#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Garmin Mailer - Single Watch, Workshop Edition (V9)

Profile (same as V8 unless noted):
    - Subject format: S1  => "Garmin FIT activities (N files) – YYYY-MM-DD"
- Submit behavior: B2 => Submit button + Enter in EMAIL field triggers send
- Multi-file: C1      => ONE email with ALL selected files attached
- Timeout: D1         => After 30s, show ❌ message + bell + enable Retry
- Labels: E1          => Label used in UI, filenames, and logs
- Checkboxes:
    • Unmount after copy (default ON, U1: just skip eject if OFF)
  • Copy only, do not send mail (default OFF)
    - NEW in V9: Auto-starts as soon as a GARMIN volume is (or becomes) mounted.
    - If already mounted when toggled ON → start immediately.
    - After success:
        - If unmounted: "Eject successful, please attach the next watch to the USB cable."
        - If not unmounted: "Copy complete. Please attach the next watch."
- NEW in V9: After a successful email, both Name and Email are cleared.

Folders:
    - ~/Documents/GarminMailer/
    - mailer.conf.json      (needed only for emailing)
    - mail-template.txt     (auto-created)
    - watch-labels.csv      (auto-created)
    - sent/YYYYMMDD/
    - devices/<device_id>/profile.json
"""

import os
import sys
import time
import json
import ssl
import smtplib
import subprocess
import re
import string
import threading
import queue
import textwrap
import xml.etree.ElementTree as ET
from datetime import datetime, date
from email.message import EmailMessage
from pathlib import Path
from typing import Optional, List, Tuple

import tkinter as tk
from tkinter import ttk, messagebox

# Optional TLS trust improvements on some Python installs
try:
    from fitparse import FitFile
    FITPARSE_OK = True
except ImportError:
    FITPARSE_OK = False

try:
    import certifi
    CERT_BUNDLE = certifi.where()
except Exception:
    CERT_BUNDLE = None

IS_WINDOWS = sys.platform.startswith("win")
IS_MAC = (sys.platform == "darwin")

# ---------------------------------------------------------------------------
# PyInstaller resource helper
# ---------------------------------------------------------------------------
def _resource_path(rel: str) -> Path:
    base = getattr(sys, "_MEIPASS", None)
    return (Path(base) / rel) if base else (Path(__file__).resolve().parent / rel)


# This will be replaced during build with actual version
BUILD_VERSION = None

def get_app_version() -> str:
    """
    Gets app version from build-time embedded version, or falls back to git tags.
    """
    default_version = "v9.2"
    
    # If version was embedded at build time, use it
    if BUILD_VERSION:
        return BUILD_VERSION
    
    # Otherwise fall back to git tags (for development)
    try:
        script_dir = Path(__file__).resolve().parent
        result = subprocess.run(
            ["git", "tag"],
            capture_output=True,
            text=True,
            check=True,
            cwd=script_dir
        )
        tags = result.stdout.strip().split('\n')
        v_tags = [t for t in tags if t and t.startswith('v')]

        if not v_tags:
            return default_version

        # Sort tags using a key that handles version numbers correctly
        def version_key(v):
            try:
                return [int(p) for p in v[1:].split('.')]
            except ValueError:
                return [0, 0, 0]
        return max(v_tags, key=version_key)
    except (subprocess.CalledProcessError, FileNotFoundError, ValueError):
        return default_version

# ---------------------------------------------------------------------------
# Paths and constants
# ---------------------------------------------------------------------------
DOCS = Path.home() / "Documents"
APP_VERSION = get_app_version()
BASE = DOCS / "GarminMailer"
SENT_ROOT = BASE / "sent"
ARCHIVE_ROOT = BASE / "archive"
LOGFILE  = BASE / "GarminMailer.log"
CONF     = BASE / "mailer.conf.json"
TEMPLATE = BASE / "mail-template.txt"
DEVICES_DIR = BASE / "devices"
LABELS_CSV = BASE / "watch-labels.csv"       # device_id,label
DEVMODE_FLAG = BASE / ".devmode"
DETECT_TIMEOUT = 30                          # seconds to wait for mount

EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")
GARMIN_NS = {"g": "http://www.garmin.com/xmlschemas/GarminDevice/v2"}

# Ensure base folders exist
for p in [BASE, SENT_ROOT, ARCHIVE_ROOT, DEVICES_DIR]:
    p.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
def log_line(msg: str) -> None:
    line = f"{datetime.now():%Y-%m-%d %H:%M:%S}  {msg}"
    print(line)
    with LOGFILE.open("a", encoding="utf-8") as f:
        f.write(line + "\n")

# ---------------------------------------------------------------------------
# PyInstaller resource helper
# ---------------------------------------------------------------------------
def _resource_path(rel: str) -> Path:
    base = getattr(sys, "_MEIPASS", None)
    return (Path(base) / rel) if base else (Path(__file__).resolve().parent / rel)

# ---------------------------------------------------------------------------
# First-run file creation
# ---------------------------------------------------------------------------
def ensure_template_exists() -> None:
    if TEMPLATE.exists():
        return
    src = _resource_path("default-mail-template.txt")
    try:
        if src.exists():
            TEMPLATE.write_text(src.read_text(encoding="utf-8"), encoding="utf-8")
        else:
            TEMPLATE.write_text("Hi {name},\n\nAttached is the latest Garmin FIT file.\n\n- Garmin Mailer\n", encoding="utf-8")
    except Exception:
        pass

def ensure_labels_csv_exists() -> None:
    if LABELS_CSV.exists():
        return
    try:
        with LABELS_CSV.open("w", encoding="utf-8", newline="") as f:
            f.write(
                    "# watch-labels.csv\n"
                    "# Format: device_id,label\n"
                    "# Add one line per watch to map Garmin device IDs to your workshop label numbers.\n"
                    "# Example:\n"
                    "# A1B2C3D4,21\n"
                    "# E7F8G9H0,7\n"
                    )
    except Exception:
        pass

# ---------------------------------------------------------------------------
# Config, template, labels
# ---------------------------------------------------------------------------
def read_config() -> dict:
    if not CONF.exists():
        raise RuntimeError(
                f"Config not found: {CONF}\n"
                "Copy 'example.mailer.conf.json' from the GarminMailer application folder into "
                f"'{CONF}' (rename it to 'mailer.conf.json') and fill in your SMTP server, port, "
                "username, and App Password."
                )
    data = json.loads(CONF.read_text(encoding="utf-8"))
    for k in ["smtp_server", "smtp_port", "username", "password"]:
        if k not in data:
            raise RuntimeError(f"Config missing key: {k}")
    return data

def read_mail_body_with_name(name: str) -> str:
    """
    Read template; replace {name} if present. If not present, content is unchanged.
    """
    ensure_template_exists()
    try:
        body = TEMPLATE.read_text(encoding="utf-8")
        if "{name}" in body:
            return body.replace("{name}", name)
        return body
    except Exception:
        fallback = "Hi {name},\n\nAttached is the latest Garmin FIT file.\n\n- Garmin Mailer"
        return fallback.replace("{name}", name) if "{name}" in fallback else fallback

def load_labels_map() -> dict:
    ensure_labels_csv_exists()
    mapping = {}
    try:
        with LABELS_CSV.open("r", encoding="utf-8") as f:
            for raw in f:
                line = raw.strip()
                if not line or line.startswith("#"):
                    continue
                if line.lower().startswith("device_id,"):
                    continue
                parts = [p.strip() for p in line.split(",")]
                if len(parts) >= 2:
                    did, lab = parts[0], parts[1]
                    if did and lab:
                        mapping[did] = lab
    except Exception:
        pass
    return mapping

# ---------------------------------------------------------------------------
# Utility
# ---------------------------------------------------------------------------
def sanitize_email_for_filename(email: str) -> str:
    email_at = email.replace("@", "-at-")
    return re.sub(r"[^A-Za-z0-9._-]", "", email_at)

def sanitize_name(name: str) -> str:
    no_spaces = re.sub(r"\s+", "", name.strip())
    return re.sub(r"[^A-Za-z0-9._-]", "", no_spaces)

# ---------------------------------------------------------------------------
# Garmin device info
# ---------------------------------------------------------------------------
def parse_garmin_device_xml(root_dir: Path) -> Tuple[Optional[str], Optional[str]]:
    dev_xml = root_dir / "GARMIN" / "GarminDevice.xml"
    try:
        if dev_xml.exists():
            tree = ET.parse(dev_xml)
            root = tree.getroot()
            device_id_el = root.find("g:Id", GARMIN_NS)
            model_el = root.find("g:Model/g:Description", GARMIN_NS)
            return (
                    device_id_el.text if device_id_el is not None else None,
                    model_el.text if model_el is not None else None,
                    )
    except Exception:
        pass
    return (None, None)

# ---------------------------------------------------------------------------
# Instant volume scan (used for auto-start in Copy-only)
# ---------------------------------------------------------------------------
def find_current_garmin_volume() -> Optional[Path]:
    """Return a mounted GARMIN volume if exactly one is present, else None."""
    if IS_MAC:
        vols_root = Path("/Volumes")
        try:
            cands = [e for e in vols_root.iterdir() if e.is_dir() and (e / "GARMIN").is_dir()]
        except Exception:
            cands = []
    else:
        cands = []
        for c in string.ascii_uppercase:
            root = Path(f"{c}:\\")
            try:
                if root.exists() and (root / "GARMIN").is_dir():
                    cands.append(root)
            except Exception:
                pass
    return cands[0] if len(cands) == 1 else None

# ---------------------------------------------------------------------------
# Single-watch detection (USB Mass Storage only)
# ---------------------------------------------------------------------------
def mac_find_single_volume(deadline: float, tick_cb) -> Optional[Path]:
    vols_root = Path("/Volumes")
    last_left = 10**9
    while time.time() < deadline:
        left = int(deadline - time.time())
        if left != last_left:
            tick_cb(max(0, left))
            last_left = left

        if not vols_root.exists():
            time.sleep(0.25)
            continue
        candidates = []
        try:
            for entry in vols_root.iterdir():
                if entry.is_dir() and (entry / "GARMIN").is_dir():
                    candidates.append(entry)
        except PermissionError:
            pass

        if len(candidates) == 1:
            return candidates[0]
        if len(candidates) > 1:
            time.sleep(0.25)
            continue

        time.sleep(0.25)
    tick_cb(0)
    return None

def win_find_single_root(deadline: float, tick_cb) -> Optional[Path]:
    last_left = 10**9
    while time.time() < deadline:
        left = int(deadline - time.time())
        if left != last_left:
            tick_cb(max(0, left))
            last_left = left

        candidates: List[Path] = []
        for c in string.ascii_uppercase:
            root = Path(f"{c}:\\")
            try:
                if root.exists() and (root / "GARMIN").is_dir():
                    candidates.append(root)
            except Exception:
                pass

        if len(candidates) == 1:
            return candidates[0]
        if len(candidates) > 1:
            time.sleep(0.25)
            continue

        time.sleep(0.25)
    tick_cb(0)
    return None

# ---------------------------------------------------------------------------
# FIT file helpers
# ---------------------------------------------------------------------------
def list_fit_files(root: Path) -> List[Path]:
    """
    Return a list of .fit files found on the mounted Garmin volume.
    Looks in:
        - <root>/GARMIN/Activity
      - <root>/Activity
    """
    results: List[Path] = []
    candidates = [
            root / "GARMIN" / "Activity",
            root / "Activity",
            ]
    seen = set()
    for folder in candidates:
        if not folder.is_dir():
            continue
        try:
            for f in folder.iterdir():
                if not f.is_file():
                    continue
                if f.name.startswith(".") or f.name.startswith("~"):
                    continue
                if f.suffix.lower() != ".fit":
                    continue
                key = str(f.resolve()) if f.exists() else str(f)
                if key in seen:
                    continue
                seen.add(key)
                results.append(f)
        except PermissionError:
            continue
        except Exception:
            continue
    return results

# ---------------------------------------------------------------------------
# Picker: multi-select with Time + Size
# ---------------------------------------------------------------------------
class FileChoiceDialog(tk.Toplevel):
    @staticmethod
    def choose(parent: tk.Tk, files: List[Path], archive_only_mode: bool, preselect_single: bool = False) -> Optional[List[Path]]:
        dlg = FileChoiceDialog(parent, files, archive_only_mode, preselect_single)
        parent.wait_window(dlg)
        return dlg.selected

    def __init__(self, parent: tk.Tk, files: List[Path], archive_only_mode: bool, preselect_single: bool = False):
        super().__init__(parent)
        if archive_only_mode:
            self.title("Choose recent activities to archive")
            dialog_text = "Select one or more recent activities to archive (Cmd/Ctrl-click or Shift-click):"
        else:
            self.title("Choose today's activities")
            dialog_text = "Multiple activities found for today. Select which file(s) to email (Cmd/Ctrl-click or Shift-click):"

        self.resizable(False, False)
        self.selected: Optional[List[Path]] = None

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(
                frm, text=dialog_text
        ).pack(anchor="w", pady=(0, 8))

        if archive_only_mode:
            columns = ("date", "time", "size")
        else:
            columns = ("time", "size")

        tree = ttk.Treeview(
                frm,
                columns=columns,
                show="headings",
                height=min(10, len(files)),
                selectmode="extended"
        )

        col_width = 130
        if archive_only_mode:
            tree.heading("date", text="Date")
            tree.column("date", width=col_width, anchor="center")

        tree.heading("time", text="Time")
        tree.heading("size", text="Size") # No change needed for size
        tree.column("time", width=col_width, anchor="center")
        tree.column("size", width=col_width, anchor="center")

        files_sorted = sorted(files, key=lambda p: p.stat().st_mtime)
        self._iid_to_path: dict[str, Path] = {}

        def fmt_size(n: int) -> str:
            if n >= 1024 * 1024:
                return f"{n / (1024*1024):.1f} MB"
            if n >= 1024:
                return f"{n / 1024:.0f} KB"
            return f"{n} B"

        for f in files_sorted:
            mod_time = datetime.fromtimestamp(f.stat().st_mtime)
            date_str = mod_time.strftime("%d-%b-%Y")
            time_str = mod_time.strftime("%H:%M:%S")
            size = fmt_size(f.stat().st_size)
            if archive_only_mode:
                iid = tree.insert("", "end", values=(date_str, time_str, size))
            else:
                iid = tree.insert("", "end", values=(time_str, size))

            # Pre-select the item if preselect_single is True
            if preselect_single:
                tree.selection_set(iid)
            self._iid_to_path[iid] = f

        tree.pack(fill="both", expand=True)

        def finalize_selection():
            sels = tree.selection()
            if not sels:
                messagebox.showinfo("Garmin Mailer", "Please select at least one activity.")
                return
            self.selected = [self._iid_to_path[i] for i in sels]
            self.destroy()

        tree.bind("<Double-1>", lambda _e: finalize_selection())

        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side="right", padx=(6, 0))
        ttk.Button(btns, text="Select", command=finalize_selection).pack(side="right")

        self.grab_set()
        self.transient(parent)
        self.wait_visibility()
        self.focus()
        self.lift()

# ---------------------------------------------------------------------------
# Email (supports multiple attachments, S1 subject style)
# ---------------------------------------------------------------------------
def send_email_gmail(conf: dict, to_addr: str, attachments: List[Path], body_text: str) -> None:
    """
    Send one email to to_addr with one or more attachments.
    Subject includes count if multiple attachments (S1).
    """
    count = len(attachments)
    if count <= 1:
        subject = f"Garmin FIT {datetime.now():%Y-%m-%d}"
    else:
        subject = f"Garmin FIT activities ({count} files) – {datetime.now():%Y-%m-%d}"

    msg = EmailMessage()
    msg["From"] = conf.get("from_address", conf["username"])
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.set_content(body_text if body_text else "Attached is the latest Garmin FIT file(s).")

    for file_path in attachments:
        data = file_path.read_bytes()
        msg.add_attachment(
                data,
                maintype="application",
                subtype="octet-stream",
                filename=file_path.name
                )

    smtp_server = conf["smtp_server"]
    smtp_port = int(conf["smtp_port"])
    username = conf["username"]
    password = conf["password"]

    ctx = ssl.create_default_context(cafile=CERT_BUNDLE) if CERT_BUNDLE else ssl.create_default_context()

    if smtp_port == 587:
        with smtplib.SMTP(smtp_server, smtp_port) as smtp:
            smtp.starttls(context=ctx)
            smtp.login(username, password)
            smtp.send_message(msg)
    elif smtp_port == 465:
        with smtplib.SMTP_SSL(smtp_server, smtp_port, context=ctx) as smtp:
            smtp.login(username, password)
            smtp.send_message(msg)
    else:
        raise ValueError(f"Unsupported SMTP port: {smtp_port}. Only 465 (SSL) and 587 (STARTTLS) are supported.")

# ---------------------------------------------------------------------------
# Background worker thread (no Tk calls inside; uses queues)
# ---------------------------------------------------------------------------
class Worker(threading.Thread):
    """
    Steps:
        1) Ensure first-run files (template, labels).
      2) (If emailing) Load config.
      3) Wait up to DETECT_TIMEOUT for ONE Garmin device.
      4) Map device_id to label (if available).
      5) Decide files for today; if >1, ask UI to pick (multi-select).
      6) Copy selected files to sent/YYYYMMDD/ with new naming convention.
      7) Eject the watch once after copy is done (if unmount_after_copy=True).
      8) If not copy_only: Send ONE email with all selected files attached.
      9) Log, update profile, and notify UI.
    """
    def __init__(
            self,
            ui_queue: "queue.Queue[str]",
            name_val: str,
            email_val: str,
            parent: tk.Tk,
            cancel_event: threading.Event,
            unmount_after_copy: bool,
            archive_only: bool,
            ):
        super().__init__(daemon=True)
        self.ui_queue = ui_queue
        self.name_val = name_val
        self.email_val = email_val
        self.parent = parent
        self.cancel_event = cancel_event
        self.start_ts = time.time()
        self.labels_map = load_labels_map()
        self.pick_reply_queue: "queue.Queue[List[str] | None]" = queue.Queue()
        self.unmount_after_copy = unmount_after_copy
        self.archive_only = archive_only
        self.saved_paths: List[Path] = []

    def post(self, msg: str) -> None:
        self.ui_queue.put(msg)

    def _check_cancel(self) -> bool:
        if self.cancel_event.is_set():
            self.post("ERROR|Cancelled by user.")
            return True
        return False

    def run(self) -> None:
        ensure_template_exists()
        ensure_labels_csv_exists()

        name = self.name_val.strip()
        email = self.email_val.strip()
        name_sane = sanitize_name(name) if name else ""
        email_sane = sanitize_email_for_filename(email) if email else ""

        # Load config only if emailing
        conf = None
        if not self.archive_only:
            try:
                conf = read_config()
            except Exception as e:
                self.post(f"ERROR|Config: {e}")
                return

        # Detect single watch (up to 30 seconds)
        if self._check_cancel():
            return
        self.post("STEP|Detecting Garmin watch...|SPIN_ON")
        self.post(f"COUNT|{DETECT_TIMEOUT}")

        deadline = time.time() + DETECT_TIMEOUT
        def tick(n: int) -> None:
            self.post(f"COUNT|{n}")

        root = mac_find_single_volume(deadline, tick) if IS_MAC else win_find_single_root(deadline, tick)

        if self._check_cancel():
            return

        if not root:
            self.post("STEP|Detection timed out.|SPIN_OFF")
            self.post("COUNT|HIDE")
            self.post("ERROR|❌ No Garmin watch detected. Connect the watch and press Retry.")
            return

        device_id, model = parse_garmin_device_xml(root)
        label = self.labels_map.get(device_id or "", None)
        human_name = f"Garmin watch {label}" if label else "Garmin watch"
        self.post("STEP|" + human_name + " found|SPIN_OFF")
        self.post("COUNT|HIDE")

        # Per-device dir
        dev_id_for_fs = device_id or "unknown"
        (DEVICES_DIR / dev_id_for_fs).mkdir(parents=True, exist_ok=True)

        # List and decide files
        files = list_fit_files(root)
        if not files:
            if self.unmount_after_copy:
                mac_eject(root) if IS_MAC else win_eject_drive(root)
            self.post("ERROR|No .fit files found on the watch.")
            return

        # Selection logic
        if self.archive_only:
            # Show the 5 most recent FIT files (by mtime), let user multi-select
            files_sorted = sorted(files, key=lambda p: p.stat().st_mtime, reverse=True)
            recent = files_sorted[:5] if len(files_sorted) > 5 else files_sorted
            if not recent:
                if self.unmount_after_copy:
                    mac_eject(root) if IS_MAC else win_eject_drive(root)
                self.post("ERROR|No .fit files found on the watch.")
                return
            self.post("ASK_PICK|" + json.dumps([str(p) for p in recent]))
            selected_paths = self._receive_pick_selection()
            if not selected_paths:
                self.post("ERROR|No file selected.")
                return
        else:
            # Email mode: only consider files with today's modification date
            today_date = date.today()
            todays_files = [f for f in files if datetime.fromtimestamp(f.stat().st_mtime).date() == today_date]

            if not todays_files:
                self.post("ERROR|No activity files from today were found on the watch.")
                return

            # Always show the file picker, pre-selecting the single file if it exists
            if len(todays_files) == 1:
                preselect_single = True
            else:
                preselect_single = False
            self.post("ASK_PICK|" + json.dumps([str(p) for p in todays_files]) + f"|PRESELECT:{preselect_single}")
            selected_paths = self._receive_pick_selection()
            if not selected_paths:
                self.post("ERROR|No file was selected to email.")
                return

        selected = [Path(s) for s in selected_paths]

        for src in selected:
            # Determine the correct date string and save directory for each file.
            # In email mode, this is always today's date.
            # In archive mode, it's the activity date from the file.
            activity_date_str = datetime.now().strftime("%Y%m%d")

            if not self.archive_only:
                # Email mode: use today's date for directory and filename
                save_root = SENT_ROOT
                if label:
                    newname = f"{activity_date_str}_{label}_{email_sane}_{src.name}"
                else:
                    newname = f"{activity_date_str}_{email_sane}_{src.name}"
            else:
                # Archive mode: determine activity date from file for directory and filename
                save_root = ARCHIVE_ROOT
                # 1. Fallback to file modification date
                activity_date_str = datetime.fromtimestamp(src.stat().st_mtime).strftime("%Y%m%d")
                # 2. Try to get the actual recording date from the FIT file
                if FITPARSE_OK:
                    try:
                        fitfile = FitFile(src)
                        time_created = None
                        for record in fitfile.get_messages('file_id'):
                            if record.get_value('time_created'):
                                time_created = record.get_value('time_created')
                                break
                        if time_created:
                            activity_date_str = time_created.strftime("%Y%m%d")
                    except Exception:
                        pass # Fallback to file modification date on any parsing error
                newname = f"{activity_date_str}_{label}_{src.name}" if label else f"{activity_date_str}_{src.name}"

            # Create the dated directory and define the final destination path
            save_dir = save_root / activity_date_str
            save_dir.mkdir(parents=True, exist_ok=True)
            dest = save_dir / newname
            try:
                dest.write_bytes(src.read_bytes())
            except PermissionError:
                tip = "Grant Full Disk Access to Terminal or Python in macOS System Settings."
                self.post(f"ERROR|Permission denied reading FIT. {tip}")
                return
            except Exception as e:
                self.post(f"ERROR|Copy failed: {e}")
                return
            self.saved_paths.append(dest)

        # Eject after copy
        eject_result = None
        # In Archive mode, always eject. In Email mode, respect the checkbox.
        should_eject = self.archive_only or self.unmount_after_copy
        if should_eject:
            if IS_MAC:
                eject_result = mac_eject(root)
            else:
                eject_result = win_eject_drive(root)

        # Email or archive-only completion
        if self.archive_only:
            # Log and profile
            prof = {
                    "device_id": device_id,
                    "model": model,
                    "last_archived_files": [p.name for p in self.saved_paths],
                    "last_action": "archive_only",
                    "last_time": datetime.now().isoformat(timespec="seconds"),
                    }
            try:
                (DEVICES_DIR / dev_id_for_fs / "profile.json").write_text(json.dumps(prof, indent=2), encoding="utf-8")
            except Exception:
                pass

            elapsed = int(time.time() - self.start_ts)
            for src, dest in zip(selected, self.saved_paths):
                log_line(
                        f"ARCHIVED  label={(label or '')}  file={dest}  src={src.name}  "
                        f"device_id={(device_id or '')}  model={model}  duration={elapsed}s  mode=ARCHIVE_ONLY"
                        )

            if eject_result:
                text = "Eject successful, please attach the next watch to the USB cable."
            else:
                text = "Archive complete. Please eject and attach the next watch."
            self.post("DONE|" + text + "|100|" + str(save_dir) + "|MODE:ARCHIVE_ONLY")
            return

        # Send single email with all attachments
        if self._check_cancel():
            return
        body_text = read_mail_body_with_name(name)
        self.post(f"STEP|Sending email... ({len(self.saved_paths)} attachment(s))|90")
        try:
            send_email_gmail(conf, email, self.saved_paths, body_text)  # type: ignore[arg-type]
        except smtplib.SMTPAuthenticationError:
            self.post("ERROR|AUTH: Brevo rejected login. Check your username/password in mailer.conf.json.")
            return
        except ssl.SSLError as e:
            self.post(f"ERROR|SSL: {e}. Tip: install certifi.")
            return
        except Exception as e:
            self.post(f"ERROR|Send failed: {e}")
            return

        # Update profile and log
        prof = {
                "device_id": device_id,
                "model": model,
                "last_sent_files": [p.name for p in self.saved_paths],
                "last_sent_time": datetime.now().isoformat(timespec="seconds"),
                }
        try:
            (DEVICES_DIR / dev_id_for_fs / "profile.json").write_text(json.dumps(prof, indent=2), encoding="utf-8")
        except Exception:
            pass

        elapsed = int(time.time() - self.start_ts)
        for src, dest in zip(selected, self.saved_paths):
            log_line(
                    f"SENT  label={(label or '')}  name={name}  email={email}  file={dest}  "
                    f"src={src.name}  device_id={(device_id or '')}  model={model}  duration={elapsed}s  mode=EMAIL"
                    )

        self.post("DONE|Email sent.|100|" + str(save_dir) + "|MODE:EMAIL")

    def _receive_pick_selection(self) -> Optional[List[str]]:
        try:
            selected_paths = self.pick_reply_queue.get(timeout=180)
        except queue.Empty:
            return None
        if selected_paths is None:
            return None
        if isinstance(selected_paths, (str, Path)):
            return [str(selected_paths)]
        elif isinstance(selected_paths, (list, tuple, set)):
            return [str(p) for p in selected_paths]
        else:
            return None

# ---------------------------------------------------------------------------
# Platform-specific eject helpers
# ---------------------------------------------------------------------------
def mac_eject(volume: Path) -> bool:
    try:
        res = subprocess.run(
                ["diskutil", "unmount", str(volume)],
                check=False, stdout=subprocess.PIPE, stderr=subprocess.PIPE
                )
        return res.returncode == 0
    except Exception:
        return False

def win_eject_drive(root: Path) -> bool:
    # Placeholder on Windows; return True for flow purposes
    return True

# ---------------------------------------------------------------------------
# Main Tk app
# ---------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Garmin Mailer")
        self.geometry("680x510")
        self.resizable(False, False)

        # Style
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TLabel", font=("Arial", 13))
        style.configure("Small.TLabel", font=("Arial", 11))
        style.configure("TButton", font=("Arial", 13))
        style.configure("TEntry", padding=4)

        # State
        self.queue: "queue.Queue[str]" = queue.Queue()
        self.worker: Optional[Worker] = None
        self.running = False
        self.cancel_event = threading.Event()
        self.current_submission_key: Optional[str] = None
        self._last_sent_dir: Optional[Path] = None

        # Layout
        frm = ttk.Frame(self, padding=16)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)

        # Name row
        ttk.Label(frm, text="Name").grid(row=0, column=0, sticky="w")
        self.name_var = tk.StringVar()
        self.name_entry = ttk.Entry(frm, textvariable=self.name_var, width=48)
        self.name_entry.grid(row=0, column=1, sticky="ew", padx=(8, 0))

        # Email row + Submit button inline
        ttk.Label(frm, text="Recipient email").grid(row=1, column=0, sticky="w")
        email_row = ttk.Frame(frm)
        email_row.grid(row=1, column=1, sticky="ew", padx=(8, 0))
        email_row.columnconfigure(0, weight=1)
        self.email_var = tk.StringVar()
        self.email_entry = ttk.Entry(email_row, textvariable=self.email_var, width=48) # type: ignore
        self.email_entry.grid(row=0, column=0, sticky="ew")
        self.submit_btn = ttk.Button(email_row, text="Ready", command=self._submit, state="disabled")
        self.submit_btn.grid(row=0, column=1, sticky="e", padx=(8, 0))

        # Checkboxes
        opts_row = ttk.Frame(frm)
        opts_row.grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 0))
        self.unmount_var = tk.BooleanVar(value=True)      # default checked
        self.archive_only_var = tk.BooleanVar(value=False)    # default unchecked
        self.archive_only_cb = ttk.Checkbutton(opts_row, text="Archive only, do not send mail", variable=self.archive_only_var, command=self._on_archive_only_toggle)
        self.archive_only_cb.pack(side="left")

        # Unmount checkbox is only visible in dev mode
        self.unmount_cb = ttk.Checkbutton(opts_row, text="Unmount after copy", variable=self.unmount_var)
        if DEVMODE_FLAG.exists():
            self.unmount_cb.pack(side="left", padx=(0, 16), before=self.archive_only_cb)

        self.hint = ttk.Label(frm, text="Attach your Garmin via USB (Mass Storage).", style="Small.TLabel")
        self.hint.grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 10))

        self.pb = ttk.Progressbar(frm, orient="horizontal", mode="determinate", maximum=100)
        self.pb.grid(row=4, column=0, columnspan=2, sticky="ew")

        self.status_var = tk.StringVar(value="Waiting for name and email")
        ttk.Label(frm, textvariable=self.status_var).grid(row=5, column=0, columnspan=2, sticky="w", pady=(10, 0))

        # Open folder + Help buttons
        self.open_folder_btn = ttk.Button(frm, text="Open 'sent' folder", command=self._open_folder, state="disabled")
        self.open_folder_btn.grid(row=6, column=0, sticky="w", pady=(8, 0))
        self.help_btn = ttk.Button(frm, text="Help", command=self._show_help)
        self.help_btn.grid(row=6, column=1, sticky="e", pady=(8, 0))

        # Timer and countdown
        self.timer_var = tk.StringVar(value="Timer: 0s")
        ttk.Label(frm, textvariable=self.timer_var, style="Small.TLabel").grid(row=7, column=0, sticky="w", pady=(8, 0))
        self.timer_seconds = 0
        self.timer_running = False

        self.detect_countdown_var = tk.StringVar(value="")
        self.detect_countdown_lbl = ttk.Label(frm, textvariable=self.detect_countdown_var, style="Small.TLabel")
        self.detect_countdown_lbl.grid(row=7, column=1, sticky="e", pady=(8, 0))
        self._hide_detect_countdown()

        # Action row (Cancel + Retry)
        btn_row = ttk.Frame(frm)
        btn_row.grid(row=8, column=0, columnspan=2, sticky="ew", pady=(14, 0))
        self.cancel_btn = ttk.Button(btn_row, text="Cancel", command=self._cancel, state="disabled")
        self.cancel_btn.pack(side="left")
        self.retry_btn = ttk.Button(btn_row, text="Retry", command=self._retry, state="disabled")
        self.retry_btn.pack(side="right")

        # Version label
        is_frozen = getattr(sys, "_MEIPASS", None) is not None
        version_str = f"Version: {APP_VERSION}"
        if not is_frozen:
            version_str += " (local)"
        ttk.Label(frm, text=version_str, style="Small.TLabel", foreground="gray").grid(row=9, column=0, sticky="w", pady=(10, 0))

        # Bindings
        self.after(100, self.name_entry.focus_set)
        self.email_entry.bind("<Return>", lambda _e: self._submit())  # Enter in Email triggers submit (B2)
        self.name_var.trace_add("write", self._validate_form)
        self.email_var.trace_add("write", self._validate_form)
        self.after(100, self._drain_queue)

        # Initial UI reflect + start watcher for auto-start in archive-only mode
        self._reflect_archive_only_state()
        self.after(1000, self._watch_mount_if_archive_only)

    # --- UI helpers ---------------------------------------------------------
    def _validate_form(self, *_):
        if self.archive_only_var.get():
            self.submit_btn.configure(state="normal" if not self.running else "disabled")
            if not self.running:
                self.status_var.set("Archive-only mode: attach a watch to begin")
            return

        name_ok = bool(self.name_var.get().strip())
        email_ok = EMAIL_RE.match(self.email_var.get().strip()) is not None
        self.submit_btn.configure(state="normal" if (name_ok and email_ok and not self.running) else "disabled")
        if not self.running:
            self.status_var.set("Waiting for name and email" if not (name_ok and email_ok) else "Ready to submit")

    def _on_archive_only_toggle(self):
        self._reflect_archive_only_state()
        self._validate_form()

        # NEW: If toggled ON and a GARMIN volume is already mounted → start immediately.
        if self.archive_only_var.get() and not self.running:
            vol = find_current_garmin_volume()
            if vol is not None:
                # Reset timer for this auto-run
                self.current_submission_key = None
                self._start_flow(name="", email="", reset_timer=True,
                                 unmount_after_copy=self.unmount_var.get(),
                                 archive_only=True)
            else:
                self.status_var.set("Archive-only mode: waiting for GARMIN volume...")

    def _reflect_archive_only_state(self):
        archive_only = self.archive_only_var.get()
        # Enable/disable inputs
        state = "disabled" if archive_only else "normal"
        self.name_entry.configure(state=state)
        self.email_entry.configure(state=state)
        self.open_folder_btn.configure(text="Open 'archive' folder" if archive_only else "Open 'sent' folder")
        if archive_only:
            self.status_var.set("Archive-only mode: attach a watch to begin")
        else:
            self.status_var.set("Waiting for name and email")

    def _watch_mount_if_archive_only(self):
        """
        Poll every 1s: if archive-only mode is ON, not running, and exactly one GARMIN volume is mounted → auto-start.
        """
        try:
            if self.archive_only_var.get() and not self.running:
                vol = find_current_garmin_volume()
                if vol is not None:
                    # Reset timer for this auto-run
                    self.current_submission_key = None
                    self._start_flow(name="", email="", reset_timer=True,
                                     unmount_after_copy=self.unmount_var.get(),
                                     archive_only=True)
                else:
                    # keep a friendly status while waiting
                    if self.status_var.get().strip() == "" or "Archive-only" not in self.status_var.get():
                        self.status_var.set("Archive-only mode: waiting for GARMIN volume...")
        finally:
            self.after(1000, self._watch_mount_if_archive_only)

    def _show_help(self):
        help_text = textwrap.dedent(f"""
            Garmin Mailer copies FIT activities from a connected Garmin watch and either emails them to the email address specified or archives them for later reference.

            Email vs archive-only:
              - Email (default): enter Name and Recipient email, then click Ready. The app emails the selected .FIT files and saves today's activities into {SENT_ROOT}/YYYYMMDD.
              - Archive only: check "Archive only, do not send mail". Name/email are disabled and the app auto-starts when exactly one GARMIN volume is mounted. Files are renamed by activity date and stored under {ARCHIVE_ROOT}/YYYYMMDD with no email sent.

            Storage under {BASE}:
              - sent/YYYYMMDD/: copies of files that were emailed
              - archive/YYYYMMDD/: archive-only copies (organized by activity date)
              - devices/<device_id>/profile.json: remembers the label/model for each watch
              - mail-template.txt: edit the outgoing email body
              - watch-labels.csv: map Garmin device_ids to workshop labels (CSV with "device_id,label" per line)
              - GarminMailer.log: timestamped record of actions and errors

            Configuring emailing:
              1. Copy example.mailer.conf.json (from the GarminMailer folder) into {CONF}.
              2. Fill smtp_server, smtp_port, username, and password (use an app password).
              3. Restart Garmin Mailer so it reads the settings. Archive-only mode works without this file, but emailing requires it.

            Updating watch labels:
              1. Open {LABELS_CSV}.
              2. Each line is "device_id,label" (for example A1B2C3D4,21).
              3. Save the file; the next time that watch connects the label populates automatically.
            """).strip()
        messagebox.showinfo("Garmin Mailer Help", help_text, parent=self)

    # Timer helpers
    def _start_timer(self, reset: bool) -> None:
        if reset:
            self.timer_seconds = 0
        if not self.timer_running:
            self.timer_running = True
            self._tick_timer()

    def _tick_timer(self) -> None:
        if not self.timer_running:
            return
        self.timer_var.set(f"Timer: {self.timer_seconds}s")
        self.timer_seconds += 1
        self.after(1000, self._tick_timer)

    def _stop_timer(self) -> None:
        self.timer_running = False

    # Countdown helpers
    def _show_detect_countdown(self, seconds_left: Optional[int] = None) -> None:
        self.detect_countdown_lbl.grid()
        if seconds_left is not None:
            self.detect_countdown_var.set(f"Time left: {seconds_left}s")

    def _hide_detect_countdown(self) -> None:
        self.detect_countdown_var.set("")
        self.detect_countdown_lbl.grid_remove()

    # Submit flow
    def _submit(self):
        if self.running:
            return

        archive_only = self.archive_only_var.get()
        name = self.name_var.get().strip()
        email = self.email_var.get().strip()

        if not archive_only:
            if not name:
                messagebox.showinfo("Garmin Mailer", "Please enter a name.")
                return
            if not EMAIL_RE.match(email):
                messagebox.showinfo("Garmin Mailer", "Please enter a valid email address.")
                return

        # Timer reset per watch
        submission_key = f"{'ARCHIVE_ONLY' if archive_only else name}|{'ARCHIVE_ONLY' if archive_only else email}"
        reset_timer = (self.current_submission_key is None) or (submission_key != self.current_submission_key)
        self.current_submission_key = submission_key

        self._start_flow(name, email, reset_timer, self.unmount_var.get(), archive_only)

    def _set_pb_indeterminate(self, on: bool) -> None:
        try:
            if on:
                self.pb.config(mode="indeterminate")
                self.pb.start(12)
            else:
                self.pb.stop()
                self.pb.config(mode="determinate")
        except Exception:
            pass

    def _start_flow(self, name: str, email: str, reset_timer: bool, unmount_after_copy: bool, archive_only: bool) -> None:
        self.running = True
        self.retry_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")
        self.submit_btn.configure(state="disabled")
        self.cancel_event.clear()
        self.open_folder_btn.configure(state="disabled")
        self._set_status("Detecting Garmin watch...", None)
        self._show_detect_countdown(DETECT_TIMEOUT)
        self._set_pb_indeterminate(True)
        self._start_timer(reset=reset_timer)
        self.worker = Worker(
                self.queue, name, email, self, self.cancel_event,
                unmount_after_copy=unmount_after_copy,
                archive_only=archive_only
                )
        self.worker.start()

    def _cancel(self) -> None:
        if not self.running:
            return
        self.cancel_event.set()
        self._set_status("Cancelling...", None)
        self._set_pb_indeterminate(True)

    def _retry(self) -> None:
        if self.running:
            return
        if not self.timer_running:
            self._start_timer(reset=False)
        self.pb["value"] = 0
        self.status_var.set("Detecting Garmin watch...")
        self._show_detect_countdown(DETECT_TIMEOUT)
        name = self.name_var.get().strip()
        email = self.email_var.get().strip()
        archive_only = self.archive_only_var.get()
        if archive_only or (name and EMAIL_RE.match(email)):
            self._set_pb_indeterminate(True)
            self.running = True
            self.retry_btn.configure(state="disabled")
            self.cancel_btn.configure(state="normal")
            self.cancel_event.clear()
            self.worker = Worker(
                    self.queue, name, email, self, self.cancel_event,
                    unmount_after_copy=self.unmount_var.get(),
                    archive_only=archive_only
                    )
            self.worker.start()
        else:
            messagebox.showinfo("Garmin Mailer", "Enter a valid name and email first (or enable Archive-only).")

    def _drain_queue(self) -> None:
        try:
            while True:
                msg = self.queue.get_nowait()
                parts = msg.split("|", 1)
                kind = parts[0]

                if kind == "COUNT":
                    val = parts[1]
                    if val == "HIDE":
                        self._hide_detect_countdown()
                    else:
                        try:
                            secs = int(val)
                            self._show_detect_countdown(secs)
                        except Exception:
                            pass

                elif kind == "STEP":
                    # STEP|text|progress_or_spin
                    sub = parts[1].split("|")
                    text, prog = sub[0], sub[1] if len(sub) > 1 else None
                    if prog == "SPIN_ON":
                        self._set_status(text, None)
                        self._set_pb_indeterminate(True)
                    elif prog == "SPIN_OFF":
                        self._set_pb_indeterminate(False)
                        self._set_status(text, None)
                    else:
                        try:
                            val = int(prog) if prog is not None else None
                        except Exception:
                            val = None
                        self._set_pb_indeterminate(False)
                        self._set_status(text, val)

                elif kind == "DONE":
                    # DONE|message|progress|/path/to/sent_dir|MODE:XXXX
                    sub = parts[1].split("|")
                    text = sub[0]
                    prog = int(sub[1]) if len(sub) > 1 else 100
                    self._last_sent_dir = Path(sub[2]) if len(sub) > 2 else None
                    mode_tag = sub[3] if len(sub) > 3 else "MODE:EMAIL"
                    mode = "EMAIL" if mode_tag.endswith("EMAIL") else "ARCHIVE_ONLY"
                    self._set_pb_indeterminate(False)
                    self._hide_detect_countdown()
                    self._set_status(text, prog)
                    self._on_success(mode, message=text)

                elif kind == "ERROR":
                    text = parts[1]
                    self._set_pb_indeterminate(False)
                    self._hide_detect_countdown()
                    self._set_status(text, None)  # may include emoji
                    self._on_error(text)

                elif kind == "ASK_PICK":
                    # ASK_PICK|[ "path1", "path2", ... ]
                    sub_parts = parts[1].split("|PRESELECT:")
                    raw_paths = sub_parts[0]
                    preselect_single = False
                    if len(sub_parts) > 1:
                        preselect_single = sub_parts[1].lower() == 'true'

                    try:
                        paths = [Path(s) for s in json.loads(raw_paths)]
                    except Exception:
                        paths = []

                    # Pass copy_only status to the dialog
                    archive_only_mode = False
                    if self.worker:
                        archive_only_mode = self.worker.archive_only

                    chosen = FileChoiceDialog.choose(self, paths, archive_only_mode, preselect_single) if paths else None

                    if self.worker and hasattr(self.worker, "pick_reply_queue"):
                        if chosen:
                            self.worker.pick_reply_queue.put([str(p) for p in chosen])
                            self._set_status(f"Selected {len(chosen)} file(s).", None)
                        else:
                            self.worker.pick_reply_queue.put(None)
        except queue.Empty:
            pass
        finally:
            self.after(100, self._drain_queue)

    def _set_status(self, text: str, prog: Optional[int]) -> None:
        self.status_var.set(text)
        if prog is not None:
            self.pb["value"] = max(0, min(100, prog))
        self.update_idletasks()

    def _on_success(self, mode: str, message: str) -> None:
        self._stop_timer()
        self.running = False
        self.cancel_btn.configure(state="disabled")
        self.open_folder_btn.configure(state="normal")
        self.submit_btn.configure(state="normal")
        try:
            self.bell()
        except Exception:
            pass

        if mode == "EMAIL":
            if IS_MAC:
                # Show a native macOS notification
                try:
                    subprocess.run(
                            ["osascript", "-e", 'display notification "Email sent." with title "Garmin Mailer"'],
                            check=False
                            )
                except Exception:
                    pass

            # Show a popup message
            email = self.email_var.get()
            attachment_count = len(self.worker.saved_paths) if self.worker is not None and hasattr(self.worker, "saved_paths") else 0
            message = f"Email successfully sent to {email} with {attachment_count} attachment{'s' if attachment_count != 1 else ''}."
            messagebox.showinfo("Garmin Mailer", message)
        else:
            self.status_var.set(message)
            self.pb["value"] = 100
            self.open_folder_btn.configure(state="normal")
            messagebox.showinfo("Garmin Mailer", message)
            # Ready for next cycle (fields remain disabled in copy-only)
            self._validate_form()
            self.submit_btn.configure(state="normal")

    def _on_error(self, text: str) -> None:
        self._stop_timer()
        self.running = False
        self.cancel_btn.configure(state="disabled")
        try:
            self.bell()  # audible ding on error/timeouts
        except Exception:
            pass
        messagebox.showerror("Garmin Mailer Error", text)
        self.retry_btn.configure(state="normal")
        self._validate_form()

    def _open_folder(self) -> None:
        archive_only = self.archive_only_var.get()
        if archive_only:
            target = self._last_sent_dir or ARCHIVE_ROOT
        else:
            target = self._last_sent_dir or SENT_ROOT
        try:
            if IS_MAC:
                subprocess.run(["open", str(target)], check=False)
            else:
                os.startfile(str(target))  # type: ignore[attr-defined]
        except Exception:
            messagebox.showinfo("Garmin Mailer", f"Open this folder manually:\n{target}")

# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    try:
        app = App()
        app.mainloop()
    except KeyboardInterrupt:
        pass
