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
    - Devices/<device_id>/profile.json
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
import xml.etree.ElementTree as ET
from datetime import datetime, date
from email.message import EmailMessage
from pathlib import Path
from typing import Optional, List, Tuple

import tkinter as tk
from tkinter import ttk, messagebox

# Optional TLS trust improvements on some Python installs
try:
    import certifi
    CERT_BUNDLE = certifi.where()
except Exception:
    CERT_BUNDLE = None

IS_WINDOWS = sys.platform.startswith("win")
IS_MAC = (sys.platform == "darwin")

# ---------------------------------------------------------------------------
# Paths and constants
# ---------------------------------------------------------------------------
DOCS = Path.home() / "Documents"
BASE = DOCS / "GarminMailer"
SENT_ROOT = BASE / "sent"
LOGFILE  = BASE / "GarminMailSend.log"
CONF     = BASE / "mailer.conf.json"
TEMPLATE = BASE / "mail-template.txt"
DEVICES_DIR = BASE / "Devices"
LABELS_CSV = BASE / "watch-labels.csv"       # device_id,label
DETECT_TIMEOUT = 30                          # seconds to wait for mount

EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")
GARMIN_NS = {"g": "http://www.garmin.com/xmlschemas/GarminDevice/v2"}

# Ensure base folders exist
for p in [BASE, SENT_ROOT, DEVICES_DIR]:
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
                "Create JSON with the Gmail SMTP server, port 465, your username and App Password."
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
def sanitize_localpart(email: str) -> str:
    lp = email.split("@")[0]
    return re.sub(r"[^A-Za-z0-9._-]", "", lp)

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
    def choose(parent: tk.Tk, files_today: List[Path]) -> Optional[List[Path]]:
        dlg = FileChoiceDialog(parent, files_today)
        parent.wait_window(dlg)
        return dlg.selected

    def __init__(self, parent: tk.Tk, files: List[Path]):
        super().__init__(parent)
        self.title("Choose today's activities")
        self.resizable(False, False)
        self.selected: Optional[List[Path]] = None

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(
                frm,
                text="Multiple activities found for today. Select one or more (Cmd/Ctrl-click or Shift-click):"
                ).pack(anchor="w", pady=(0, 8))

        columns = ("time", "size")
        tree = ttk.Treeview(
                frm,
                columns=columns,
                show="headings",
                height=min(10, len(files)),
                selectmode="extended"
                )
        tree.heading("time", text="Time (24-hour)")
        tree.heading("size", text="Size")
        tree.column("time", width=180, anchor="w")
        tree.column("size", width=100, anchor="e")

        files_sorted = sorted(files, key=lambda p: p.stat().st_mtime)
        self._iid_to_path: dict[str, Path] = {}

        def fmt_size(n: int) -> str:
            if n >= 1024 * 1024:
                return f"{n / (1024*1024):.1f} MB"
            if n >= 1024:
                return f"{n / 1024:.0f} KB"
            return f"{n} B"

        for f in files_sorted:
            ts = datetime.fromtimestamp(f.stat().st_mtime).strftime("%H:%M")
            size = fmt_size(f.stat().st_size)
            iid = tree.insert("", "end", values=(ts, size))
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

    ctx = ssl.create_default_context(cafile=CERT_BUNDLE) if CERT_BUNDLE else ssl.create_default_context()
    with smtplib.SMTP_SSL(conf["smtp_server"], int(conf["smtp_port"]), context=ctx) as smtp:
        smtp.login(conf["username"], conf["password"])
        smtp.send_message(msg)

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
            copy_only: bool,
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
        self.copy_only = copy_only

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
        local = sanitize_localpart(email) if email else ""
        name_sane = sanitize_name(name) if name else ""

        # Load config only if emailing
        conf = None
        if not self.copy_only:
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
        if self.copy_only:
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
                # Email mode: prefer today's files; if none, fall back to latest single file
                today_date = date.today()
                todays = [f for f in files if datetime.fromtimestamp(f.stat().st_mtime).date() == today_date]
                if len(todays) == 0:
                    selected_paths = [str(max(files, key=lambda p: p.stat().st_mtime))]
                elif len(todays) == 1:
                    selected_paths = [str(todays[0])]
                else:
                    self.post("ASK_PICK|" + json.dumps([str(p) for p in todays]))
                    selected_paths = self._receive_pick_selection()
                    if selected_paths is None:
                        self.post("ERROR|No file selected.")
                        return

        selected = [Path(s) for s in selected_paths]

        # Prepare sent dir and copy with proper naming
        day = datetime.now().strftime("%Y%m%d")
        sent_dir = SENT_ROOT / day
        sent_dir.mkdir(parents=True, exist_ok=True)

        saved_paths: List[Path] = []
        for src in selected:
            if not self.copy_only:
                if label:
                    newname = f"{day}_{label}_{name_sane}_{local}_{src.name}"
                else:
                    newname = f"{day}_{name_sane}_{local}_{src.name}"
            else:
                newname = f"{day}_{label}_{src.name}" if label else f"{day}_{src.name}"
            dest = sent_dir / newname
            try:
                dest.write_bytes(src.read_bytes())
            except PermissionError:
                tip = "Grant Full Disk Access to Terminal or Python in macOS System Settings."
                self.post(f"ERROR|Permission denied reading FIT. {tip}")
                return
            except Exception as e:
                self.post(f"ERROR|Copy failed: {e}")
                return
            saved_paths.append(dest)

        # Eject after copy
        # NEW RULE: In Copy-Only mode, ALWAYS eject
        eject_result = None
        if self.copy_only or self.unmount_after_copy:
            if IS_MAC:
                eject_result = mac_eject(root)
            else:
                eject_result = win_eject_drive(root)

        # Email or copy-only completion
        if self.copy_only:
            # Log and profile
            prof = {
                    "device_id": device_id,
                    "model": model,
                    "last_copied_files": [p.name for p in saved_paths],
                    "last_action": "copy_only",
                    "last_time": datetime.now().isoformat(timespec="seconds"),
                    }
            try:
                (DEVICES_DIR / dev_id_for_fs / "profile.json").write_text(json.dumps(prof, indent=2), encoding="utf-8")
            except Exception:
                pass

            elapsed = int(time.time() - self.start_ts)
            for src, dest in zip(selected, saved_paths):
                log_line(
                        f"COPIED  label={(label or '')}  file={dest}  src={src.name}  "
                        f"device_id={(device_id or '')}  model={model}  duration={elapsed}s  mode=COPY_ONLY"
                        )

            text = "Eject successful, please attach the next watch to the USB cable."
            self.post("DONE|" + text + "|100|" + str(sent_dir) + "|MODE:COPY_ONLY")
            return

        # Send single email with all attachments
        if self._check_cancel():
            return
        body_text = read_mail_body_with_name(name)
        self.post(f"STEP|Sending email... ({len(saved_paths)} attachment(s))|90")
        try:
            send_email_gmail(conf, email, saved_paths, body_text)  # type: ignore[arg-type]
        except smtplib.SMTPAuthenticationError:
            self.post("ERROR|AUTH: Gmail rejected login. Use an App Password in mailer.conf.json.")
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
                "last_sent_files": [p.name for p in saved_paths],
                "last_sent_time": datetime.now().isoformat(timespec="seconds"),
                }
        try:
            (DEVICES_DIR / dev_id_for_fs / "profile.json").write_text(json.dumps(prof, indent=2), encoding="utf-8")
        except Exception:
            pass

        elapsed = int(time.time() - self.start_ts)
        for src, dest in zip(selected, saved_paths):
            log_line(
                    f"SENT  label={(label or '')}  name={name}  email={email}  file={dest}  "
                    f"src={src.name}  device_id={(device_id or '')}  model={model}  duration={elapsed}s  mode=EMAIL"
                    )

        self.post("DONE|Email sent.|100|" + str(sent_dir) + "|MODE:EMAIL")

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
        self.email_entry = ttk.Entry(email_row, textvariable=self.email_var, width=48)
        self.email_entry.grid(row=0, column=0, sticky="ew")
        self.submit_btn = ttk.Button(email_row, text="Submit", command=self._submit, state="disabled")
        self.submit_btn.grid(row=0, column=1, sticky="e", padx=(8, 0))

        # Checkboxes
        opts_row = ttk.Frame(frm)
        opts_row.grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 0))
        self.unmount_var = tk.BooleanVar(value=True)      # default checked
        self.copyonly_var = tk.BooleanVar(value=False)    # default unchecked
        self.unmount_cb = ttk.Checkbutton(opts_row, text="Unmount after copy", variable=self.unmount_var)
        self.copyonly_cb = ttk.Checkbutton(opts_row, text="Copy only, do not send mail", variable=self.copyonly_var, command=self._on_copyonly_toggle)
        self.unmount_cb.pack(side="left", padx=(0, 16))
        self.copyonly_cb.pack(side="left")

        self.hint = ttk.Label(frm, text="Attach your Garmin via USB (Mass Storage).", style="Small.TLabel")
        self.hint.grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 10))

        self.pb = ttk.Progressbar(frm, orient="horizontal", mode="determinate", maximum=100)
        self.pb.grid(row=4, column=0, columnspan=2, sticky="ew")

        self.status_var = tk.StringVar(value="Waiting for name and email")
        ttk.Label(frm, textvariable=self.status_var).grid(row=5, column=0, columnspan=2, sticky="w", pady=(10, 0))

        # Open Sent Folder button
        self.open_sent_btn = ttk.Button(frm, text="Open Sent Folder", command=self._open_sent, state="disabled")
        self.open_sent_btn.grid(row=6, column=0, columnspan=2, sticky="w", pady=(8, 0))

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

        # Bindings
        self.after(100, self.name_entry.focus_set)
        self.email_entry.bind("<Return>", lambda _e: self._submit())  # Enter in Email triggers submit (B2)
        self.name_var.trace_add("write", self._validate_form)
        self.email_var.trace_add("write", self._validate_form)
        self.after(100, self._drain_queue)

        # Initial UI reflect + start watcher for auto-start in copy-only mode
        self._reflect_copyonly_state()
        self.after(1000, self._watch_mount_if_copyonly)

    # --- UI helpers ---------------------------------------------------------
    def _validate_form(self, *_):
        if self.copyonly_var.get():
            self.submit_btn.configure(state="normal" if not self.running else "disabled")
            if not self.running:
                self.status_var.set("Copy-only mode: attach a watch to begin")
            return

        name_ok = bool(self.name_var.get().strip())
        email_ok = EMAIL_RE.match(self.email_var.get().strip()) is not None
        self.submit_btn.configure(state="normal" if (name_ok and email_ok and not self.running) else "disabled")
        if not self.running:
            self.status_var.set("Waiting for name and email" if not (name_ok and email_ok) else "Ready to submit")

    def _on_copyonly_toggle(self):
        self._reflect_copyonly_state()
        self._validate_form()

        # NEW: If toggled ON and a GARMIN volume is already mounted → start immediately.
        if self.copyonly_var.get() and not self.running:
            vol = find_current_garmin_volume()
            if vol is not None:
                # Reset timer for this auto-run
                self.current_submission_key = None
                self._start_flow(name="", email="", reset_timer=True,
                                 unmount_after_copy=self.unmount_var.get(),
                                 copy_only=True)
            else:
                self.status_var.set("Copy-only mode: waiting for GARMIN volume...")

    def _reflect_copyonly_state(self):
        copy_only = self.copyonly_var.get()
        # Enable/disable inputs
        state = "disabled" if copy_only else "normal"
        self.name_entry.configure(state=state)
        self.email_entry.configure(state=state)
        if copy_only:
            self.status_var.set("Copy-only mode: attach a watch to begin")
        else:
            self.status_var.set("Waiting for name and email")

    def _watch_mount_if_copyonly(self):
        """
        Poll every 1s: if copy-only mode is ON, not running, and exactly one GARMIN volume is mounted → auto-start.
        """
        try:
            if self.copyonly_var.get() and not self.running:
                vol = find_current_garmin_volume()
                if vol is not None:
                    # Reset timer for this auto-run
                    self.current_submission_key = None
                    self._start_flow(name="", email="", reset_timer=True,
                                     unmount_after_copy=self.unmount_var.get(),
                                     copy_only=True)
                else:
                    # keep a friendly status while waiting
                    if self.status_var.get().strip() == "" or "Copy-only" not in self.status_var.get():
                        self.status_var.set("Copy-only mode: waiting for GARMIN volume...")
        finally:
            self.after(1000, self._watch_mount_if_copyonly)

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

        copy_only = self.copyonly_var.get()
        name = self.name_var.get().strip()
        email = self.email_var.get().strip()

        if not copy_only:
            if not name:
                messagebox.showinfo("Garmin Mailer", "Please enter a name.")
                return
            if not EMAIL_RE.match(email):
                messagebox.showinfo("Garmin Mailer", "Please enter a valid email address.")
                return

        # Timer reset per watch
        submission_key = f"{'COPYONLY' if copy_only else name}|{'COPYONLY' if copy_only else email}"
        reset_timer = (self.current_submission_key is None) or (submission_key != self.current_submission_key)
        self.current_submission_key = submission_key

        self._start_flow(name, email, reset_timer, self.unmount_var.get(), copy_only)

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

    def _start_flow(self, name: str, email: str, reset_timer: bool, unmount_after_copy: bool, copy_only: bool) -> None:
        self.running = True
        self.retry_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")
        self.submit_btn.configure(state="disabled")
        self.cancel_event.clear()
        self.open_sent_btn.configure(state="disabled")
        self._set_status("Detecting Garmin watch...", None)
        self._show_detect_countdown(DETECT_TIMEOUT)
        self._set_pb_indeterminate(True)
        self._start_timer(reset=reset_timer)
        self.worker = Worker(
                self.queue, name, email, self, self.cancel_event,
                unmount_after_copy=unmount_after_copy,
                copy_only=copy_only
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
        copy_only = self.copyonly_var.get()
        if copy_only or (name and EMAIL_RE.match(email)):
            self._set_pb_indeterminate(True)
            self.running = True
            self.retry_btn.configure(state="disabled")
            self.cancel_btn.configure(state="normal")
            self.cancel_event.clear()
            self.worker = Worker(
                    self.queue, name, email, self, self.cancel_event,
                    unmount_after_copy=self.unmount_var.get(),
                    copy_only=copy_only
                    )
            self.worker.start()
        else:
            messagebox.showinfo("Garmin Mailer", "Enter a valid name and email first (or enable Copy-only).")

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
                    mode = "EMAIL" if mode_tag.endswith("EMAIL") else "COPY_ONLY"
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
                    raw = parts[1]
                    try:
                        paths = [Path(s) for s in json.loads(raw)]
                    except Exception:
                        paths = []
                    chosen = FileChoiceDialog.choose(self, paths) if paths else None
                    if self.worker and hasattr(self.worker, "pick_reply_queue"):
                        if chosen:
                            self.worker.pick_reply_queue.put([str(p) for p in chosen])
                            self._set_status(f"Selected {len(chosen)} file(s).", None)
                        else:
                            self.worker.pick_reply_queue.put(None)
                            self._set_status("No file selected.", None)
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
        self.open_sent_btn.configure(state="normal")
        try:
            self.bell()
        except Exception:
            pass

        if mode == "EMAIL":
            if IS_MAC:
                try:
                    subprocess.run(
                            ["osascript", "-e", 'display notification "Email sent." with title "Garmin Mailer"'],
                            check=False
                            )
                except Exception:
                    pass
            self.status_var.set("Email sent.")
            self.pb["value"] = 100

            # NEW in V9: clear BOTH name and email ~2s after success
            def _clear_inputs():
                self.name_var.set("")
                self.email_var.set("")
                self.status_var.set("Waiting for name and email")
                self._validate_form()
            self.after(2000, _clear_inputs)
            self.submit_btn.configure(state="normal")
        else:
            # COPY_ONLY: show prompt for next watch
            self.status_var.set(message)
            self.pb["value"] = 100
            messagebox.showinfo("Garmin Mailer", message)
            # Ready for next cycle (fields remain disabled in copy-only)
            self._validate_form()
            self.submit_btn.configure(state="normal")

    def _on_error(self, _text: str) -> None:
        self._stop_timer()
        self.running = False
        self.cancel_btn.configure(state="disabled")
        try:
            self.bell()  # audible ding on error/timeouts
        except Exception:
            pass
        self.retry_btn.configure(state="normal")
        self._validate_form()

    def _open_sent(self) -> None:
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
