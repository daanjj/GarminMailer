# GarminMailer AI Assistant Instructions

## Project Overview
GarminMailer is a cross-platform Tkinter desktop application that detects Garmin watches via USB Mass Storage, copies FIT activity files, and either emails them or archives them. Designed for running workshops where participants' activity data needs to be quickly processed.

## Core Architecture & Data Flow

### Single-File Application
- **Main file**: `garmin_mailer.py` (~1300 lines) - contains the entire application
- **Threading model**: Main UI thread + background `Worker` thread communicating via `queue.Queue`
- **Cross-platform detection**: Platform-specific USB volume detection for macOS (`/Volumes/`) and Windows (drive letters)

### Key Data Structures & State Management
- **Runtime paths**: All user data stored in `~/Documents/GarminMailer/`
  - `sent/YYYYMMDD/` - emailed files with naming: `YYYYMMDD_label_email_originalname.fit`
  - `archive/YYYYMMDD/` - archive-only files with naming: `YYYYMMDD_label_originalname.fit`
  - `devices/<device_id>/profile.json` - per-device metadata
  - `watch-labels.csv` - device ID to workshop label mapping
- **Threading communication**: UI thread reads worker messages via `queue.Queue` with pipe-delimited protocol (`"STEP|text|progress"`, `"ERROR|message"`, etc.)

### Core Workflows
1. **Device Detection**: 30-second timeout waiting for exactly one GARMIN volume to be mounted
2. **File Selection**: Parse `GarminDevice.xml` for device info, list FIT files from `GARMIN/Activity/` or `Activity/`
3. **Dual Modes**: 
   - Email mode: requires name/email, processes today's files only
   - Archive mode: auto-starts on mount, shows 5 most recent files for selection

## Development Patterns

### Version Management
- **Git-based versioning**: `get_app_version()` uses `git tag` to find highest `v*` tag
- **Frozen app versioning**: Uses PyInstaller `_MEIPASS` detection + platform-specific version embedding
- **Version files**: macOS uses `version.txt`, Windows uses EXE resources via `win32api`

### Cross-Platform Patterns
```python
IS_WINDOWS = sys.platform.startswith("win")
IS_MAC = (sys.platform == "darwin")

# Platform-specific implementations
if IS_MAC:
    result = mac_find_single_volume(deadline, tick_cb)
else:
    result = win_find_single_root(deadline, tick_cb)
```

### PyInstaller Resource Handling
```python
def _resource_path(rel: str) -> Path:
    base = getattr(sys, "_MEIPASS", None)
    return (Path(base) / rel) if base else (Path(__file__).resolve().parent / rel)
```

### UI-Worker Communication Protocol
Worker posts messages: `"STEP|Detecting watch...|SPIN_ON"`, `"ASK_PICK|['/path1', '/path2']"`, `"DONE|Email sent|100|/save/path|MODE:EMAIL"`

## Build & Release System

### Local Development Build
```bash
# Use VSCode task: "Build and Restart GarminMailer" 
# or manually:
pyinstaller --noconfirm GarminMailer.spec && open dist/GarminMailer.app
```

### Release Process
- **Trigger**: Push git tags matching `v*` pattern
- **GitHub Actions**: Builds both macOS (.app bundle) and Windows (.exe) versions
- **macOS**: Uses `--onedir --windowed` with `.icns` icon
- **Windows**: Uses `--onefile --noconsole` with version resource injection

### Dependencies
- **Core**: `tkinter` (GUI), `fitparse` (FIT file parsing), `certifi` (SSL)
- **Windows-only**: `pywin32` for version info extraction (optional import)
- **Email**: Uses Python's built-in `smtplib` with `EmailMessage`

## Configuration & Customization

### Application Configuration
- **Config file**: `config.json` in project root
  - `devmode`: boolean - Shows/hides advanced UI controls (default: false)
  - `only_today`: boolean - In email mode, filter to today's files only (default: true)

### Email Configuration
- **Required file**: `~/Documents/GarminMailer/mailer.conf.json`
- **Template**: Copy from `example.mailer.conf.json` (Brevo SMTP by default)
- **Email template**: `mail-template.txt` with `{name}` placeholder substitution

### Device Labeling
- **CSV format**: `watch-labels.csv` with `device_id,label` pairs
- **Auto-generation**: Files auto-created with example content on first run
- **Usage**: Labels appear in UI, filenames, and logs for workshop organization

## Error Handling Patterns
- **Worker errors**: Posted to UI queue as `"ERROR|message"` 
- **Permission issues**: Specific macOS "Full Disk Access" guidance
- **Config validation**: Explicit missing key detection with user-friendly messages
- **Graceful degradation**: Optional imports (`fitparse`, `win32api`) with fallback behavior

## Testing & Debugging
- **Dev mode**: Create `~/Documents/GarminMailer/.devmode` to show additional UI controls
- **Logging**: All actions logged to `~/Documents/GarminMailer/GarminMailer.log` with timestamps
- **UI feedback**: Progress bars, status updates, and platform-specific notifications