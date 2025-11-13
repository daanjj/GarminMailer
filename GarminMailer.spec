# -*- mode: python ; coding: utf-8 -*-

import os
from pathlib import Path

version_kwargs = {}
version_file_env = os.environ.get("GARMIN_MAILER_VERSION_FILE")
if version_file_env:
    vf = Path(version_file_env)
    if vf.exists():
        version_kwargs["version"] = str(vf)
else:
    # Fallback: look for version.txt in current directory
    version_file = Path("version.txt")
    if version_file.exists():
        version_kwargs["version"] = str(version_file)

import sys

# Add version.txt to data files if it exists (for macOS builds)
datas = []
if os.path.exists('version.txt'):
    datas.append(('version.txt', '.'))

# Add default template file
if os.path.exists('default-mail-template.txt'):
    datas.append(('default-mail-template.txt', '.'))

a = Analysis(
    ['garmin_mail_gui.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='GarminMailer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    **version_kwargs,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='GarminMailer',
)
app = BUNDLE(
    coll,
    name='GarminMailer.app',
    icon='icon/icon.icns',
    bundle_identifier=None,
)
