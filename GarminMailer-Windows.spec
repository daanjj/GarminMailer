# -*- mode: python ; coding: utf-8 -*-

import os
from pathlib import Path

# Add default template file if it exists
datas = []
if os.path.exists('default-mail-template.txt'):
    datas.append(('default-mail-template.txt', '.'))

if os.path.exists(os.path.join('icon', 'icon.ico')):
    datas.append((os.path.join('icon', 'icon.ico'), 'icon'))

if os.path.exists(os.path.join('icon', 'GarminMailer icon.png')):
    datas.append((os.path.join('icon', 'GarminMailer icon.png'), 'icon'))

a = Analysis(
    ['garmin_mailer.py'],
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
    a.binaries,
    a.datas,
    [],
    name='GarminMailer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=r'icon\icon.ico',
)