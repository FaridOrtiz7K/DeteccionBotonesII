# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['versionF1.py'],
    pathex=[],
    binaries=[],
    datas=[('AutoHotkey_1.1.37.02', 'AutoHotkey_1.137.02'), ('utils', 'utils'), ('img', 'img')],
    hiddenimports=['utils.ahk_manager', 'utils.ahk_click_down', 'utils.ahk_enter', 'utils.ahk_managerCopyDelete', 'utils.ahk_writer'],
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
    name='VersionF1',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['img\\icono.ico'],
)
