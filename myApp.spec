# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['myApp.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\Admin\\OneDrive\\Desktop\\MMO_tool\\auto_submit_form_IRS\\venv\\Lib\\site-packages\\selenium_stealth\\js', 'selenium_stealth/js/')],
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
    name='myApp',
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
)
