# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['justinb.py'],
    pathex=[],
    binaries=[],
    datas=[('KorRV.json', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='justinb',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=True,         # macOS GUI 앱용 필수 옵션
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.icns',            # .icns 아이콘 파일
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='justinb.app',          # .app 이름 지정!
)
