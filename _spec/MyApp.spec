# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['D:\\我的数据\\文档\\GitHub\\Y2026S\\run.py'],
    pathex=[],
    binaries=[],
    datas=[('D:\\我的数据\\文档\\GitHub\\Y2026S\\run_ai.vbs', '.'), ('D:\\我的数据\\文档\\GitHub\\Y2026S\\AItest_ai.jsx', '.')],
    hiddenimports=['get_best'],
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
    name='MyApp',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    contents_directory='_internal',
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='MyApp',
)
