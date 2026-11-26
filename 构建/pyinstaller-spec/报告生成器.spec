# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['C:\\Users\\admin\\Desktop\\报告生成器\\构建\\exe_launcher.py'],
    pathex=['C:\\Users\\admin\\Desktop\\报告生成器\\app'],
    binaries=[],
    datas=[('C:\\Users\\admin\\Desktop\\报告生成器\\app\\templates', 'templates'), ('C:\\Users\\admin\\Desktop\\报告生成器\\app\\static', 'static')],
    hiddenimports=['docxcompose.composer', 'win32timezone'],
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
    name='报告生成器',
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
)
