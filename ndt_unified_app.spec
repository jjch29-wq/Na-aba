# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['ndt_unified_app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'docx',
        'docx.oxml',
        'docx.oxml.ns',
        'docx.oxml.shared',
        'docx.shared',
        'docx.enum.text',
        'docx.text.paragraph',
        'docx.table',
        'lxml',
        'lxml._elementpath',
        'lxml.etree',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
        'PIL.ImageFile',
        'PIL.JpegImagePlugin',
        'PIL.PngImagePlugin',
        'PIL.BmpImagePlugin',
        'PIL.GifImagePlugin',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.simpledialog',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='NDT-Procedure-Manager',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # 콘솔 창 숨김
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,              # 아이콘 없음 (추후 .ico 파일 지정 가능)
)
