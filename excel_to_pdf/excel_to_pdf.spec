# -*- mode: python ; coding: utf-8 -*-
# Ordinal 380 / pywin32 uyumluluğu için onedir (klasör) modunda derleme önerilir.

from PyInstaller.utils.hooks import collect_all

# pywin32 DLL'lerini ve modüllerini doğru topla (ordinal hatası önlemi)
pywin32_datas, pywin32_binaries, pywin32_hiddenimports = collect_all('pywin32')

a = Analysis(
    ['excel_to_pdf.py'],
    pathex=[],
    binaries=pywin32_binaries,
    datas=pywin32_datas,
    hiddenimports=[
        'win32com',
        'win32com.client',
        'win32com.client.gencache',
        'pywintypes',
        'pythoncom',
    ] + pywin32_hiddenimports,
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
    name='excel_to_pdf',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='excel.ico' if __import__('os').path.exists('excel.ico') else None,
)

# Klasör modu: dist/excel_to_pdf/ içinde .exe + _internal (DLL'ler)
# Ordinal 380 hatasını önler; kurulum için klasörü zipleyip dağıtabilirsiniz.
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='excel_to_pdf',
)
