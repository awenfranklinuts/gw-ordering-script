# -*- mode: python ; coding: utf-8 -*-
# Single-file, single-click build. neto_api is a normal `import neto_api` in
# gw_order_tool.py, so PyInstaller's import analysis picks it up automatically.
# Build with: pyinstaller "GW Order Tool.spec" (see build.bat).
#
# Selenium and the bundled chromedriver.exe used to be embedded here for
# neto_scraper. Neto stock now comes from the HTTP API (neto_api.py), which is
# stdlib-only, so both are gone — along with the whole class of failures they
# caused: "Unable to obtain driver for chrome" when Selenium Manager couldn't
# reach the internet, and driver/Chrome major-version mismatches after a Chrome
# auto-update. The exe is also considerably smaller as a result.
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = []
# pdfminer.six (used by pdfplumber) loads its CMap data files via package
# resources at runtime, which PyInstaller's import analysis won't pick up on
# its own — collect_all ships that data so PDF parsing works in the exe.
tmp_ret = collect_all('pdfminer')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['gw_order_tool.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    name='GW Order Tool',
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
