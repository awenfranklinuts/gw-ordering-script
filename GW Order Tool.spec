# -*- mode: python ; coding: utf-8 -*-
# Single-file, single-click build: everything gw_order_tool.py needs (selenium,
# including its bundled Selenium Manager binaries that auto-locate/download a
# matching chromedriver on the target machine) gets embedded into one exe via
# collect_all('selenium') below. neto_scraper is a normal `import neto_scraper`
# in gw_order_tool.py now, so PyInstaller's own import analysis picks it up
# automatically — it no longer needs to ship as a separate .py file next to the
# exe. Build with: pyinstaller "GW Order Tool.spec" (see build.bat).
import os
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = ['neto_scraper']  # belt-and-suspenders; real import already covers this
tmp_ret = collect_all('selenium')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# If a chromedriver.exe sits next to this spec (build.yml / build.bat download one
# matching the current Chrome stable release before invoking pyinstaller), embed it
# as a data file so neto_scraper.create_driver() can use it directly at runtime
# instead of depending on Selenium Manager downloading one on the target machine
# (that dependency on runtime network access + Chrome auto-detection is what caused
# "Unable to obtain driver for chrome" on other people's devices). Optional: if the
# file isn't present (e.g. a from-source build with no internet), the build still
# succeeds and the app falls back to Selenium Manager, same as before.
if os.path.exists('chromedriver.exe'):
    datas.append(('chromedriver.exe', '.'))


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
