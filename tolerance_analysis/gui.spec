# -*- mode: python ; coding: utf-8 -*-
r"""PyInstaller 打包配置（onedir）。

构建（在 IDE 自带终端，非 sandbox）：
    ..\.venv\Scripts\python.exe -m PyInstaller --noconfirm gui.spec

产物：dist\公差分析\公差分析.exe + 同级 tol_config_模板.xlsx。
注意：目标电脑仍需安装 Zemax OpticStudio 并有可用 license；
ZOSAPI*.dll 运行时从本机 Zemax 安装目录加载，不打进 exe。
"""

import os

block_cipher = None

_TEMPLATE = "tol_config_模板.xlsx"
datas = []
if os.path.isfile(_TEMPLATE):
    datas.append((_TEMPLATE, "."))

hiddenimports = [
    "clr",
    "pythonnet",
    "toltool",
    "toltool.pipeline",
    "toltool.zos_connect",
    "toltool.excel_io",
    "toltool.tde_builder",
    "toltool.mfe_builder",
    "toltool.tsc_builder",
    "toltool.tol_runner",
    "toltool.ztd_reader",
]

a = Analysis(
    ["gui.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["tkinter", "matplotlib", "pytest"],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="公差分析",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="公差分析",
)

if os.path.isfile(_TEMPLATE):
    import shutil
    _dst = os.path.join("dist", "公差分析", _TEMPLATE)
    if os.path.isdir(os.path.dirname(_dst)):
        shutil.copyfile(_TEMPLATE, _dst)
