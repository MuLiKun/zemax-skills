# zemax-skills

基于 Zemax OpticStudio ZOS-API 的光学自动化工作区，目前包含**公差分析自动化程序**与配套的 Trae Skills。

## 仓库内容

```
agentstudy/
├── tolerance_analysis/        公差分析自动化程序（主项目）
│   ├── gui.py                 图形界面入口（PySide6，暗色主题，推荐）
│   ├── main.py / tol_run.py   命令行入口
│   ├── toltool/               核心代码包（连接/建表/运行/读结果）
│   ├── 方案A_全自动 / 方案B_手动观察
│   ├── tol_config_*.xlsx      配置/模板
│   └── 公差分析程序_使用说明.md / _需求文档.md
├── .trae/skills/              Trae 技能（zemax-zosapi-connector / zemax-tolerance-analysis）
└── .gitignore
```

> `.venv/`（虚拟环境）不入库，使用者需自行创建并安装依赖。

## 公差分析程序能做什么

- 连接 OpticStudio（交互扩展 / 独立实例）
- 按 Excel 配置自动写公差表（TDE）、评价函数（MFE）、生成 TSC 脚本
- 跑脚本式蒙特卡洛公差分析，保存 ZTD / Worst / Best Case
- 读取 ZTD，把点列 / GENC / 几何 MTF 等各 REPORT 分项独立成列做统计

## 快速上手

详见 [tolerance_analysis/公差分析程序_使用说明.md](tolerance_analysis/公差分析程序_使用说明.md)。

图形界面（推荐，暗色主题，实时日志 + 运行计时）：

```powershell
.\.venv\Scripts\python.exe -u tolerance_analysis\gui.py
```

命令行：

```powershell
cd tolerance_analysis
..\.venv\Scripts\python.exe -u main.py
```

调用顺序：入口 → `pipeline`（依次 excel_io / zos_connect / tde_builder / mfe_builder / tsc_builder）→ `tol_runner` 跑蒙卡 →（tol_run.py 再调）`ztd_reader` 出统计。

## 环境

- Zemax OpticStudio 2023 R1（ZOS-API / .NET）
- Python + pythonnet、openpyxl、numpy、PySide6（GUI）
