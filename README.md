# zemax-skills

基于 Zemax OpticStudio ZOS-API 的光学自动化工作区，目前包含**公差分析自动化程序**与配套的 Trae Skills。

## 仓库内容

```
agentstudy/
├── tolerance_analysis/        公差分析自动化程序（主项目）
│   ├── gui.py                 图形界面入口（PySide6，暗色主题，推荐）
│   ├── main.py / tol_run.py   命令行入口
│   ├── check_stage1.py / check_field_mapping.py / make_backup.py  基础检查、视场映射检查与本地备份脚本
│   ├── toltool/               核心代码包（连接/建表/视场映射/运行/读结果）
│   ├── 方案A_全自动 / 方案B_手动观察
│   ├── tol_config_*.xlsx      配置/模板
│   └── 公差分析程序_使用说明.md / _需求文档.md / 开发进度与测试记录.md
├── .trae/skills/              Trae 技能（zemax-zosapi-connector / zemax-tolerance-analysis）
└── .gitignore
```

> `.venv/`（虚拟环境）不入库，使用者需自行创建并安装依赖。

## 公差分析程序能做什么

- 连接 OpticStudio（交互扩展 / 独立实例）
- 按 Excel 配置自动写公差表（TDE）、评价函数（MFE）、生成 TSC 脚本
- 支持运行前配置校验，不连接 Zemax 即可检查路径、Excel 与 MFE/REPORT 映射
- 支持高级 Excel 视场映射：按目标归一化视场匹配/插入 tol 工作副本视场，自动改写 RSCE、GENC/GMTT/GMTS/GMTA 和 REPORT 标签，并输出 `field_mapping.txt` 轻量复核报告；当前视场归一化按 X/Y 主轴处理，不再支持斜向二维视场投影
- 支持普通标准模板模式：GUI/CLI 可直接选择模板、公差等级、MC 次数、保存数量、补偿方式与是否保存 WC/BC；默认自动使用主波长、自动识别公差结束面，并启用标准模板视场映射（目标 `0,0.5,0.9`）；运行期临时标准配置使用唯一文件名，避免覆盖用户导出的同名配置
- 支持使用 Zemax 当前设置模式：复用 zmx 中已有 TDE/MFE，自动从当前 MFE 生成 REPORT/TSC；补偿器判断以 TDE 中的 `COMP` 操作数为准
- 跑脚本式蒙特卡洛公差分析，在每次运行的时间戳结果目录内保存工作副本、ZTD、日志与配置快照
- 读取 ZTD，把自定义脚本总值、REPORT 分项和 TDE 中的 `COMP` 补偿器项独立成列做统计，并导出统计 Excel；Cpk1.33 上下限双边输出，按方向标黄用户关注侧
- 支持独立分析已有 ZTD；REPORT 表头优先从同名 TSC 的 `REPORT "标签" 行号` 反推，COMP 列优先用同目录 `run_config.json` 中的 TDE 顺序定位

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

普通标准模板模式（不需要手写 Excel，当前为后台/CLI/GUI 最小可用版；模板内容集中在 `tolerance_analysis/toltool/standard_templates.py`，可后续替换为正式标准）。当前行为：主波长自动识别、结束面自动取像面前一面、优化补偿未填后焦补偿面时自动取像面前一面、标准模板视场映射默认开启并使用 `0,0.5,0.9`，运行期临时配置生成 `<镜头名>_标准模板配置_<8位uuid>.xlsx`，GUI 可勾选保存 WC/BC：

```powershell
.\.venv\Scripts\python.exe -u tolerance_analysis\tol_run.py --standard --zmx "镜头.zmx" --outdir "输出目录" --connect standalone --standard-template 快速摸底 --tolerance-level 标准 --num-runs 20
```

使用 Zemax 当前设置模式（不需要 Excel 配置，保留当前 TDE/MFE；如需补偿器优化，必须由当前 TDE 中的 `COMP` 操作数触发）：

```powershell
.\.venv\Scripts\python.exe -u tolerance_analysis\tol_run.py --current-settings --zmx "镜头.zmx" --outdir "输出目录" --connect standalone --num-runs 20 --comp-mode 全部优化DLS
```

一阶段基础检查（不连接 Zemax，含典型负向配置校验）：

```powershell
..\.venv\Scripts\python.exe -u tolerance_analysis\check_stage1.py --python ..\.venv\Scripts\python.exe
```

本地备份快照（输出到 `_backups/`，不入库）：

```powershell
..\.venv\Scripts\python.exe -u tolerance_analysis\make_backup.py
```

调用顺序：入口 → `pipeline`（依次 excel_io / zos_connect / tde_builder / mfe_builder / tsc_builder）→ `tol_runner` 跑蒙卡 →（tol_run.py 再调）`ztd_reader` 出统计。

## 环境

- Zemax OpticStudio 2023 R1（ZOS-API / .NET）
- Python + pythonnet、openpyxl、numpy、PySide6（GUI）

## 变更履历

> 仅记录功能层面的主要变更，便于追溯。日期格式 YYYY-MM-DD。

- **2026-06-29** 视场映射与标准模板临时配置修正：视场归一化改为 X/Y 主轴规则，修复纯 X 方向正负归一化不稳定问题，并移除斜向二维视场投影逻辑；普通标准模板运行期临时配置改为 UUID 唯一文件名，避免覆盖用户导出的同名标准配置。
- **2026-06-27** 当前设置与 COMP 统计增强：新增使用 Zemax 当前 TDE/MFE 模式；补偿器判断改为扫描 TDE 中的 `COMP`；ZTD 统计纳入 COMP 补偿器项并按 TDE 顺序定位；独立分析已有 ZTD 可结合 `run_config.json` 复现 COMP 列；视场映射复核输出由 `mapped_excel.xlsx` 简化为 `field_mapping.txt`。
- **2026-06-26** 普通标准模板模式增强：GUI 去除波长选择并默认使用 Zemax 主波长；运行期自动识别像面前一面作为公差结束面；优化补偿未填后焦补偿面时自动取像面前一面；标准模板视场映射默认开启，目标为 `0,0.5,0.9`；GUI 新增「保存 WC/BC」勾选项；`used_excel.xlsx` 记录运行期解析后的实际配置。
- **2026-06-26** 第二阶段普通标准模板模式启动：新增后台/CLI/GUI 最小可用版，支持 `tol_run.py --standard` 或 GUI「普通标准模板」自动生成标准模板配置并复用现有 TDE/MFE/TSC/蒙卡/统计链路。
- **2026-06-26** 高级 Excel 视场映射增强：目标归一化视场列改为可选覆盖项；空值时从 RSCE `Param4` 或 GENC/GMTT/GMTS/GMTA 原 `Param3` 自动推断；启用后输出映射复核信息，GUI/run.log 打印最终映射表和 MFE/REPORT 改写数量。
- **2026-06-26** 第一阶段完成并冻结：完成多样本 GUI 端到端回归、工作副本隔离、运行产物追溯、ZTD 统计和 Stage 1.5 视场映射接入。
- **2026-06-25** 运行产物隔离与追踪增强：每次运行进入时间戳结果目录，保存 `run.log`、`run_config.json`、`used_excel.xlsx`，并在日志中打印输出清单与保存策略。
- **2026-06-23** GUI 与统计可用性增强：GUI 记住上次文件/目录/连接模式；统计 Excel 的 Cpk1.33 上下限双边输出，并按方向标黄用户关注侧。
- **2026-06-23** 公差运行诊断增强：GUI 日志会输出 Zemax 公差工具关键设置与运行结果，失败时附带 `Succeeded`、`ErrorMessage`、`NumberToSave`、`SaveTolDataFile`、`TolDataFile` 等状态，并补充 Zemax 数据目录下 `Tolerance` 的 ZTD 查找路径。
- **2026-06-22** ZTD 统计模块增强：GUI 支持独立分析已有 ZTD，统计 Excel 默认启用并输出到 ZTD 同目录；ZTD 表头优先从同名 TSC 反推，找不到 TSC 时保留 colN 兜底。
- **2026-06-21** GUI 移植性增强：Zemax 安装目录改为分层智能查找（显式 → INI 配置 → 环境变量 `ZEMAX_ZOS_DIR` → 注册表+多盘 glob 搜索）；自动找不到时弹窗引导手动选目录并写入 `zemax_config.ini` 永久记住，换电脑免改代码。
- **2026-06-21** 新增 PyInstaller 打包（`gui.spec`，onedir 无黑窗），可生成独立 `公差分析.exe`，Excel 模板与配置模板随 exe 外置。
- **2026-06-21** 新增 PySide6 图形界面 `gui.py`（暗色主题、实时日志、运行计时、Standalone/交互扩展连接模式切换、强制停止）。
- **初始版本** 公差分析自动化核心：ZOS-API 连接、按 Excel 重建 TDE/MFE/TSC、脚本式蒙特卡洛、ZTD 分项统计；提供命令行入口（`main.py`/`tol_run.py`）与方案 A/B。
