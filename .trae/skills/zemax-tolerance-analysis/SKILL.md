---
name: "zemax-tolerance-analysis"
description: "Runs Zemax tolerance analysis, merit-function operands, solves, and optimization via ZOS-API. Invoke when user asks for tolerancing, sensitivity/Monte Carlo, operands, macros, or to use the bundled ZOS-API sample code."
---

# Zemax 公差分析与操作数

用于执行公差分析（灵敏度 / 反向 / 蒙特卡洛）、构建/读取评价函数操作数、设置求解、运行优化，或复用本机 ZOS-API 示例代码。与 `zemax-zosapi-connector` 互补：先用那个建立连接，再套用这里的模式。

所有路径、文件名、数值均为**运行时需核实的示例**，会随机器和 OpticStudio 版本变化。

## 本机参考来源

不确定某个 API 名称/操作数/设置时，查这些权威资料（先核实路径存在）：

1. 用户手册 PDF：`C:\Program Files\Ansys Zemax OpticStudio 2023 R1.00\OpticStudio_UserManual_{zh,en}.pdf`。相关章节：公差分析、优化操作数、编程选项卡/关于 ZPL（MODIFYSETTINGS 关键字）。
2. ZOS-API 示例代码：`C:\Users\<user>\Documents\Zemax\ZOS-API Sample Code\`（C#/C++/MATLAB/Python/VB.NET）。路径若不同，用 `TheApplication.SamplesDir` 定位。

示例编号（各语言文件夹共用，如 `PythonStandalone_NN_*.py`）：
**14 Seq_Tolerance（公差主参考）**、15 Seq_Optimization（求解/MF向导/优化）、07 TiltDecenterAndMFOperand（坐标断点/求解/`MFE.GetOperandValue`）、03 open_file_and_optimise、04 FFTMTF、11/12 序列搭建、18 多重结构、19 面属性、22 点列、26 工程偏好。请求某流程时打开对应示例，勿凭空猜 API 名。

## 本项目已封装的实现（优先参考）

本工作区已把下述实测经验固化为可复用 Python 程序（包 `toltool/`，入口 `gui.py` / `main.py` / `tol_run.py` / 方案A/B）。改公差分析逻辑时优先看这些文件，而非从零写：

- `toltool/pipeline.py`：单一来源 `prepare_session()`（连接→TDE→MFE→TSC→Save）+ `run_montecarlo()`。
- `toltool/tol_runner.py`：配置并运行脚本式蒙卡（界面解耦，progress_cb/cancel_flag）。
- `toltool/ztd_reader.py`：读 ZTD MonteCarloData 出分项统计。
- `gui.py`：PySide6 暗色界面，QThread 后台跑蒙卡 + 信号回调日志/计时，业务全复用 pipeline；不重复任何公差逻辑。
- 使用说明见 `公差分析程序_使用说明.md`，需求见 `公差分析程序_需求文档.md`。

## 连接

用 `zemax-zosapi-connector`。GUI 文件用 `ConnectAsExtension($instance)` + `$sys=$app.PrimarySystem`（每次连接消耗一次「交互扩展」等待状态，需重新点击）；批处理用 `CreateNewApplication()`。

## 公差分析工作流（两步）

1. **定义公差**：序列模式用 `$sys.TDE.SEQToleranceWizard`，设量级、关掉不需要的项、`OK()`。需精细覆盖则直接写 TDE 行。
2. **运行工具**：`$tol=$sys.Tools.OpenTolerancing()`，配置后 `RunAndWaitForCompletion()`（阻塞）或 `Run()`+轮询（非阻塞），最后 `Close()`。

运行前与用户确认：SetupMode（灵敏度/反向/蒙卡）、Criterion（判据）、CriterionComp（补偿器）、NumberOfRuns/NumberToSave。

### 实测要点（2023 R1 验证，必须遵守）

**工具设置用索引列表，不是枚举属性：**
- `SetupModeIndex`：[0]灵敏度 [1]反极值 [2]反增量 [3]跳过灵敏度。**无独立蒙卡模式；蒙卡靠 `NumberOfRuns>0` 触发，SetupModeIndex 仍用 0。**
- `CriterionIndex`（16项）：[0]RMS光斑半径 [3]RMS波前 **[4]评价函数** [5-10]各类MTF [11]瞄准误差 **[15]自定义脚本**。
- `CriterionCompIndex`：[0]全部优化(DLS) **[1]近轴焦点** [2]无 [3]全部优化(OD)。
- `MonteCarloStatisticIndex`：[0]正态 [1]均匀 [2]抛物线。
- **PowerShell** 列表用 `foreach` 或方括号 `[$i]`（圆括号报 MethodNotFound），`.Count` 取长度（非 `.Length`）。**Python 端 `ListOf*` 返回 None**，改用 `GetXxxAt(i)`+`NumberOfXxx` 遍历。
- 保存 ZTD：**`SaveTolDataFile=$true`（默认 False，必须显式设）+ `TolDataFile=绝对路径`**（用 `os.path.abspath`，**不要用纯文件名**）。实测纯文件名依赖 Zemax 相对目录解析，在 standalone + 跨盘（镜头不在 C 盘）场景下失效——`Succeeded=True` 却不落盘任何标准目录，造成「假成功」；改绝对路径后稳定。运行后按多候选路径（绝对路径 / 镜头目录 / `Documents\Zemax\Tolerance`）回读确认文件真存在，回读不到应判失败。`OutputFile` 实测不自动落盘。
- 保存 Worst/Best：`IsSaveBestWorstUsed=true` + `FilePrefix`。**这是单一开关，无法分别只存 Worst 或只存 Best**；`BestWorstOutputFolder` 实测只读，输出位置由 Zemax 决定，运行后回读。

### 分项展示：自定义脚本判据（TSC）

Criterion 是单一标量，一次只输出一个判据分布。要把点列/GENC/MTF **分项**成列，用 `CriterionIndex=15` + 一个 `.TSC` 脚本：

- `.TSC` 放在 `<DataDir>\Tolerance\`，Zemax 列入 `ListOfCriteriaScripts`，`CriterionScript` 取其**整数索引**（Python 端列表项是文件名字符串如 `'05304_tol.TSC'`，取索引）。
- TSC 语法（非 ZPL）：`LOADMERIT xxx.MF` → `OPTIMIZE n`（优化补偿器，会接管运行参数里的补偿器设置）→ `REPORT "标签" N`（把当前 MF 第 N 行操作数值作为独立结果列）。每个 REPORT 行在蒙卡结果里是独立一列。
- 先 `$sys.MFE.SaveMeritFunction('<DataDir>\MeritFunction\xxx.MF')`，REPORT 行号与 MFE 行号一一对应。

### 完成判据与长运行

- 完成：`IsRunning` 变 False 且 `Succeeded==True`。
- 扩展模式轮询不要快于约 2 秒，否则 IPC 管道易断。
- **写 TSC 的权限坑**：后台隐藏进程（`Start-Process -WindowStyle Hidden`）启动的 python 写不了 `Documents\Zemax\Tolerance`（errno13）；前台终端进程正常。非代码问题，是进程启动方式差异。
- **TSC 模式无逐次进度（实测定论，勿再尝试做 1/N 进度）**：运行期间 `Progress` 全程 0、结束才跳 100，无任何标量属性随运行变化；MC 样本文件**跑完才一次性写**，且只写 `NumberToSave` 个（非 `NumberOfRuns` 个）。TSC 的 `SAVE` 命令只能存固定一个文件、不支持变量，**无法逐圈保存**。要保留全部样本只能把 `NumberToSave=NumberOfRuns`，但仍是跑完才写。GUI 只能诚实显示「总数 + 已用时间」，做不到逐次进度。

### 读取结果：ToleranceDataViewer

工具一次只开一个：先 `$tool.Close()`，再 `$dv=$sys.Tools.OpenToleranceDataViewer()`，设 `$dv.FileName='xxx.ZTD'` → `RunAndWaitForCompletion()`。
- `$dv.Summary`：各 REPORT 分项标称值 + 每项公差灵敏度贡献。
- `$dv.MonteCarloData.Values`：矩阵，前若干列即 TSC 的 REPORT 分项（顺序同 REPORT）。
- `$dv.SensitivityData`：灵敏度数据。

**读取陷阱（否则全 NaN / 循环不执行）：**
1. **PowerShell 变量名大小写不敏感**：`$R`(行数) 与 `$r`(计数) 是同一变量，`for($r=0;$r -lt $R)` 会清零 `$R`。循环/边界变量用不冲突的名字（`$i/$j/$nrow/$ncol`）。Python 无此坑。
2. **`Values.Rows` 对大矩阵首次返回 0**（惰性）；用已知规模硬编码 `nrow=NumberOfRuns`、`ncol`。
3. **不要预读 `Values`**（`GetValueAt(0,0)` warmup / `.Data` / `.Rows` 轮询都会干扰）；拿到 `Values` 后**直接行优先连续读满所有列**，不跳列。
4. 中文路径：Python 按 UTF-8 处理；PowerShell 用 `-File 脚本.ps1` 且脚本内全 ASCII，中文路径由 `-ZTD '...'` 命令行参数传入。
5. 取数循环内不做文件 I/O；先收进 List，循环后一次性写。

## 评价函数操作数

操作数在 `$sys.MFE`。直接求值不编辑用 `GetOperandValue(opType, 8个数值输入，未用补0)`。自动构建用 `$mfe.SEQOptimizationWizard`（设 `Data`/`Ring`/玻璃空气边界 → `Apply()`），再 `SaveMeritFunction`/`LoadMeritFunction`。操作数代码（GLCR/RSCE/RWCE/DIMX…）查手册「优化操作数」章节。

## 求解 / 优化 / 宏

- 求解：`$tools.SetAllRadiiVariable()`、`cell.MakeSolveVariable()`、`CreateSolveType(...)`+`SetSolveData(...)`（拾取/位置/F数）。参考示例 07。
- 优化：`OpenLocalOptimization()`（DLS + Cycles.Automatic + NumberOfCores）；全局/锤形用 `OpenGlobalOptimization()`/`OpenHammerOptimization()` + `RunAndWaitWithTimeout(秒)`。参考示例 15。
- 分析设置脚本化：`GetSettings().SaveTo(cfg)` → `ModifySettings(cfg,'关键字','值')` → `LoadFrom(cfg)` → `ApplyAndWaitForCompletion()`。关键字查手册「编程选项卡 > 关于 ZPL > 关键字」。

## 操作注意

- 启动耗时运行前与用户确认 SetupMode/Criterion/次数。
- **不覆盖原始文件**：分析前先 SaveAs 副本。
- 不确定 API 名/操作数/关键字时，打开对应示例或查手册，勿猜。
