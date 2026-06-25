# Debug Session: ztd-sensitivity-index

Status: [RESOLVED_PENDING_CLEANUP]

## Symptom

样本 `F:\个人文件\P3111\5.21.1.zmx` 在 Standalone GUI 公差分析中，Zemax 公差工具运行成功，ZTD 已保存，但读取 ZTD 并导出统计 Excel 时失败：

- 结果目录：`F:\个人文件\P3111\公差\公差分析_5.21.1_20260626_004518`
- ZTD：`F:\个人文件\P3111\公差\公差分析_5.21.1_20260626_004518\5.21.1_tol.ZTD`
- Zemax Succeeded：True
- ErrorMessage：无
- Progress：100
- 失败信息：`cannot parse Sensitivity data: 索引超出了数组界限。`

## Initial Hypotheses

1. H1: `5.21.1` 这个 ZTD 的 Sensitivity 矩阵列数/行数少于 `ztd_reader` 预期，读取完整 Sensitivity 时越界。
2. H2: 当前自动统计虽然传了 `num_items=13`，但底层仍先解析完整 Sensitivity，导致无关列触发越界。
3. H3: 该镜头的 ZTD 只包含 Monte Carlo 数据，Sensitivity 数据结构为空或不完整。
4. H4: REPORT/TSC 表头与 ZTD 实际列数量不一致，导致标签映射越界。
5. H5: 统计 Excel 导出阶段使用了超出 `data` 实际列数的列索引。

## Evidence Collected

1. 失败目录产物完整，已生成：
   - `5.21.1_tol.ZTD`
   - `5.21.1_tol.ZDA`
   - `5.21.1_tol.zmx`
   - `5.21.1_tolMC_T0001.zmx` ~ `T0005.zmx`
   - `5.21.1_tolMC_BEST.zmx`
   - `5.21.1_tolMC_WORST.zmx`
   - `run.log`
   - `run_config.json`
   - `used_excel.xlsx`
2. 失败目录缺少 `5.21.1_tol_统计.xlsx`。
3. `run.log` 显示 Zemax 公差工具成功：`Succeeded=True`，`Progress=100`，ZTD 已保存。
4. 失败发生在 `ZTD 自动统计分项: 13 项` 后，进入 `正在读取 ZTD 并导出统计 Excel…` 阶段。
5. 错误为 `cannot parse Sensitivity data: 索引超出了数组界限。`，说明 `ToleranceDataViewer` 加载 ZTD 时解析 Sensitivity 数据失败，而不是 TDE/MFE/TSC/MonteCarlo 运行失败。

## Evidence After Instrumentation

用户复测结果：

- `RunAndWaitForCompletion=False`
- `Succeeded=False`
- `ErrorMessage=cannot parse Sensitivity data: 索引超出了数组界限。`
- `MonteCarloData.Values.Rows=5`
- `MonteCarloData.Values.Cols=110`

结论：H2/H3 成立。Zemax API 在加载 ZTD 时因 Sensitivity 数据解析失败标记工具失败，但 Monte Carlo 主矩阵实际可访问，因此统计 Excel 可降级从 Monte Carlo 主矩阵读取。

## Fix Applied

在 `ztd_reader.read_ztd()` 中增加最小降级逻辑：

- 如果 `ToleranceDataViewer` 返回失败；
- 且错误信息包含 `Sensitivity`；
- 且 `MonteCarloData.Values.Rows/Cols` 可读且大于 0；
- 则不直接失败，继续按现有逻辑读取 Monte Carlo 主矩阵，并在结果提示中写入降级说明。

## Verification

用户再次运行 `5.21.1.zmx`，结果目录：`F:\个人文件\P3111\公差\公差分析_5.21.1_20260626_005222`。

验证结果：

- 日志包含 Sensitivity 降级提示。
- 统计 Excel 已生成：`5.21.1_tol_统计.xlsx`。
- ZTD、ZDA、工作副本、5 个 MC case、Best/Worst、run.log、run_config.json、used_excel.xlsx 均生成。
- GUI 显示 `✅ 分析完成。`

## Next Steps

1. 一阶段收尾前可删除或归档本调试记录。
2. 做一阶段收尾审查。
