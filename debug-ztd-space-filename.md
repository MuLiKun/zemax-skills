# Debug Session: ztd-space-filename

Status: [RESOLVED_PENDING_CLEANUP]

## Symptom

样本 `F:\个人文件\P3111\05304 - 副本.zmx` 在 Standalone GUI 公差分析中，Zemax 公差工具返回成功，但程序未找到目标 ZTD：

- 结果目录：`F:\个人文件\P3111\公差\公差分析_05304_-_副本_20260626_003443`
- 期望 ZTD：`05304 - 副本_tol.ZTD`
- Zemax Succeeded：True
- ErrorMessage：无
- Progress：100
- 已生成 MC case / Best / Worst 的可能性待确认

## Initial Hypotheses

1. H1: Zemax 对 `TolDataFile` 中的空格或 ` - ` 处理不稳定，实际保存成了被改名/截断的 ZTD 文件。
2. H2: 当前程序只检查精确文件名，未扫描工作目录中实际生成的 `.ZTD` 候选文件。
3. H3: `FilePrefix` 包含空格导致 MC case 能保存，但 ZTD 保存路径/文件名规则不同。
4. H4: 结果目录内确实没有 ZTD，但 Zemax 把 ZTD 保存到了其他 Zemax 默认目录或使用了 sanitized 文件名。
5. H5: ZTD 未生成是 Zemax 针对空格文件名的限制，应在工具侧统一使用安全的工作副本/输出前缀。

## Evidence Collected

User-provided log shows Zemax tool success but target ZTD missing.

检查失败目录：

- 仅存在 `05304 - 副本_tol.zmx`、`05304 - 副本_tol.ZDA`、`run.log`、`run_config.json`、`used_excel.xlsx`。
- 结果目录无 `.ZTD`。
- 结果目录无 MC case / Best / Worst。
- `C:\Users\12970\Documents\Zemax\Tolerance` 未找到 `*副本*.ZTD`。

结论：问题不是统计读取，也不是只漏查 ZTD 路径；Zemax 在文件名前缀含空格时未保存 ZTD 和 MC case，虽然工具状态返回成功。

## Fix Applied

最小修复：在 `pipeline.prepare_session()` 创建工作副本时，对源镜头名使用已有 `_safe_name()` 生成安全前缀；如果源名含空格或特殊字符，日志提示实际使用的 Zemax 输出前缀。

修复后 `TolDataFile` / `FilePrefix` 将基于安全工作副本名，例如：

- 原始文件：`05304 - 副本.zmx`
- 安全工作副本：`05304_-_副本_tol.zmx`
- ZTD：`05304_-_副本_tol.ZTD`

## Verification

- `check_stage1.py` 通过。
- VS Code diagnostics: `pipeline.py` 无报错。
- 用户已重新运行 `05304 - 副本.zmx`，结果目录：`F:\个人文件\P3111\公差\公差分析_05304_-_副本_20260626_004019`。
- 新日志确认 `FilePrefix` / `TolDataFile` 使用安全前缀 `05304_-_副本_tol`。
- 新结果目录已生成 ZTD、统计 Excel、5 个 MC case、Best/Worst、run.log、run_config.json、used_excel.xlsx。

## Next Steps

1. 一阶段收尾前可删除或归档本调试记录。
2. 继续执行 `5.21.1.zmx` 多点号样本回归。
