"""ztd_reader.py —— 读取公差分析 ZTD 文件并做分项统计（界面解耦）。

职责（需求文档 §7、skill §2e）：
- 打开 ToleranceDataViewer，加载 ZTD，读 MonteCarloData 矩阵。
- 矩阵列布局（probe 实测，05304_tol.ZTD 200x110）：
    col0      = 综合判据（评价函数总标量）
    col1..colN = N 个 REPORT 分项，顺序与 TSC 的 REPORT 完全一致
    其后各列 = 各单项公差灵敏度数据（本模块不解析）
- 对每个分项列统计：有效次数 N / 均值 / 标准差 / 最好 / 最差 / 2σ。
- 标签优先用调用方传入的 REPORT 标签；否则从 Summary「相对评估脚本」段解析。

读取陷阱（skill §2e，务必遵守）：
- 读结果前工具一次只能开一个；本模块自行 OpenToleranceDataViewer。
- 矩阵 Rows 对大矩阵可能惰性返回 0，不依赖它做边界；硬编码 nrow=num_runs。
- 不在主循环前对 Values 做任何预读（GetValueAt warmup / Data / TotalLength 轮询）。
- Values 后直接进 GetValueAt(i,j) 双重循环、行优先连续读满全部列、循环内禁 I/O。
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field


@dataclass
class ItemStat:
    label: str
    col: int
    n: int
    mean: float
    std: float
    best: float
    worst: float
    two_sigma: float
    nominal: float = float("nan")


@dataclass
class ZtdResult:
    succeeded: bool
    ztd_path: str
    num_runs: int
    num_cols: int
    items: list = field(default_factory=list)
    summary: str = ""
    message: str = ""


_SCRIPT_LINE = re.compile(r"^\s*(.+?)\s*=\s*\t?\s*([-+0-9.eE]+)\s*$")


def parse_summary_labels(summary: str):
    """从 Summary 的「相对评估脚本」段解析 (标签, 标称值) 列表，按出现顺序。"""
    out = []
    in_block = False
    for raw in str(summary).splitlines():
        line = raw.rstrip()
        if "相对评估脚本" in line or "Relative" in line:
            in_block = True
            continue
        if in_block:
            m = _SCRIPT_LINE.match(line)
            if m:
                label = m.group(1).strip()
                try:
                    nominal = float(m.group(2))
                except ValueError:
                    nominal = float("nan")
                out.append((label, nominal))
            elif line.strip() == "" and out:
                break
    return out


def _stats(values):
    """对一列样本计算 N/均值/标准差(样本)/最好/最差/2σ。

    'best'/'worst' 仅按数值大小返回最小/最大；调用方按指标含义解读
    （点列/GENC 越小越好，MTF 越大越好），统计本身不预设方向。
    """
    vals = [v for v in values if v is not None and not math.isnan(v)]
    n = len(vals)
    if n == 0:
        return 0, float("nan"), float("nan"), float("nan"), float("nan"), float("nan")
    mean = sum(vals) / n
    if n > 1:
        var = sum((v - mean) ** 2 for v in vals) / (n - 1)
        std = math.sqrt(var)
    else:
        std = 0.0
    return n, mean, std, min(vals), max(vals), 2.0 * std


def read_ztd(zos_system, ztd_path: str, num_runs: int,
             report_labels=None, num_items=None) -> ZtdResult:
    """打开 ZTD，统计 col1..colN 各分项。

    zos_system: 已连接的 PrimarySystem。
    ztd_path: ZTD 文件绝对路径。
    num_runs: 蒙特卡洛次数（用作行数硬编码边界，规避 Rows 惰性返回 0）。
    report_labels: 调用方提供的分项标签列表（顺序同 TSC REPORT）；优先使用。
    num_items: 分项数；缺省由 report_labels / Summary 推断。
    """
    dv = zos_system.Tools.OpenToleranceDataViewer()
    try:
        dv.FileName = ztd_path
        ok = bool(dv.RunAndWaitForCompletion())
        if not ok or not dv.Succeeded:
            return ZtdResult(False, ztd_path, num_runs, 0,
                             message=str(dv.ErrorMessage) or "加载 ZTD 失败")

        summary = ""
        try:
            summary = str(dv.Summary)
        except Exception:
            pass

        parsed = parse_summary_labels(summary)
        if report_labels:
            labels = list(report_labels)
        elif parsed:
            labels = [p[0] for p in parsed]
        else:
            labels = []

        if num_items is not None:
            n_items = int(num_items)
        elif labels:
            n_items = len(labels)
        elif parsed:
            n_items = len(parsed)
        else:
            n_items = 0

        nominals = {p[0]: p[1] for p in parsed}

        mc = dv.MonteCarloData
        v = mc.Values
        ncol = int(v.Cols)

        nrow = int(num_runs)
        if n_items <= 0:
            n_items = max(0, ncol - 1)

        first_col = 1
        last_col = min(first_col + n_items, ncol)
        cols = list(range(first_col, last_col))

        buckets = {c: [] for c in cols}
        for i in range(nrow):
            for j in cols:
                buckets[j].append(v.GetValueAt(i, j))

        items = []
        for idx, c in enumerate(cols):
            label = labels[idx] if idx < len(labels) else f"col{c}"
            n, mean, std, vmin, vmax, two_s = _stats(buckets[c])
            items.append(ItemStat(
                label=label, col=c, n=n, mean=mean, std=std,
                best=vmin, worst=vmax, two_sigma=two_s,
                nominal=nominals.get(label, float("nan")),
            ))

        return ZtdResult(
            succeeded=True, ztd_path=ztd_path, num_runs=nrow,
            num_cols=ncol, items=items, summary=summary,
        )
    finally:
        try:
            dv.Close()
        except Exception:
            pass


def format_table(result: ZtdResult) -> str:
    """把 ZtdResult 渲染成对齐文本表，便于命令行汇报。"""
    if not result.succeeded:
        return f"读取失败: {result.message}"
    lines = [
        f"ZTD: {result.ztd_path}",
        f"蒙特卡洛次数: {result.num_runs}  矩阵列数: {result.num_cols}  "
        f"分项数: {len(result.items)}",
        "",
        f"{'分项':<14}{'N':>5}{'标称':>14}{'均值':>14}"
        f"{'标准差':>14}{'最好':>14}{'最差':>14}{'2σ':>14}",
    ]
    for it in result.items:
        lines.append(
            f"{it.label:<14}{it.n:>5}{it.nominal:>14.6g}{it.mean:>14.6g}"
            f"{it.std:>14.6g}{it.best:>14.6g}{it.worst:>14.6g}"
            f"{it.two_sigma:>14.6g}"
        )
    return "\n".join(lines)
