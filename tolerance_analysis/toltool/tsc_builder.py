"""tsc_builder.py —— 由 Excel REPORT 配置生成 .TSC 公差脚本。

职责（需求文档 §4.6/§7、skill §2c）：
- 读「输入_REPORT」启用行，生成自定义脚本判据用的 .TSC：
    LOADMERIT <名>.MF
    <优化行>            # 按补偿器模式：无/OPTIMIZE n/OPTIMIZE n OD
    REPORT "<标签>" <MF行号>
- 有后焦补偿面时用双 MF：
    LOADMERIT <comp>.MF → 补偿优化行 → LOADMERIT <report>.MF → REPORT 各行
- TSC 放 <DataDir>\\Tolerance\\，Zemax 会列入 ListOfCriteriaScripts，
  CriterionScript 取其整数索引（由 tol_runner 处理）。
- REPORT 行号与 MFE 行号一一对应；每条 REPORT 在蒙卡结果中成为独立分项列。

补偿器模式 → 优化行映射：
  无            → 不写优化行
  全部优化(DLS) → OPTIMIZE n
  全部优化(OD)  → OPTIMIZE n OD

注意（skill §2e-4）：TSC 标签尽量用 ASCII，避免读取端编码问题。
"""

from __future__ import annotations

import os


def _optimize_line(comp_mode: str, optimize_cycles: int) -> str | None:
    """按补偿器模式返回 TSC 优化行；无补偿返回 None（不写）。"""
    key = str(comp_mode or "").strip().lower().replace(" ", "")
    n = int(optimize_cycles)
    if key in ("无", "none", ""):
        return None
    if key in ("全部优化od", "全部优化(od)", "od"):
        return f"OPTIMIZE {n} OD"
    return f"OPTIMIZE {n}"


def build_tsc_lines(report_rows: list[dict], mf_name: str,
                    optimize_cycles: int = 4, comp_mode: str = "无",
                    comp_mf_name: str | None = None) -> list[str]:
    """根据 REPORT 行生成 TSC 文本行列表。

    report_rows: 「输入_REPORT」各行 dict（含 启用/标签/MF行号）。
    mf_name: 报告 MF 的 .MF 文件名（LOADMERIT，TSC 同数据目录）。
    comp_mode: 补偿器模式，决定优化行（无/OPTIMIZE/OPTIMIZE OD）。
    comp_mf_name: 补偿专用 MF 文件名；提供则用双 MF（先补偿 MF + 优化，
                  再载报告 MF 出 REPORT）。
    """
    def _mf(name: str) -> str:
        return name if name.lower().endswith(".mf") else name + ".MF"

    report_mf = _mf(mf_name)
    opt_line = _optimize_line(comp_mode, optimize_cycles)

    lines: list[str] = []
    if comp_mf_name:
        lines.append(f"LOADMERIT {_mf(comp_mf_name)}")
        if opt_line:
            lines.append(opt_line)
        lines.append(f"LOADMERIT {report_mf}")
    else:
        lines.append(f"LOADMERIT {report_mf}")
        if opt_line:
            lines.append(opt_line)

    count = 0
    for row in report_rows:
        if str(row.get("启用", "")).strip().upper() != "Y":
            continue
        label = str(row.get("标签", "")).strip()
        mf_line = row.get("MF行号")
        if not label or mf_line in (None, ""):
            continue
        n = int(float(mf_line))
        lines.append(f'REPORT "{label}" {n}')
        count += 1

    if count == 0:
        raise ValueError("REPORT sheet 中没有启用的有效行，无法生成 TSC。")
    return lines


def tsc_dir(zos_system) -> str:
    """返回 <DataDir>\\Tolerance\\ 目录（不存在则创建）。

    实测：TDE 无 ToleranceDirectory 属性；数据根目录可由
    MFE.MeritFunctionDirectory 的父目录得到（= App.ZemaxDataDir）。
    """
    mf_dir = zos_system.MFE.MeritFunctionDirectory
    data_dir = os.path.dirname(mf_dir.rstrip("\\/"))
    d = os.path.join(data_dir, "Tolerance")
    os.makedirs(d, exist_ok=True)
    return d


def default_tsc_path(zos_system, base_name: str) -> str:
    if not base_name.lower().endswith(".tsc"):
        base_name += ".TSC"
    return os.path.join(tsc_dir(zos_system), base_name)


def write_tsc(lines: list[str], tsc_path: str) -> str:
    """写 TSC 文件。

    目标文件可能被 OpticStudio 瞬时占用（刚 Save 同名镜头时），
    遇 PermissionError 短暂重试，仍失败则回退到带 _api 后缀的新文件名。
    """
    import time
    text = "\n".join(lines) + "\n"

    def _do_write(path: str) -> None:
        with open(path, "w", encoding="utf-16") as f:
            f.write(text)

    last_err = None
    for attempt in range(20):
        try:
            _do_write(tsc_path)
            return tsc_path
        except PermissionError as e:
            last_err = e
            time.sleep(1.0)

    root, ext = os.path.splitext(tsc_path)
    alt = root + "_api" + ext
    try:
        _do_write(alt)
        return alt
    except PermissionError:
        raise last_err



def build_and_write(zos_system, report_rows: list[dict], mf_name: str,
                    base_name: str, optimize_cycles: int = 4,
                    comp_mode: str = "无",
                    comp_mf_name: str | None = None) -> tuple[int, str]:
    """生成并写出 TSC。返回 (REPORT 条数, TSC 路径)。"""
    lines = build_tsc_lines(report_rows, mf_name, optimize_cycles,
                            comp_mode=comp_mode, comp_mf_name=comp_mf_name)
    path = default_tsc_path(zos_system, base_name)
    write_tsc(lines, path)
    n_report = sum(1 for ln in lines if ln.startswith("REPORT"))
    return n_report, path
