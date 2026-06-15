"""mfe_builder.py —— 由 Excel 评价函数配置重建 MFE 并保存 .MF。

职责（需求文档 §4.5/§7）：
- 按「输入_评价函数」各行重建评价函数编辑器（MFE）。
- SaveMeritFunction() 存成 .MF，供 TSC 的 LOADMERIT 使用。

实测 API（probe 验证）：
- mfe.AddOperand() -> IMFERow；mfe.RemoveOperandAt(i)
- row.ChangeType(MeritOperandType.XXX) -> bool
- 参数列通过 row.GetOperandCell(MeritColumn.ParamN) 访问；
  Param1/2/3 为整数列、Param4 等为浮点列；用 cell.Value(字符串) 通用赋值。
- row.Target / row.Weight 直接属性。
- 各操作数 Param 语义不同（用户在 Excel 直接填 Param1..Param8）：
  RSCE: P1采样 P2波长 P4视场系数
  GENC: P1采样 P2波长 P3视场号 P4能量比 P5...
  GMTT/GMTS: P1采样 P2波长 P3视场号 P4空间频率
"""

from __future__ import annotations

import os

_PARAM_COLS = ["Param1", "Param2", "Param3", "Param4",
               "Param5", "Param6", "Param7", "Param8"]


def _to_cell_str(v) -> str | None:
    if v is None or (isinstance(v, str) and v.strip() == ""):
        return None
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def build_mfe(zos_system, mfe_rows: list[dict], clear: bool = True) -> int:
    """按 Excel 行重建 MFE。返回写入的操作数行数。"""
    import ZOSAPI

    mfe = zos_system.MFE
    MOT = ZOSAPI.Editors.MFE.MeritOperandType
    MC = ZOSAPI.Editors.MFE.MeritColumn

    rows_sorted = sorted(
        mfe_rows,
        key=lambda r: float(r.get("行号")) if r.get("行号") not in (None, "") else 1e9,
    )

    if clear:
        while mfe.NumberOfOperands > 1:
            mfe.RemoveOperandAt(mfe.NumberOfOperands)
        first = mfe.GetOperandAt(1)
        first.ChangeType(MOT.BLNK)

    written = 0
    for row in rows_sorted:
        op = str(row.get("操作数", "")).strip().upper()
        if not op:
            continue
        op_enum = getattr(MOT, op, None)
        if op_enum is None:
            raise ValueError(f"MFE 不支持的操作数: {op}")
        r = mfe.AddOperand()
        if not r.ChangeType(op_enum):
            raise RuntimeError(f"MFE ChangeType 失败: {op}")
        for nm in _PARAM_COLS:
            sval = _to_cell_str(row.get(nm))
            if sval is None:
                continue
            cell = r.GetOperandCell(getattr(MC, nm))
            cell.Value = sval
        if row.get("目标") not in (None, ""):
            r.Target = float(row.get("目标"))
        if row.get("权重") not in (None, ""):
            r.Weight = float(row.get("权重"))
        written += 1
    return written


def save_mf(zos_system, mf_path: str) -> str:
    """保存当前 MFE 到 .MF 文件。返回路径。"""
    zos_system.MFE.SaveMeritFunction(mf_path)
    return mf_path


def default_mf_path(zos_system, base_name: str) -> str:
    """返回 <DataDir>\\MeritFunction\\<base_name>.MF 路径。"""
    mf_dir = zos_system.MFE.MeritFunctionDirectory
    if not base_name.lower().endswith(".mf"):
        base_name += ".MF"
    return os.path.join(mf_dir, base_name)


def build_and_save(zos_system, mfe_rows: list[dict], base_name: str) -> tuple[int, str]:
    """重建 MFE 并保存 .MF。返回 (行数, .MF 路径)。"""
    n = build_mfe(zos_system, mfe_rows)
    path = default_mf_path(zos_system, base_name)
    save_mf(zos_system, path)
    return n, path


def build_comp_mf(zos_system, base_name: str, freq_lp: float = 34.0,
                  sampling: int = 3, wave: int = 2) -> tuple[int, str]:
    """构建后焦补偿专用评价函数并保存 .MF（不沿用报告 MF）。

    内容：单条 GMTA（几何 MTF 平均），视场=1（中心视场），
    目标=1，权重=1，空间频率=freq_lp（补偿线对，默认 34 lp/mm）。
    GMTA 参数语义：P1采样 P2波长 P3视场号 P4空间频率。
    base_name 自动加后缀 _comp。返回 (行数, .MF 路径)。
    """
    import ZOSAPI

    mfe = zos_system.MFE
    MOT = ZOSAPI.Editors.MFE.MeritOperandType
    MC = ZOSAPI.Editors.MFE.MeritColumn

    while mfe.NumberOfOperands > 1:
        mfe.RemoveOperandAt(mfe.NumberOfOperands)
    first = mfe.GetOperandAt(1)
    first.ChangeType(MOT.BLNK)

    r = mfe.AddOperand()
    if not r.ChangeType(MOT.GMTA):
        raise RuntimeError("MFE ChangeType 失败: GMTA")
    params = {"Param1": sampling, "Param2": wave, "Param3": 1, "Param4": freq_lp}
    for nm, val in params.items():
        sval = _to_cell_str(val)
        if sval is None:
            continue
        r.GetOperandCell(getattr(MC, nm)).Value = sval
    r.Target = 1.0
    r.Weight = 1.0

    comp_base = base_name
    if comp_base.lower().endswith(".mf"):
        comp_base = comp_base[:-3]
    comp_base += "_comp"
    path = default_mf_path(zos_system, comp_base)
    save_mf(zos_system, path)
    return 1, path
