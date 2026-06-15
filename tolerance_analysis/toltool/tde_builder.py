"""tde_builder.py —— 由 Excel 配置生成 TDE 公差表（基于 Zemax 原生向导）。

职责（需求文档 §4.3/§4.4/§7、§10.2）：
- 按「输入_公差向导」调用 Zemax 原生 SEQToleranceWizard 生成基础公差表，
  由 Zemax 自行做元件分组、玻璃/空气识别、阿贝数换算（物理正确）。
- 再按「输入_公差明细」对生成的 TDE 做覆盖/追加/删除（明细优先级最高）。

为什么用原生向导（实测对比结论）：
- 自写逐面展开与 Zemax 原生向导在「成对面分组、表面/元件公差区分、
  玻璃面识别、阿贝换算」四处不等价；改用 SEQToleranceWizard 规避。

SEQToleranceWizard 真实属性（probe 实测，2023 R1）：
- 范围：StartAtSurface / StopAtSurface（全局一组，向导不支持逐类范围/跳过面）。
- 半径：IsSurfaceRadiusUsed + SurfaceRadiusUnitType(DefaultAndFringes:
  0=Default,1=Fringes) + SurfaceRadiusFringes / SurfaceRadiusPercent。
- 厚度：IsSurfaceThicknessUsed + SurfaceThickness。
- 面偏心：IsSurfaceDecenterX/YUsed + SurfaceDecenterX/Y。
- 面倾斜：IsSurfaceTiltX/YUsed + SurfaceTiltXUnitType(DefaultAndDegrees) +
  SurfaceTiltXDegrees / SurfaceTiltYDegrees。
- 元件偏心：IsElementDecenterX/YUsed + ElementDecenterX/Y。
- 元件倾斜：IsElementTiltX/YUsed + ElementTiltXDegrees / ElementTiltYDegrees。
- 不规则：IsSurfaceSandAIrregularityUsed + SurfaceSandAIrregularityFringes。
- 折射率：IsIndexUsed + Index。
- 阿贝%：IsIndexAbbePercentageUsed + IndexAbbePercentage。
- 焦点补偿：IsFocusCompensationUsed。
- 测试波长：TestWavelength（um）。
- 流程：Initialize() → 设属性 → OK()（落 TDE）。
- pythonnet 约束：单位枚举须用枚举成员对象赋值，不能传 int。
"""

from __future__ import annotations

from dataclasses import dataclass


# 明细表用：操作数 → 是否成对面
PAIRED_OPS = {"TTHI", "TEDX", "TEDY", "TETX", "TETY"}


@dataclass
class TolItem:
    op: str
    surf1: int
    surf2: int
    vmin: float
    vmax: float
    comment: str = ""

    def key(self) -> tuple:
        return (self.op, self.surf1, self.surf2)


def _is_yes(v) -> bool:
    return str(v).strip().upper() in ("Y", "YES", "1", "TRUE", "是")


def _num(v, default=None):
    if v is None or (isinstance(v, str) and v.strip() == ""):
        return default
    return float(v)


# ---------------------------------------------------------------------------
# 公差向导 sheet → 原生向导设置
# ---------------------------------------------------------------------------

# 每个公差类别 → (启用开关属性, 数值属性) 列表（部分类别需设单位枚举，单独处理）
_CATEGORY_MAP = {
    "半径": ("IsSurfaceRadiusUsed", None),               # 单位特殊处理
    "厚度": ("IsSurfaceThicknessUsed", "SurfaceThickness"),
    "面偏心X": ("IsSurfaceDecenterXUsed", "SurfaceDecenterX"),
    "面偏心Y": ("IsSurfaceDecenterYUsed", "SurfaceDecenterY"),
    "面倾斜X": ("IsSurfaceTiltXUsed", None),              # 单位特殊处理
    "面倾斜Y": ("IsSurfaceTiltYUsed", None),
    "元件偏心X": ("IsElementDecenterXUsed", "ElementDecenterX"),
    "元件偏心Y": ("IsElementDecenterYUsed", "ElementDecenterY"),
    "元件倾斜X": ("IsElementTiltXUsed", "ElementTiltXDegrees"),
    "元件倾斜Y": ("IsElementTiltYUsed", "ElementTiltYDegrees"),
    "面不规则": ("IsSurfaceSandAIrregularityUsed", "SurfaceSandAIrregularityFringes"),
    "折射率": ("IsIndexUsed", "Index"),
    "阿贝%": ("IsIndexAbbePercentageUsed", "IndexAbbePercentage"),
}

# 向导支持的全部开关（用于先全部关闭、再按 Excel 打开）
_ALL_SWITCHES = [
    "IsSurfaceRadiusUsed", "IsSurfaceThicknessUsed",
    "IsSurfaceDecenterXUsed", "IsSurfaceDecenterYUsed",
    "IsSurfaceTiltXUsed", "IsSurfaceTiltYUsed",
    "IsElementDecenterXUsed", "IsElementDecenterYUsed",
    "IsElementTiltXUsed", "IsElementTiltYUsed",
    "IsSurfaceSandAIrregularityUsed", "IsSurfaceZernikeIrregularityUsed",
    "IsIndexUsed", "IsIndexAbbePercentageUsed",
]


def _wizard_range(wizard_rows: list[dict]) -> tuple[int, int]:
    """取各启用行 起始/结束面 的最小/最大作为全局统一范围。"""
    starts, stops = [], []
    for row in wizard_rows:
        if not _is_yes(row.get("启用")):
            continue
        s0 = _num(row.get("起始面"))
        s1 = _num(row.get("结束面"))
        if s0 is not None:
            starts.append(int(s0))
        if s1 is not None:
            stops.append(int(s1))
    if not starts or not stops:
        return 0, 0
    return min(starts), max(stops)


def run_native_wizard(zos_system, wizard_rows: list[dict],
                      test_wavelength_um: float = 0.0,
                      focus_compensation: bool = False) -> int:
    """按 Excel 公差向导调用 SEQToleranceWizard 生成 TDE。返回 TDE 行数。

    全局统一范围 = 各启用行起始/结束面的最小/最大。
    原生向导无“跳过面”概念，由 Zemax 自动识别玻璃/空气/光阑。
    """
    import ZOSAPI

    wiz = zos_system.TDE.SEQToleranceWizard
    wiz.Initialize()

    for sw in _ALL_SWITCHES:
        try:
            setattr(wiz, sw, False)
        except Exception:
            pass

    enabled = {}
    for row in wizard_rows:
        if not _is_yes(row.get("启用")):
            continue
        cat = str(row.get("公差类别", "")).strip()
        if cat not in _CATEGORY_MAP:
            raise ValueError(f"未知公差类别: {cat!r}（向导）")
        enabled[cat] = _num(row.get("数值"))

    DF = ZOSAPI.Wizards.DefaultAndFringes
    DD = ZOSAPI.Wizards.DefaultAndDegrees

    for cat, val in enabled.items():
        switch_attr, value_attr = _CATEGORY_MAP[cat]
        setattr(wiz, switch_attr, True)
        if val is None:
            continue
        if cat == "半径":
            wiz.SurfaceRadiusUnitType = DF.Fringes
            wiz.SurfaceRadiusFringes = float(val)
        elif cat == "面倾斜X":
            wiz.SurfaceTiltXUnitType = DD.Degrees
            wiz.SurfaceTiltXDegrees = float(val)
        elif cat == "面倾斜Y":
            wiz.SurfaceTiltYUnitType = DD.Degrees
            wiz.SurfaceTiltYDegrees = float(val)
        else:
            setattr(wiz, value_attr, float(val))

    s0, s1 = _wizard_range(wizard_rows)
    if s0 and s1:
        wiz.StartAtSurface = int(s0)
        wiz.StopAtSurface = int(s1)

    if test_wavelength_um and test_wavelength_um > 0:
        wiz.TestWavelength = float(test_wavelength_um)
    wiz.IsFocusCompensationUsed = bool(focus_compensation)

    wiz.OK()
    return zos_system.TDE.NumberOfOperands


# ---------------------------------------------------------------------------
# 公差明细 sheet → 在已生成的 TDE 上覆盖/追加/删除
# ---------------------------------------------------------------------------

def _find_tde_row(tde, op: str, s1: int, s2: int):
    """在 TDE 中按 (类型, Param1, Param2) 查找匹配行，返回行号或 0。"""
    import ZOSAPI
    T = ZOSAPI.Editors.TDE.ToleranceOperandType
    want = getattr(T, op, None)
    for i in range(1, tde.NumberOfOperands + 1):
        r = tde.GetOperandAt(i)
        try:
            same_type = (r.Type == want)
        except Exception:
            same_type = (str(r.Type) == op)
        if not same_type:
            continue
        p1 = int(r.Param1)
        p2 = int(r.Param2) if op in PAIRED_OPS else 0
        if p1 == s1 and (op not in PAIRED_OPS or p2 == s2):
            return i
    return 0


def apply_detail_to_tde(zos_system, detail_rows: list[dict]) -> int:
    """按明细行在已生成的 TDE 上覆盖/追加/删除。返回净变更条数。"""
    import ZOSAPI

    tde = zos_system.TDE
    T = ZOSAPI.Editors.TDE.ToleranceOperandType
    changed = 0

    for row in detail_rows:
        action = str(row.get("操作", "")).strip()
        op = str(row.get("操作数", "")).strip().upper()
        if not op:
            continue
        op_enum = getattr(T, op, None)
        if op_enum is None:
            raise ValueError(f"TDE 不支持的操作数: {op}（明细）")
        s1 = int(_num(row.get("面1"), 0))
        s2 = int(_num(row.get("面2"), 0))
        if op not in PAIRED_OPS:
            s2 = 0

        if action in ("删除", "delete", "DELETE"):
            idx = _find_tde_row(tde, op, s1, s2)
            if idx:
                tde.RemoveOperandAt(idx)
                changed += 1
            continue

        vmin = _num(row.get("Min"), 0.0)
        vmax = _num(row.get("Max"), 0.0)
        comment = str(row.get("注释") or "")

        if action in ("覆盖", "override", "OVERRIDE", "覆盖向导"):
            idx = _find_tde_row(tde, op, s1, s2)
            r = tde.GetOperandAt(idx) if idx else tde.AddOperand()
        elif action in ("追加", "append", "APPEND", "新增"):
            r = tde.AddOperand()
        else:
            raise ValueError(f"未知明细操作: {action!r}（应为 追加/覆盖/删除）")

        if not r.ChangeType(op_enum):
            raise RuntimeError(f"明细 ChangeType 失败: {op}")
        r.Param1 = int(s1)
        if op in PAIRED_OPS:
            r.Param2 = int(s2)
        r.Min = float(vmin)
        r.Max = float(vmax)
        if comment:
            r.Comment = comment
        changed += 1

    return changed


def build_and_write(zos_system, wizard_rows: list[dict],
                    detail_rows: list[dict], center_wave: int = 0,
                    test_wavelength_um: float = 0.0,
                    focus_compensation: bool = False,
                    comp_surface: int = 0,
                    comp_min=None, comp_max=None) -> int:
    """完整流程：原生向导生成 → 明细覆盖 → COMP 补偿器 → 返回 TDE 总行数。

    顺序铁律：向导 → 明细 → COMP 最后追加（否则被向导清掉）。
    comp_surface>0 才追加 COMP；留空=0=跳过（非 bug，需在 Excel 填面号）。

    center_wave 保留兼容旧签名（原生向导用 TestWavelength 设波长，
    若只给了 center_wave 而无 test_wavelength_um，则忽略 center_wave，
    由向导默认/调用方传 test_wavelength_um 控制）。
    """
    run_native_wizard(zos_system, wizard_rows,
                      test_wavelength_um=test_wavelength_um,
                      focus_compensation=focus_compensation)
    if detail_rows:
        apply_detail_to_tde(zos_system, detail_rows)
    if comp_surface and int(comp_surface) > 0:
        add_back_focus_compensator(zos_system, int(comp_surface),
                                   comp_min, comp_max)
    return zos_system.TDE.NumberOfOperands


def add_back_focus_compensator(zos_system, comp_surface: int,
                               vmin=None, vmax=None,
                               comment: str = "后焦补偿") -> int:
    """在 TDE 末尾追加一个 COMP 补偿器（后焦：该面厚度作补偿器）。

    官方语义：COMP 的 Param1=面号，Param2=0 表示对该面厚度做补偿。
    必须在向导与明细之后追加，否则会被向导清掉。
    vmin/vmax 为 None 时留 0（由 Zemax 自由调整）。返回新行号。
    """
    import ZOSAPI
    T = ZOSAPI.Editors.TDE.ToleranceOperandType

    tde = zos_system.TDE
    r = tde.AddOperand()
    if not r.ChangeType(T.COMP):
        raise RuntimeError("COMP ChangeType 失败")
    r.Param1 = int(comp_surface)
    r.Param2 = 0
    r.Min = 0.0 if vmin is None else float(vmin)
    r.Max = 0.0 if vmax is None else float(vmax)
    if comment:
        r.Comment = comment
    return tde.NumberOfOperands
