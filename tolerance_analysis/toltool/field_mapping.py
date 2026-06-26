"""field_mapping.py —— 高级 Excel 模式的视场号/归一化视场映射。

该模块只在运行参数显式启用时介入主流程：
- 读取 tol 工作副本视场；
- 按目标归一化视场匹配 Zemax 视场号；
- 可选把缺失目标视场插入到工作副本；
- 改写内存中的 MFE/REPORT 配置，不回写原始 Excel。
"""

from __future__ import annotations

import copy
import math
from dataclasses import asdict, dataclass


DEFAULT_TARGETS = "0,-0.25,0.25,-0.5,0.5,-0.7,0.7,-0.9,0.9,-1,1"
_FIELD_OPS = {"GENC", "GMTT", "GMTS", "GMTA"}


@dataclass
class FieldItem:
    field_no: int
    x: float
    y: float
    field_abs: float
    normalized: float


@dataclass
class FieldMatch:
    target_normalized: float
    field_no: int | None
    actual_normalized: float | None
    delta: float | None
    need_insert: bool
    suggested_x: float | None
    suggested_y: float | None
    report_label: str


@dataclass
class FieldMappingResult:
    enabled: bool
    insert_strategy: str
    threshold: float
    targets: list[float]
    original_fields: list[FieldItem]
    inserted_fields: list[FieldMatch]
    final_fields: list[FieldItem]
    final_matches: list[FieldMatch]
    messages: list[str]

    def to_dict(self) -> dict:
        data = asdict(self)
        return data


def yes(value) -> bool:
    return str(value).strip().upper() in ("Y", "YES", "1", "TRUE", "是")


def parse_targets(value) -> list[float]:
    text = str(value or DEFAULT_TARGETS)
    targets: list[float] = []
    for part in text.replace("；", ",").replace("，", ",").split(","):
        part = part.strip()
        if not part:
            continue
        targets.append(float(part))
    return sorted(targets or parse_targets(DEFAULT_TARGETS), key=sort_key)


def sort_key(value: float) -> tuple[float, int, float]:
    if abs(value) < 1e-12:
        return (0.0, 0, 0.0)
    return (abs(value), 0 if value < 0 else 1, value)


def report_label(target: float) -> str:
    if abs(target) < 1e-12:
        return "F0"
    return f"F{target:g}"


def _read_fields(zos_system) -> list[tuple[float, float]]:
    fields = zos_system.SystemData.Fields
    return [
        (float(fields.GetField(i).X), float(fields.GetField(i).Y))
        for i in range(1, fields.NumberOfFields + 1)
    ]


def _normalize_value(x: float, y: float, max_abs_y: float, max_row: tuple[float, float, float] | None) -> float:
    if max_abs_y > 0 and abs(y) >= abs(x):
        return y / max_abs_y
    if max_row is None or max_row[2] <= 0:
        return 0.0
    max_x, max_y, max_abs = max_row
    return (x * max_x + y * max_y) / (max_abs * max_abs)


def build_field_items(fields_xy: list[tuple[float, float]]) -> list[FieldItem]:
    raw = []
    for i, (x, y) in enumerate(fields_xy, start=1):
        x = float(x)
        y = float(y)
        raw.append((i, x, y, math.hypot(x, y)))
    max_abs_y = max((abs(r[2]) for r in raw), default=0.0)
    max_row = max(((r[1], r[2], r[3]) for r in raw), key=lambda r: r[2], default=None)
    rows = [
        FieldItem(
            field_no=i,
            x=x,
            y=y,
            field_abs=field_abs,
            normalized=_normalize_value(x, y, max_abs_y, max_row),
        )
        for i, x, y, field_abs in raw
    ]
    return sorted(rows, key=lambda r: (sort_key(r.normalized), r.field_no))


def _nearest(rows: list[FieldItem], target: float) -> FieldItem | None:
    if not rows:
        return None
    return min(rows, key=lambda r: (abs(r.normalized - target), r.field_no))


def _suggest_insert_xy(rows: list[FieldItem], target: float) -> tuple[float, float] | None:
    if not rows:
        return None
    max_abs_x = max(abs(r.x) for r in rows)
    max_abs_y = max(abs(r.y) for r in rows)
    if max_abs_y > 0 and max_abs_y >= max_abs_x:
        return 0.0, target * max_abs_y
    if max_abs_x > 0 and max_abs_x > max_abs_y:
        return target * max_abs_x, 0.0
    edge = max(rows, key=lambda r: (r.field_abs, r.field_no))
    return edge.x * target, edge.y * target


def build_matches(rows: list[FieldItem], targets: list[float], threshold: float) -> list[FieldMatch]:
    matches: list[FieldMatch] = []
    for target in sorted(targets, key=sort_key):
        hit = _nearest(rows, target)
        sx_sy = _suggest_insert_xy(rows, target)
        if hit is None:
            matches.append(FieldMatch(target, None, None, None, True,
                                      sx_sy[0] if sx_sy else None,
                                      sx_sy[1] if sx_sy else None,
                                      report_label(target)))
            continue
        delta = abs(hit.normalized - target)
        need_insert = delta > threshold
        matches.append(FieldMatch(
            target_normalized=target,
            field_no=hit.field_no,
            actual_normalized=hit.normalized,
            delta=delta,
            need_insert=need_insert,
            suggested_x=sx_sy[0] if sx_sy else None,
            suggested_y=sx_sy[1] if sx_sy else None,
            report_label=report_label(target),
        ))
    return matches


def _add_field(zos_system, x: float, y: float) -> None:
    fields = zos_system.SystemData.Fields
    try:
        fields.AddField(float(x), float(y), 1.0)
    except TypeError:
        fields.AddField(float(x), float(y))


def _find_match(matches: list[FieldMatch], target: float) -> FieldMatch | None:
    for item in matches:
        if abs(item.target_normalized - target) < 1e-9:
            return item
    return None


def _num(value):
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _apply_to_mfe_rows(mfe_rows: list[dict], matches: list[FieldMatch]) -> list[dict]:
    new_rows = copy.deepcopy(mfe_rows)
    for row in new_rows:
        op = str(row.get("操作数") or "").strip().upper()
        target = _num(row.get("归一化视场"))
        if target is None and op == "RSCE":
            target = _num(row.get("Param4"))
        if target is None:
            continue
        match = _find_match(matches, target)
        if match is None or match.field_no is None:
            continue
        if op in _FIELD_OPS:
            row["Param3"] = match.field_no
        elif op == "RSCE":
            row["Param4"] = target
    return new_rows


def _base_report_name(label: str) -> str:
    text = str(label or "").strip()
    if not text:
        return "REPORT"
    for marker in ("_F", "-F"):
        idx = text.upper().rfind(marker)
        if idx > 0:
            return text[:idx]
    return text


def _label_with_field(base: str, field_label: str) -> str:
    base = str(base or "REPORT").strip("_-") or "REPORT"
    return f"{base}_{field_label}"


def _apply_to_report_rows(report_rows: list[dict], mfe_rows: list[dict], matches: list[FieldMatch]) -> list[dict]:
    new_rows = copy.deepcopy(report_rows)
    by_line: dict[int, dict] = {}
    for row in mfe_rows:
        line = row.get("行号")
        try:
            by_line[int(float(line))] = row
        except (TypeError, ValueError):
            continue
    for row in new_rows:
        try:
            mf_line = int(float(row.get("MF行号")))
        except (TypeError, ValueError):
            continue
        mfe_row = by_line.get(mf_line)
        if not mfe_row:
            continue
        target = _num(mfe_row.get("归一化视场"))
        if target is None and str(mfe_row.get("操作数") or "").strip().upper() == "RSCE":
            target = _num(mfe_row.get("Param4"))
        if target is None:
            continue
        match = _find_match(matches, target)
        if match is None:
            continue
        row["标签"] = _label_with_field(_base_report_name(row.get("标签")), match.report_label)
    return new_rows


def process(zos_system, cfg, run_params: dict, log=print) -> tuple[object, FieldMappingResult]:
    enabled = yes(run_params.get("启用视场映射", "N"))
    strategy = str(run_params.get("视场插入策略") or "禁用").strip()
    messages: list[str] = []
    if not enabled:
        result = FieldMappingResult(False, strategy, 0.05, [], [], [], [], [], messages)
        return cfg, result

    threshold = float(run_params.get("视场匹配阈值") or 0.05)
    targets = parse_targets(run_params.get("目标归一化视场") or DEFAULT_TARGETS)
    original = build_field_items(_read_fields(zos_system))
    initial_matches = build_matches(original, targets, threshold)
    inserted: list[FieldMatch] = []

    auto_insert = strategy.replace(" ", "").lower() in ("自动插入", "auto", "insert", "y", "yes")
    if auto_insert:
        for item in initial_matches:
            if not item.need_insert:
                continue
            if item.suggested_x is None or item.suggested_y is None:
                raise RuntimeError(f"目标归一化视场 {item.target_normalized:g} 无法计算插入视场")
            _add_field(zos_system, item.suggested_x, item.suggested_y)
            inserted.append(item)
        if inserted:
            zos_system.Save()
            messages.append(f"已插入 {len(inserted)} 个缺失视场到 tol 工作副本")
    else:
        missing = [m for m in initial_matches if m.need_insert]
        if missing:
            labels = ", ".join(m.report_label for m in missing)
            raise RuntimeError(
                f"启用视场映射后存在目标视场偏差大于阈值但未插入：{labels}。"
                f"请将视场插入策略设为自动插入，或关闭视场映射后手动确认 Excel。")

    final_fields = build_field_items(_read_fields(zos_system))
    final_matches = build_matches(final_fields, targets, threshold)
    missing_after = [m for m in final_matches if m.need_insert]
    if auto_insert and missing_after:
        labels = ", ".join(m.report_label for m in missing_after)
        raise RuntimeError(f"视场插入后仍有目标视场未满足阈值：{labels}")

    new_cfg = copy.deepcopy(cfg)
    new_cfg.mfe = _apply_to_mfe_rows(cfg.mfe, final_matches)
    new_cfg.report = _apply_to_report_rows(cfg.report, new_cfg.mfe, final_matches)

    result = FieldMappingResult(True, strategy, threshold, targets, original, inserted, final_fields, final_matches, messages)
    return new_cfg, result
