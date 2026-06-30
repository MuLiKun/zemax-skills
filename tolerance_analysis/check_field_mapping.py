from __future__ import annotations

import argparse
import math
import os
import sys
from dataclasses import dataclass

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

from toltool import zos_connect


DEFAULT_TARGETS = "0,-0.25,0.25,-0.5,0.5,-0.7,0.7,-0.9,0.9,-1,1"


@dataclass
class FieldRow:
    field_no: int
    x: float
    y: float
    field_abs: float
    normalized: float


def _parse_targets(text: str) -> list[float]:
    targets: list[float] = []
    for part in str(text or "").split(","):
        part = part.strip()
        if not part:
            continue
        targets.append(float(part))
    return targets or _parse_targets(DEFAULT_TARGETS)


def _sort_key_normalized(value: float) -> tuple[float, int, float]:
    if abs(value) < 1e-12:
        return (0.0, 0, 0.0)
    return (abs(value), 0 if value < 0 else 1, value)


def _normalize_value(x: float, y: float, max_abs_x: float, max_abs_y: float) -> float:
    if max_abs_y > 0 and abs(y) >= abs(x):
        return y / max_abs_y
    if max_abs_x > 0:
        return x / max_abs_x
    return 0.0


def build_mapping(fields: list[tuple[float, float]]) -> list[FieldRow]:
    raw = []
    for i, (x, y) in enumerate(fields, start=1):
        raw.append((i, float(x), float(y), math.hypot(float(x), float(y))))
    max_abs_x = max((abs(r[1]) for r in raw), default=0.0)
    max_abs_y = max((abs(r[2]) for r in raw), default=0.0)
    rows = [
        FieldRow(
            field_no=i,
            x=x,
            y=y,
            field_abs=field_abs,
            normalized=_normalize_value(x, y, max_abs_x, max_abs_y),
        )
        for i, x, y, field_abs in raw
    ]
    return sorted(rows, key=lambda r: (_sort_key_normalized(r.normalized), r.field_no))


def _nearest(rows: list[FieldRow], target: float) -> FieldRow | None:
    if not rows:
        return None
    return min(rows, key=lambda r: (abs(r.normalized - target), r.field_no))


def _suggest_insert_xy(rows: list[FieldRow], target: float) -> tuple[float, float] | None:
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


def _report_label(target: float) -> str:
    if abs(target) < 1e-12:
        return "F0"
    text = f"{target:g}"
    return f"F{text}"


def print_mapping(rows: list[FieldRow], targets: list[float], threshold: float) -> None:
    print("\n视场号 ↔ 归一化视场映射（排序：0、负、正，按归一化视场绝对值递增）")
    print("视场号        X              Y        视场绝对值      归一化视场")
    for r in rows:
        print(f"{r.field_no:>6}  {r.x:>12.6g}  {r.y:>12.6g}  {r.field_abs:>12.6g}  {r.normalized:>12.6f}")

    print("\n目标归一化视场推荐（评价函数与 REPORT 使用同一映射）")
    print("目标归一化视场  推荐视场号  实际归一化视场        偏差    REPORT标签    建议")
    need_insert: list[tuple[float, float, float, float, str]] = []
    for target in sorted(targets, key=_sort_key_normalized):
        label = _report_label(target)
        hit = _nearest(rows, target)
        if hit is None:
            print(f"{target:>14.6f}          -               -           -    {label:<10}  无视场数据")
            continue
        delta = abs(hit.normalized - target)
        suggestion = "OK"
        insert_xy = _suggest_insert_xy(rows, target)
        if delta > threshold and insert_xy is not None:
            suggestion = "建议插入"
            need_insert.append((target, insert_xy[0], insert_xy[1], delta, label))
        print(f"{target:>14.6f}  {hit.field_no:>10}  {hit.normalized:>14.6f}  {delta:>10.6f}    {label:<10}  {suggestion}")

    if need_insert:
        print(f"\n以下目标与最近已有视场差距大于阈值 {threshold:g}，可考虑插入到工作副本视场表：")
        print("目标归一化视场      建议X          建议Y        偏差    REPORT标签")
        for target, x, y, delta, label in need_insert:
            print(f"{target:>14.6f}  {x:>12.6g}  {y:>12.6g}  {delta:>10.6f}    {label}")
        print("\n说明：当前脚本只给出插入建议，不会修改 zmx。正式功能应只修改工作副本，并要求用户确认。")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="读取 zmx 视场并输出视场号与归一化视场映射")
    parser.add_argument("--zmx", required=True, help="待读取的 zmx/zos 文件路径")
    parser.add_argument("--connect", choices=["standalone", "extension"], default="standalone")
    parser.add_argument("--targets", default=DEFAULT_TARGETS, help="目标归一化视场，逗号分隔")
    parser.add_argument("--threshold", type=float, default=0.05, help="大于该归一化视场偏差时提示建议插入")
    parser.add_argument("--zos-dir", default=None, help="Zemax OpticStudio 安装目录，可选")
    args = parser.parse_args(argv)

    if not os.path.isfile(args.zmx):
        raise FileNotFoundError(args.zmx)

    sess = zos_connect.ZosSession(zos_dir=args.zos_dir)
    try:
        sess.connect(mode=args.connect)
        copy = sess.open_as_copy(args.zmx)
        info = sess.read_lens_info()
        print(f"工作副本: {copy}")
        print(f"视场类型: {info.field_type}")
        print(f"视场数量: {len(info.fields)}")
        rows = build_mapping(info.fields)
        print_mapping(rows, _parse_targets(args.targets), args.threshold)
        return 0
    finally:
        if args.connect == "standalone":
            sess.close()


if __name__ == "__main__":
    raise SystemExit(main())
