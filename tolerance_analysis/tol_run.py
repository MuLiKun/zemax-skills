"""tol_run.py —— 公差分析自动化主入口（命令行，含 ZTD 读取）。

用法见 公差分析程序_需求文档.md §6。本阶段已实现：
  --init-template  生成配置模板
  --read-only      连接 Zemax，读取并打印镜头信息（不跑分析）
  (默认)           读配置 → 建 TDE/MFE/TSC → 跑蒙卡 → 读 ZTD 出统计

准备与跑蒙卡的逻辑统一由 toltool.pipeline 提供（与方案A/方案B 共享
单一来源），本文件在其上补充 ZTD 读取与命令行参数解析。
"""

from __future__ import annotations

import argparse
import sys

from toltool import excel_io


def _cmd_init_template(args) -> int:
    path = args.config
    if not path:
        print("错误：--init-template 需要 --config 指定输出路径。", file=sys.stderr)
        return 2
    try:
        out = excel_io.generate_template(path, overwrite=args.overwrite)
    except FileExistsError as e:
        print(f"{e}（加 --overwrite 可覆盖）", file=sys.stderr)
        return 1
    print(f"已生成配置模板: {out}")
    return 0


def _cmd_read_only(args) -> int:
    from toltool import zos_connect

    sess = zos_connect.ZosSession()
    print(f"ZOS 目录: {sess.zos_dir}")
    sess.connect(mode=args.connect)
    print(f"已连接: {sess.sys.SystemFile}")
    copy = sess.open_as_copy(args.zmx) if args.zmx else None
    if copy:
        print(f"工作副本: {copy}")
    info = sess.read_lens_info()
    print(f"\n面数={info.num_surfaces} 光阑面={info.stop_surface} "
          f"孔径={info.aperture_type}:{info.aperture_value}")
    print(f"波长(um)={info.wavelengths_um} 主波长号={info.primary_wave}")
    print(f"视场({info.field_type})={info.fields}")
    print("\n面号  注释            半径         厚度         材料")
    for s in info.surfaces:
        print(f"{s.index:>3}  {s.comment[:12]:<12}  "
              f"{s.radius:>12.5g}  {s.thickness:>10.5g}  {s.material}")
    return 0


def _yes(v) -> bool:
    return str(v).strip().upper() in ("Y", "YES", "1", "TRUE", "是")


def _as_int(v, default: int) -> int:
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return default


def _cmd_run(args) -> int:
    if not args.config:
        print("错误：运行需要 --config 指定配置 Excel。", file=sys.stderr)
        return 2
    if not args.zmx:
        print("错误：运行需要 --zmx 指定待分析镜头。", file=sys.stderr)
        return 2

    from toltool import pipeline, ztd_reader

    prep = pipeline.prepare_session(
        args.zmx, args.config, outdir=args.outdir, connect=args.connect)

    result = pipeline.run_montecarlo(prep)
    if not result.succeeded:
        print(f"公差分析失败: {result.message}", file=sys.stderr)
        return 1
    print(f"分析完成。ZTD: {result.ztd_path}")
    if result.bestworst_folder:
        print(f"Worst/Best 输出目录: {result.bestworst_folder}")

    if args.no_read:
        print("（--no-read：跳过 ZTD 读取，请在 OpticStudio 中查看结果）")
        return 0

    report_meta = [
        r for r in prep.cfg.report if _yes(r.get("启用")) and r.get("标签")
    ]
    report_labels = [str(r.get("标签")).strip() for r in report_meta]
    num_runs = _as_int(prep.rp.get("蒙特卡洛次数"), 200)
    zres = ztd_reader.read_ztd(
        prep.sess.sys, result.ztd_path, num_runs=num_runs,
        report_labels=report_labels or None,
        report_meta=report_meta or None)
    print()
    print(ztd_reader.format_table(zres))
    if zres.succeeded and _yes(prep.rp.get("输出统计Excel", "N")):
        stat_path = result.ztd_path.rsplit(".", 1)[0] + "_统计.xlsx"
        out = ztd_reader.export_excel(zres, stat_path)
        print(f"统计 Excel: {out}")
    return 0 if zres.succeeded else 1


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Zemax 公差分析自动化工具")
    p.add_argument("--zmx", help="待分析的 zmx 文件路径")
    p.add_argument("--config", help="Excel 配置文件路径")
    p.add_argument("--outdir", help="输出目录，默认 = zmx 同级目录")
    p.add_argument("--connect", choices=["extension", "standalone"],
                   default="extension", help="连接方式，默认 extension")
    p.add_argument("--init-template", action="store_true",
                   help="生成带示例的空白配置模板后退出")
    p.add_argument("--overwrite", action="store_true",
                   help="生成模板时覆盖已存在文件")
    p.add_argument("--read-only", action="store_true",
                   help="只读镜头信息，不跑分析")
    p.add_argument("--no-read", action="store_true",
                   help="跑完蒙特卡洛即停，不读取 ZTD（在 OpticStudio 中查看）")
    return p


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    if args.init_template:
        return _cmd_init_template(args)
    if args.read_only:
        return _cmd_read_only(args)
    return _cmd_run(args)


if __name__ == "__main__":
    raise SystemExit(main())
