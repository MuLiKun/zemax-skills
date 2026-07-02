"""tol_run.py —— 公差分析自动化主入口（命令行，含 ZTD 读取）。

用法见 公差分析程序_需求文档.md §6。本阶段已实现：
  --init-template  生成配置模板
  --read-only      连接 Zemax，读取并打印镜头信息（不跑分析）
  (默认)           读配置 → 建 TDE/MFE/TSC → 跑蒙卡 → 读 ZTD 出统计

准备与跑蒙卡的逻辑统一由 toltool.pipeline 提供，本文件在其上补充 ZTD 读取与命令行参数解析。
"""

from __future__ import annotations

import argparse
import os
import sys

from toltool import current_settings, excel_io, standard_templates


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


def _standard_config_path(args) -> str:
    return standard_templates.make_temp_config(
        args.zmx, args.outdir, args.standard_template, args.tolerance_level,
        args.num_runs, args.num_to_save, args.center_wave, args.comp_mode,
        args.save_worst_best, product_type=args.product_type)


def _cmd_validate_only(args) -> int:
    if not args.zmx:
        print("错误：--validate-only 需要 --zmx 指定待分析镜头。", file=sys.stderr)
        return 2
    if args.current_settings:
        if not os.path.isfile(args.zmx):
            print(f"配置校验失败：运行前校验失败：\n- ZMX 文件不存在：{args.zmx}", file=sys.stderr)
            return 1
        print("当前设置模式 validate-only：已校验 zmx 路径；MFE/TDE 需连接 Zemax 后运行期读取。")
        print(f"REPORT 筛选: {current_settings.report_filter_label(args.current_report_filter)}")
        return 0
    if not args.config and not args.standard:
        print("错误：--validate-only 需要 --config，或使用 --standard / --current-settings。", file=sys.stderr)
        return 2

    from toltool import pipeline

    try:
        config = _standard_config_path(args) if args.standard else args.config
        cfg = pipeline.validate_config(args.zmx, config)
    except Exception as e:
        print(f"配置校验失败：{e}", file=sys.stderr)
        return 1

    print("配置校验通过。")
    if args.standard:
        print(f"标准模板配置: {config}")
    print(f"评价函数有效行数: {len([r for r in cfg.mfe if str(r.get('操作数') or '').strip()])}")
    print(f"REPORT 启用行数: {len([r for r in cfg.report if _yes(r.get('启用'))])}")
    return 0


def _cmd_run(args) -> int:
    if not args.zmx:
        print("错误：运行需要 --zmx 指定待分析镜头。", file=sys.stderr)
        return 2
    if not args.config and not args.standard and not args.current_settings:
        print("错误：运行需要 --config，或使用 --standard / --current-settings。", file=sys.stderr)
        return 2

    from toltool import pipeline, ztd_reader

    config = _standard_config_path(args) if args.standard else args.config
    if args.standard:
        print(f"标准模板配置: {config}")
    if args.current_settings:
        print("当前设置模式：将复用工作副本中的现有 TDE/MFE。")
        print(f"REPORT 筛选: {current_settings.report_filter_label(args.current_report_filter)}")

    prep = pipeline.prepare_session(
        args.zmx, config, outdir=args.outdir, connect=args.connect,
        use_current_settings=args.current_settings,
        current_args={
            "num_runs": args.num_runs,
            "num_to_save": args.num_to_save,
            "comp_mode": args.comp_mode,
            "save_worst_best": args.save_worst_best,
            "report_filter": args.current_report_filter,
        })

    export_stats = (not args.no_read) and _yes(prep.rp.get("输出统计Excel", "N"))
    result = pipeline.run_montecarlo(
        prep, log=lambda m: pipeline.append_run_log(prep, m),
        export_stats=export_stats)
    if not result.succeeded:
        pipeline.append_run_log(prep, "公差分析失败：" + (result.message or "未知错误"))
        print(f"公差分析失败: {result.message}", file=sys.stderr)
        return 1
    pipeline.append_run_log(prep, f"分析完成。ZTD: {result.ztd_path}")
    if result.bestworst_folder:
        pipeline.append_run_log(prep, f"Worst/Best 输出目录: {result.bestworst_folder}")

    if args.no_read:
        print("（--no-read：跳过 ZTD 读取，请在 OpticStudio 中查看结果）")
        return 0

    try:
        report_meta = [
            r for r in prep.cfg.report if _yes(r.get("启用")) and r.get("标签")
        ]
        report_labels = [str(r.get("标签")).strip() for r in report_meta]
        num_items = len(report_labels) + 1 if report_labels else None
        comp_count = sum(
            1 for r in (prep.tde_meta or [])
            if str(r.get("操作数") or "").strip().upper() == "COMP")
        if num_items:
            pipeline.append_run_log(
                prep, f"ZTD 自动统计分项: {num_items + comp_count} 项（自定义脚本 + {len(report_labels)} 个 REPORT + {comp_count} 个 COMP）")
        num_runs = _as_int(prep.rp.get("蒙特卡洛次数"), 200)
        zres = ztd_reader.read_ztd(
            prep.sess.sys, result.ztd_path, num_runs=num_runs,
            report_labels=report_labels or None,
            num_items=num_items,
            report_meta=report_meta or None,
            tde_meta=prep.tde_meta or None)
        print()
        print(ztd_reader.format_table(zres))
        if not zres.succeeded:
            pipeline.append_run_log(prep, "分析完成，但读取 ZTD 失败：" + zres.message)
            return 1
        if zres.message:
            pipeline.append_run_log(prep, "提示：" + zres.message)
        if _yes(prep.rp.get("输出统计Excel", "N")):
            stat_path = result.ztd_path.rsplit(".", 1)[0] + "_统计.xlsx"
            out = ztd_reader.export_excel(zres, stat_path)
            pipeline.append_run_log(prep, f"统计 Excel: {out}")
        return 0
    except Exception as e:
        message = f"读取/导出 ZTD 失败：{type(e).__name__}: {e}"
        pipeline.append_run_log(prep, message)
        print(message, file=sys.stderr)
        return 1


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
    p.add_argument("--validate-only", action="store_true",
                   help="只校验 zmx 路径与 Excel 配置，不连接 Zemax、不跑分析")
    p.add_argument("--no-read", action="store_true",
                   help="跑完蒙特卡洛即停，不读取 ZTD（在 OpticStudio 中查看）")
    p.add_argument("--standard", action="store_true",
                   help="使用普通标准模板模式，不需要 --config")
    p.add_argument("--current-settings", action="store_true",
                   help="使用 Zemax 当前设置模式：复用 zmx 中已有 TDE/MFE，不需要 --config")
    p.add_argument("--product-type", choices=standard_templates.PRODUCT_TYPES,
                   default=standard_templates.DEFAULT_PRODUCT_TYPE,
                   help="产品类型；当前 TX 暂复用 RX 模板")
    p.add_argument("--standard-template", choices=standard_templates.TEMPLATE_NAMES,
                   default=standard_templates.DEFAULT_TEMPLATE_NAME, help="标准模板名称")
    p.add_argument("--tolerance-level", choices=standard_templates.LEVEL_NAMES,
                   default="标准", help="标准模板公差等级")
    p.add_argument("--num-runs", type=int, default=20,
                   help="标准模板模式的蒙特卡洛次数")
    p.add_argument("--num-to-save", type=int, default=0,
                   help="标准模板模式的 MC case 保存数量")
    p.add_argument("--center-wave", type=int, default=0,
                   help="标准模板模式的中心波长号，0=自动使用主波长")
    p.add_argument("--comp-mode", default="无",
                   help="标准模板/当前设置模式的补偿器模式")
    p.add_argument("--save-worst-best", action="store_true",
                   help="标准模板/当前设置模式下保存 Zemax Worst/Best case")
    p.add_argument("--current-report-filter", choices=["all", "mtf", "common"],
                   default="all", help="当前设置模式 REPORT 筛选：all=全部有效行，mtf=仅 MTF 类，common=常用评价类")
    return p


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    if args.standard and args.current_settings:
        print("错误：--standard 与 --current-settings 不能同时使用。", file=sys.stderr)
        return 2
    if args.init_template:
        return _cmd_init_template(args)
    if args.read_only:
        return _cmd_read_only(args)
    if args.validate_only:
        return _cmd_validate_only(args)
    return _cmd_run(args)


if __name__ == "__main__":
    raise SystemExit(main())
