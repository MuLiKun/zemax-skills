"""pipeline.py —— 公差分析的可复用流程（两方案共享准备逻辑）。

把「连接 → 建 TDE → 建 MFE → 建 TSC → Save」这段两方案完全相同的
准备工作抽成 prepare_session()，避免方案A/方案B 两份入口重复维护：

  方案A（全自动）：prepare_session() + run_montecarlo()，API 直接跑完。
  方案B（手动）  ：只 prepare_session()，停下，由用户在 OpticStudio
                  公差分析对话框选 TSC、点 Run，观察原生窗口实时进度。
"""

from __future__ import annotations

import json
import os
import re
import shutil
from dataclasses import dataclass
from datetime import datetime

from . import (zos_connect, excel_io, tde_builder, mfe_builder,
               tsc_builder, tol_runner)


def _as_int(v, default: int) -> int:
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return default


def _yes(v) -> bool:
    return str(v).strip().upper() in ("Y", "YES", "1", "TRUE", "是")


def _num(v, default=None):
    if v is None or (isinstance(v, str) and str(v).strip() == ""):
        return default
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def _safe_name(name: str) -> str:
    text = re.sub(r'[<>:"/\\|?*\s]+', "_", str(name).strip())
    text = text.strip("._")
    return text or "lens"


def _make_run_dir(zmx: str, outdir: str | None) -> tuple[str, str]:
    src_base = os.path.splitext(os.path.basename(zmx))[0]
    parent = os.path.abspath(outdir) if outdir \
        else os.path.dirname(os.path.abspath(zmx))
    os.makedirs(parent, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_dir = os.path.join(parent, f"公差分析_{_safe_name(src_base)}_{stamp}")
    suffix = 1
    unique_dir = run_dir
    while os.path.exists(unique_dir):
        suffix += 1
        unique_dir = f"{run_dir}_{suffix}"
    os.makedirs(unique_dir, exist_ok=False)
    return parent, unique_dir


def _json_default(value):
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    return str(value)


def _write_json(path: str, data: dict) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=_json_default)


def _log_to_file(path: str, message: str) -> None:
    with open(path, "a", encoding="utf-8") as f:
        f.write(str(message) + "\n")


def _tee_logger(log, log_path: str):
    def emit(message: str) -> None:
        _log_to_file(log_path, message)
        log(message)
    return emit


def append_run_log(prep, message: str, log=print) -> None:
    if getattr(prep, "log_path", ""):
        _log_to_file(prep.log_path, message)
    log(message)


def _validate_paths(zmx: str, config: str) -> None:
    errors: list[str] = []
    if not os.path.isfile(zmx):
        errors.append(f"ZMX 文件不存在：{zmx}")
    if not os.path.isfile(config):
        errors.append(f"Excel 配置不存在：{config}")
    if errors:
        raise ValueError("运行前校验失败：\n" + "\n".join(f"- {e}" for e in errors))


def _validate_inputs(cfg, rp: dict) -> None:
    errors: list[str] = []
    num_runs = _as_int(rp.get("蒙特卡洛次数"), 200)
    num_to_save = _as_int(rp.get("保存数量"), 10)
    if num_runs <= 0:
        errors.append("蒙特卡洛次数必须大于 0")
    if num_to_save < 0:
        errors.append("保存数量不能小于 0")
    if num_to_save > num_runs:
        errors.append("保存数量不能大于蒙特卡洛次数")

    valid_mfe_lines: set[int] = set()
    for row in cfg.mfe:
        op = str(row.get("操作数") or "").strip()
        if not op:
            continue
        line = row.get("行号")
        if line in (None, ""):
            errors.append(f"评价函数操作数 {op} 缺少行号")
            continue
        try:
            line_no = int(float(line))
        except (TypeError, ValueError):
            errors.append(f"评价函数行号无效：{line!r}")
            continue
        if line_no <= 0:
            errors.append(f"评价函数行号必须大于 0：{line_no}")
            continue
        valid_mfe_lines.add(line_no)
    if not valid_mfe_lines:
        errors.append("评价函数工作表至少需要 1 行带操作数的有效行")

    report_count = 0
    for row in cfg.report:
        if not _yes(row.get("启用")):
            continue
        label = str(row.get("标签") or "").strip()
        mf_line = row.get("MF行号")
        if not label:
            errors.append("启用的 REPORT 行缺少标签")
            continue
        if mf_line in (None, ""):
            errors.append(f"REPORT {label} 缺少 MF行号")
            continue
        try:
            mf_line_no = int(float(mf_line))
        except (TypeError, ValueError):
            errors.append(f"REPORT {label} 的 MF行号无效：{mf_line!r}")
            continue
        if mf_line_no not in valid_mfe_lines:
            errors.append(f"REPORT {label} 的 MF行号 {mf_line_no} 未在评价函数中找到")
        report_count += 1
    if report_count == 0:
        errors.append("REPORT 至少需要启用 1 个带标签和 MF行号的分项")

    if errors:
        raise ValueError("运行前校验失败：\n" + "\n".join(f"- {e}" for e in errors))


def validate_config(zmx: str, config: str):
    """只校验文件路径和 Excel 配置，不连接 Zemax。"""
    _validate_paths(zmx, config)
    cfg = excel_io.read_config(config)
    _validate_inputs(cfg, cfg.run_params)
    return cfg


@dataclass
class Prepared:
    sess: object
    cfg: object
    rp: dict
    base: str
    outdir: str
    tsc_path: str
    mf_path: str
    n_tde: int
    n_mfe: int
    n_report: int
    lens_dir: str = ""
    parent_outdir: str = ""
    source_zmx: str = ""
    config_path: str = ""
    log_path: str = ""
    run_config_path: str = ""


def prepare_session(zmx: str, config: str, outdir: str | None = None,
                    connect: str = "extension", log=print,
                    zos_dir: str | None = None) -> Prepared:
    """连接 Zemax、打开工作副本，按 Excel 重建 TDE/MFE/TSC 并 Save。

    两方案共用。返回 Prepared，方案A 据此跑蒙卡，方案B 据此提示用户。
    """
    cfg = validate_config(zmx, config)
    rp = cfg.run_params

    parent_out, out = _make_run_dir(zmx, outdir)
    log_path = os.path.join(out, "run.log")
    log = _tee_logger(log, log_path)
    log(f"结果目录: {out}")
    log(f"日志文件: {log_path}")
    try:
        shutil.copy2(config, os.path.join(out, "used_excel.xlsx"))
    except Exception as e:
        log(f"保存 Excel 配置快照失败(忽略): {e}")

    center_wave = _as_int(rp.get("中心波长号"), 0)
    comp_surface = _as_int(rp.get("后焦补偿面"), 0)
    comp_min = _num(rp.get("补偿Min"))
    comp_max = _num(rp.get("补偿Max"))
    comp_freq = _num(rp.get("补偿线对"), 34.0)
    comp_mode = str(rp.get("补偿器模式") or "近轴焦点").strip()
    # 补偿器模式=无 → 完全不补偿：TSC 不写优化行/不走双 MF、不建 comp MF、不加 TDE COMP。
    # 模式≠无 → 写优化行 + comp MF（双 MF）；其中 TDE COMP 还需后焦补偿面>0 才落。
    comp_off = str(comp_mode).replace(" ", "").lower() in ("无", "none", "")
    comp_on = not comp_off
    add_comp_operand = comp_on and comp_surface > 0

    sess = zos_connect.ZosSession(zos_dir=zos_dir)
    log(f"ZOS 目录: {sess.zos_dir}")
    sess.connect(mode=connect)
    if connect == "standalone":
        log("已启动 Zemax 独立实例（即将载入下面的工作副本）")
    else:
        log(f"已连接交互扩展: {sess.sys.SystemFile}")

    src_base, src_ext = os.path.splitext(os.path.basename(zmx))
    copy_path = os.path.join(out, f"{src_base}_tol{src_ext}")

    copy = sess.open_as_copy(zmx, copy_path=copy_path)
    log(f"工作副本: {copy}")

    base = os.path.splitext(os.path.basename(copy))[0]
    lens_dir = os.path.dirname(os.path.abspath(copy))

    test_wl = 0.0
    if center_wave > 0:
        try:
            test_wl = float(sess.sys.SystemData.Wavelengths
                            .GetWavelength(center_wave).Wavelength)
        except Exception as e:
            log(f"读取测试波长失败(忽略): {e}")

    comp_surface_eff = comp_surface if add_comp_operand else 0
    n_tde = tde_builder.build_and_write(
        sess.sys, cfg.tol_wizard, cfg.tol_detail,
        center_wave=center_wave, test_wavelength_um=test_wl,
        comp_surface=comp_surface_eff, comp_min=comp_min, comp_max=comp_max)
    if add_comp_operand:
        log(f"已写入 TDE 公差: {n_tde} 条（含后焦补偿面 {comp_surface}，测试波长 {test_wl}um）")
    elif comp_off and comp_surface > 0:
        log(f"已写入 TDE 公差: {n_tde} 条（补偿器模式=无，已忽略后焦补偿面 {comp_surface}，测试波长 {test_wl}um）")
    else:
        log(f"已写入 TDE 公差: {n_tde} 条（不加 COMP 操作数，测试波长 {test_wl}um）")

    n_mfe, mf_path = mfe_builder.build_and_save(sess.sys, cfg.mfe, base)
    log(f"已重建 MFE: {n_mfe} 行  → {mf_path}")

    comp_mf_name = None
    if comp_on:
        if comp_surface <= 0:
            log(f"提示：补偿器模式={comp_mode} 但未填后焦补偿面，"
                f"TSC 仍写优化行/补偿 MF，但 TDE 不加 COMP 面（补偿由 TSC 负责）。")
        wave_for_mf = center_wave if center_wave > 0 else 2
        _n_comp, comp_mf_path = mfe_builder.build_comp_mf(
            sess.sys, base, freq_lp=comp_freq, wave=wave_for_mf)
        comp_mf_name = os.path.basename(comp_mf_path)
        log(f"已生成补偿专用 MF（GMTA {comp_freq}lp/mm）→ {comp_mf_path}")

    mf_name = os.path.basename(mf_path)
    optimize_cycles = _as_int(rp.get("TSC优化周期"), 4)
    n_report, tsc_path = tsc_builder.build_and_write(
        sess.sys, cfg.report, mf_name, base, optimize_cycles=optimize_cycles,
        comp_mode=comp_mode, comp_mf_name=comp_mf_name)
    log(f"已生成 TSC: {n_report} 个 REPORT 分项  → {tsc_path}")

    sess.sys.Save()

    run_config_path = os.path.join(out, "run_config.json")
    _write_json(run_config_path, {
        "source_zmx": os.path.abspath(zmx),
        "config_excel": os.path.abspath(config),
        "parent_outdir": parent_out,
        "result_outdir": out,
        "connect": connect,
        "working_copy": os.path.abspath(copy),
        "lens_dir": lens_dir,
        "tsc_path": tsc_path,
        "mf_path": mf_path,
        "counts": {
            "tde": n_tde,
            "mfe": n_mfe,
            "report": n_report,
        },
        "run_params": rp,
    })
    log(f"运行配置快照: {run_config_path}")

    return Prepared(sess=sess, cfg=cfg, rp=rp, base=base, outdir=out,
                    tsc_path=tsc_path, mf_path=mf_path,
                    n_tde=n_tde, n_mfe=n_mfe, n_report=n_report,
                    lens_dir=lens_dir, parent_outdir=parent_out,
                    source_zmx=os.path.abspath(zmx),
                    config_path=os.path.abspath(config),
                    log_path=log_path, run_config_path=run_config_path)


def make_runspec(prep: Prepared) -> tol_runner.RunSpec:
    """按 Excel 运行参数与 Prepared 生成 RunSpec（ZTD/前缀落 outdir）。"""
    rp = prep.rp
    ztd_path = os.path.join(prep.outdir, f"{prep.base}.ZTD")
    save_worst = _yes(rp.get("保存WorstCase", "Y"))
    save_best = _yes(rp.get("保存BestCase", "Y"))
    return tol_runner.RunSpec(
        tsc_name=os.path.basename(prep.tsc_path),
        num_runs=_as_int(rp.get("蒙特卡洛次数"), 200),
        num_to_save=_as_int(rp.get("保存数量"), 10),
        comp_mode=str(rp.get("补偿器模式") or "近轴焦点").strip(),
        distribution=str(rp.get("统计分布") or "正态").strip(),
        ztd_path=ztd_path,
        save_best_worst=save_worst or save_best,
        file_prefix=prep.base,
        lens_dir=prep.lens_dir,
    )


def log_run_plan(prep: Prepared, spec: tol_runner.RunSpec, log=print,
                 export_stats: bool | None = None) -> None:
    """打印本次运行的输出清单与保存策略。"""
    save_worst = _yes(prep.rp.get("保存WorstCase", "Y"))
    save_best = _yes(prep.rp.get("保存BestCase", "Y"))
    used_excel = os.path.join(prep.outdir, "used_excel.xlsx")
    stat_path = f"{os.path.splitext(spec.ztd_path)[0]}_统计.xlsx"
    log("本次运行输出清单：")
    log(f"  结果目录: {prep.outdir}")
    log(f"  工作副本目录: {prep.lens_dir}")
    log(f"  ZTD 目标文件: {spec.ztd_path}")
    if export_stats is None:
        log(f"  统计 Excel: {stat_path}（按运行参数决定是否生成）")
    elif export_stats:
        log(f"  统计 Excel: {stat_path}")
    else:
        log("  统计 Excel: 不生成")
    log(f"  运行日志: {prep.log_path}")
    log(f"  运行配置快照: {prep.run_config_path}")
    log(f"  Excel 配置快照: {used_excel}")
    log("本次保存策略：")
    log(f"  蒙特卡洛次数: {spec.num_runs}")
    log(f"  保存数量: {spec.num_to_save}")
    log(f"  Worst/Best 保存: {'开启' if spec.save_best_worst else '关闭'}"
        f"（Worst={'Y' if save_worst else 'N'}, Best={'Y' if save_best else 'N'}；Zemax API 为同一开关）")


def run_montecarlo(prep: Prepared, log=print,
                   export_stats: bool | None = None):
    """方案A：API 直接跑完蒙特卡洛。返回 RunResult。"""
    spec = make_runspec(prep)
    log_run_plan(prep, spec, log=log, export_stats=export_stats)
    log(f"开始公差分析：{spec.num_runs} 次蒙特卡洛（{spec.distribution}分布）…")

    def on_progress(progress: int, msg: str) -> None:
        log(f"  [{progress:>3}%] {msg}")

    return tol_runner.run(prep.sess.sys, spec, progress_cb=on_progress)
