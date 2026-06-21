"""pipeline.py —— 公差分析的可复用流程（两方案共享准备逻辑）。

把「连接 → 建 TDE → 建 MFE → 建 TSC → Save」这段两方案完全相同的
准备工作抽成 prepare_session()，避免方案A/方案B 两份入口重复维护：

  方案A（全自动）：prepare_session() + run_montecarlo()，API 直接跑完。
  方案B（手动）  ：只 prepare_session()，停下，由用户在 OpticStudio
                  公差分析对话框选 TSC、点 Run，观察原生窗口实时进度。
"""

from __future__ import annotations

import os
from dataclasses import dataclass

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


def prepare_session(zmx: str, config: str, outdir: str | None = None,
                    connect: str = "extension", log=print,
                    zos_dir: str | None = None) -> Prepared:
    """连接 Zemax、打开工作副本，按 Excel 重建 TDE/MFE/TSC 并 Save。

    两方案共用。返回 Prepared，方案A 据此跑蒙卡，方案B 据此提示用户。
    """
    cfg = excel_io.read_config(config)
    rp = cfg.run_params
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

    copy = sess.open_as_copy(zmx)
    log(f"工作副本: {copy}")

    base = os.path.splitext(os.path.basename(copy))[0]
    lens_dir = os.path.dirname(os.path.abspath(copy))
    out = outdir or lens_dir
    os.makedirs(out, exist_ok=True)

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

    return Prepared(sess=sess, cfg=cfg, rp=rp, base=base, outdir=out,
                    tsc_path=tsc_path, mf_path=mf_path,
                    n_tde=n_tde, n_mfe=n_mfe, n_report=n_report,
                    lens_dir=lens_dir)


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


def run_montecarlo(prep: Prepared, log=print):
    """方案A：API 直接跑完蒙特卡洛。返回 RunResult。"""
    spec = make_runspec(prep)
    log(f"开始公差分析：{spec.num_runs} 次蒙特卡洛（{spec.distribution}分布）…")

    def on_progress(progress: int, msg: str) -> None:
        log(f"  [{progress:>3}%] {msg}")

    return tol_runner.run(prep.sess.sys, spec, progress_cb=on_progress)
