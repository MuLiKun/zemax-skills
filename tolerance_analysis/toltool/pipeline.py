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


def prepare_session(zmx: str, config: str, outdir: str | None = None,
                    connect: str = "extension", log=print) -> Prepared:
    """连接 Zemax、打开工作副本，按 Excel 重建 TDE/MFE/TSC 并 Save。

    两方案共用。返回 Prepared，方案A 据此跑蒙卡，方案B 据此提示用户。
    """
    cfg = excel_io.read_config(config)
    rp = cfg.run_params
    center_wave = _as_int(rp.get("中心波长号"), 0)

    sess = zos_connect.ZosSession()
    log(f"ZOS 目录: {sess.zos_dir}")
    sess.connect(mode=connect)
    log(f"已连接: {sess.sys.SystemFile}")

    copy = sess.open_as_copy(zmx)
    log(f"工作副本: {copy}")

    base = os.path.splitext(os.path.basename(copy))[0]
    out = outdir or os.path.dirname(os.path.abspath(copy))
    os.makedirs(out, exist_ok=True)

    test_wl = 0.0
    if center_wave > 0:
        try:
            test_wl = float(sess.sys.SystemData.Wavelengths
                            .GetWavelength(center_wave).Wavelength)
        except Exception as e:
            log(f"读取测试波长失败(忽略): {e}")

    n_tde = tde_builder.build_and_write(
        sess.sys, cfg.tol_wizard, cfg.tol_detail,
        center_wave=center_wave, test_wavelength_um=test_wl)
    log(f"已写入 TDE 公差: {n_tde} 条（测试波长 {test_wl}um）")

    n_mfe, mf_path = mfe_builder.build_and_save(sess.sys, cfg.mfe, base)
    log(f"已重建 MFE: {n_mfe} 行  → {mf_path}")

    mf_name = os.path.basename(mf_path)
    optimize_cycles = _as_int(rp.get("TSC优化周期"), 4)
    n_report, tsc_path = tsc_builder.build_and_write(
        sess.sys, cfg.report, mf_name, base, optimize_cycles=optimize_cycles)
    log(f"已生成 TSC: {n_report} 个 REPORT 分项  → {tsc_path}")

    sess.sys.Save()

    return Prepared(sess=sess, cfg=cfg, rp=rp, base=base, outdir=out,
                    tsc_path=tsc_path, mf_path=mf_path,
                    n_tde=n_tde, n_mfe=n_mfe, n_report=n_report)


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
    )


def run_montecarlo(prep: Prepared, log=print):
    """方案A：API 直接跑完蒙特卡洛。返回 RunResult。"""
    spec = make_runspec(prep)
    log(f"开始公差分析：{spec.num_runs} 次蒙特卡洛（{spec.distribution}分布）…")

    def on_progress(progress: int, msg: str) -> None:
        log(f"  [{progress:>3}%] {msg}")

    return tol_runner.run(prep.sess.sys, spec, progress_cb=on_progress)
