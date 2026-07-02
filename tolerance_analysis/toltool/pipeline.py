"""pipeline.py —— 公差分析的可复用流程。

GUI 与命令行入口共享这里的核心流程：
连接 → 建 TDE → 建 MFE → 建 TSC → Save → 跑蒙卡。
"""

from __future__ import annotations

import copy
import json
import os
import re
from dataclasses import dataclass
from datetime import datetime

from . import (zos_connect, excel_io, tde_builder, mfe_builder,
               tsc_builder, tol_runner, field_mapping, current_settings)


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


def _enum_name(value) -> str:
    text = str(value or "").strip()
    if "." in text:
        text = text.rsplit(".", 1)[-1]
    text = text.upper()
    match = re.search(r"[A-Z]{3,4}", text)
    return match.group(0) if match else text


def _read_tde_meta(zos_system) -> list[dict]:
    tde = zos_system.TDE
    rows: list[dict] = []
    for i in range(1, int(tde.NumberOfOperands) + 1):
        r = tde.GetOperandAt(i)
        op = _enum_name(getattr(r, "Type", ""))
        if not op or op == "BLNK":
            continue
        row = {"行号": i, "操作数": op}
        for name in ("Param1", "Param2", "Min", "Max", "Comment"):
            try:
                row[name] = getattr(r, name)
            except Exception:
                row[name] = ""
        if op == "COMP":
            comment = str(row.get("Comment") or "").strip()
            surf = row.get("Param1")
            suffix = f"_S{surf}" if str(surf).strip() else ""
            row["标签"] = comment or f"COMP{suffix}"
            row["方向"] = ""
            row["单位"] = "mm"
        rows.append(row)
    return rows


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


def _fmt_num(value) -> str:
    if value is None:
        return "-"
    try:
        return f"{float(value):.4f}"
    except (TypeError, ValueError):
        return str(value)


def _insert_strategy_label(value) -> str:
    text = str(value or "").strip()
    if _yes(text):
        return "自动插入"
    if not text:
        return "禁用"
    return text


def field_mapping_report_lines(result) -> list[str]:
    missing = [item for item in result.final_matches if item.need_insert]
    inserted_targets = {item.target_normalized for item in result.inserted_fields}
    matched_count = len(result.final_matches) - len(missing)
    title = "视场映射预览" if getattr(result, "is_preview", False) else "视场映射报告"
    if getattr(result, "is_preview", False) and result.inserted_fields:
        title += "（已模拟补齐）"
    if result.inserted_fields:
        conclusion = f"已匹配 {matched_count}/{len(result.targets)}，本次模拟补齐 {len(result.inserted_fields)}"
    else:
        conclusion = f"已匹配 {matched_count}/{len(result.targets)}，需补齐 {len(missing)}"
    field_by_no = {field.field_no: field for field in result.final_fields}
    lines = [
        title,
        f"结论: {conclusion}",
        f"阈值: {result.threshold:g}；插入策略: {_insert_strategy_label(result.insert_strategy)}",
        "",
        "目标      视场号        X        Y    归一化    偏差  来源/状态",
    ]
    for item in result.final_matches:
        if item.need_insert:
            status = "仍需补齐"
        elif item.target_normalized in inserted_targets:
            status = "已补齐"
        else:
            status = "已有"
        field = field_by_no.get(item.field_no)
        x = _fmt_num(field.x if field else None)
        y = _fmt_num(field.y if field else None)
        lines.append(
            f"{item.report_label:<8} {str(item.field_no or '-'):>5}  "
            f"{x:>7}  {y:>7}  {_fmt_num(item.actual_normalized):>7}  "
            f"{_fmt_num(item.delta):>6}  {status}")
    if not getattr(result, "is_preview", False):
        lines.extend([
            "",
            f"已插入缺失视场: {len(result.inserted_fields)}",
            f"已改写 MFE: {result.mfe_updates} 行",
            f"已改写 REPORT: {result.report_updates} 项",
        ])
    if missing:
        lines.append("")
        lines.append("需补齐目标: " + ", ".join(item.report_label for item in missing))
    elif result.messages and not getattr(result, "is_preview", False):
        lines.append("")
        lines.extend(str(msg) for msg in result.messages)
    return lines


def _log_field_mapping(result, log) -> None:
    for line in field_mapping_report_lines(result):
        if line:
            log(line)


def _write_field_mapping_report(path: str, result) -> None:
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(field_mapping_report_lines(result)) + "\n")


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
    comp_mode_key = str(rp.get("补偿器模式") or "无").strip().replace(" ", "").lower()
    if comp_mode_key not in ("无", "none", "全部优化dls", "全部优化(dls)", "dls",
                             "全部优化od", "全部优化(od)", "od"):
        errors.append("补偿器模式仅支持：无、全部优化DLS、全部优化OD")

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


def validate_config_data(cfg):
    """只校验已读取/生成的配置内容，不连接 Zemax。"""
    _validate_inputs(cfg, cfg.run_params)
    return cfg


def validate_config(zmx: str, config: str):
    """只校验文件路径和 Excel 配置，不连接 Zemax。"""
    _validate_paths(zmx, config)
    cfg = excel_io.read_config(config)
    return validate_config_data(cfg)


def _fill_auto_standard_surfaces(zos_system, cfg, rp: dict, log=print) -> None:
    if str(rp.get("分析模式") or "").strip() != "标准模板":
        return
    try:
        end_surface = max(1, int(zos_system.LDE.NumberOfSurfaces) - 2)
    except (AttributeError, TypeError, ValueError) as e:
        log(f"自动读取镜头面数失败，保留 Excel 中的公差范围: {e}")
        return
    changed = False
    for row in cfg.tol_wizard:
        if _as_int(row.get("结束面"), 0) <= 0:
            row["结束面"] = end_surface
            changed = True
    if changed:
        log(f"标准模板自动公差范围: 1-{end_surface}")


def _operand_value(zos_system, op: str, *params: float) -> float:
    import ZOSAPI

    values = list(params[:8])
    while len(values) < 8:
        values.append(0)
    op_enum = getattr(ZOSAPI.Editors.MFE.MeritOperandType, op)
    return float(zos_system.MFE.GetOperandValue(op_enum, *values))


def _next_mfe_line(cfg) -> int:
    lines: list[int] = []
    for row in cfg.mfe:
        try:
            lines.append(int(float(row.get("行号"))))
        except (TypeError, ValueError):
            pass
    return (max(lines) + 1) if lines else 2


def _mfe_row(line_no: int, op: str, *, target=0, weight=0,
             comment: str = "", field=None, **params) -> dict:
    row = {
        "行号": line_no,
        "操作数": op,
        "目标": target,
        "权重": weight,
        "注释": comment,
        "目标归一化视场": "" if field is None else field,
        "视场映射说明": "标准模板动态生成",
        "归一化视场": "" if field is None else field,
    }
    for i in range(1, 9):
        row[f"Param{i}"] = params.get(f"Param{i}", "")
    return row


def _report_labels(cfg) -> set[str]:
    return {str(row.get("标签") or "").strip() for row in cfg.report}


def _append_standard_dynamic_metrics(zos_system, cfg, rp: dict, center_wave: int,
                                     log=print):
    if str(rp.get("分析模式") or "").strip() != "标准模板":
        return cfg
    if center_wave <= 0:
        log("标准模板动态评价项：中心波长号无效，已跳过中心指向偏移和焦距偏移百分比。")
        return cfg

    new_cfg = copy.deepcopy(cfg)
    labels = _report_labels(new_cfg)
    line = _next_mfe_line(new_cfg)
    added: list[str] = []

    if _yes(rp.get("启用中心指向偏移", "N")) and not {
        "POINTING_DY_F0_mm", "POINTING_DX_F0_mm"}.issubset(labels):
        field_no = _as_int(rp.get("中心指向视场号"), 1)
        if field_no <= 0:
            field_no = 1
        ceny0 = _operand_value(zos_system, "CENY", 16, center_wave, field_no, 0, 5)
        cenx0 = _operand_value(zos_system, "CENX", 16, center_wave, field_no, 0, 5)

        new_cfg.mfe.append(_mfe_row(line, "BLNK", comment="接收指向偏移 F0/mm")); line += 1
        ceny_line = line
        new_cfg.mfe.append(_mfe_row(line, "CENY", comment="CENY_F0_current", field=0,
                                    Param1=16, Param2=center_wave, Param3=field_no,
                                    Param4=0, Param5=5)); line += 1
        cenx_line = line
        new_cfg.mfe.append(_mfe_row(line, "CENX", comment="CENX_F0_current", field=0,
                                    Param1=16, Param2=center_wave, Param3=field_no,
                                    Param4=0, Param5=5)); line += 1
        cons_y_line = line
        new_cfg.mfe.append(_mfe_row(line, "CONS", target=ceny0, comment="CENY_F0_nominal")); line += 1
        cons_x_line = line
        new_cfg.mfe.append(_mfe_row(line, "CONS", target=cenx0, comment="CENX_F0_nominal")); line += 1
        diff_y_line = line
        new_cfg.mfe.append(_mfe_row(line, "DIFF", comment="POINTING_DY_F0_mm",
                                    Param1=ceny_line, Param2=cons_y_line)); line += 1
        diff_x_line = line
        new_cfg.mfe.append(_mfe_row(line, "DIFF", comment="POINTING_DX_F0_mm",
                                    Param1=cenx_line, Param2=cons_x_line)); line += 1
        new_cfg.report.extend([
            {"启用": "Y", "标签": "POINTING_DY_F0_mm", "MF行号": diff_y_line, "方向": "小", "单位": "mm"},
            {"启用": "Y", "标签": "POINTING_DX_F0_mm", "MF行号": diff_x_line, "方向": "小", "单位": "mm"},
        ])
        added.extend(["POINTING_DY_F0_mm", "POINTING_DX_F0_mm"])
        log(f"中心指向偏移 F0：视场号 {field_no}，CENY0={ceny0:.12g}，CENX0={cenx0:.12g}")

    if _yes(rp.get("启用焦距偏移百分比", "N")) and "EFL_DELTA_PCT" not in _report_labels(new_cfg):
        efl0 = _operand_value(zos_system, "EFFL", center_wave)
        if abs(efl0) < 1e-15:
            log("焦距偏移百分比：名义 EFL 接近 0，已跳过 EFL_DELTA_PCT。")
        else:
            new_cfg.mfe.append(_mfe_row(line, "BLNK", comment="EFL")); line += 1
            efl_line = line
            new_cfg.mfe.append(_mfe_row(line, "EFFL", comment="EFFL_current",
                                        Param1=center_wave)); line += 1
            cons_line = line
            new_cfg.mfe.append(_mfe_row(line, "CONS", target=efl0, comment="EFFL_nominal")); line += 1
            diff_line = line
            new_cfg.mfe.append(_mfe_row(line, "DIFF", comment="EFL_delta",
                                        Param1=efl_line, Param2=cons_line)); line += 1
            divi_line = line
            new_cfg.mfe.append(_mfe_row(line, "DIVI", comment="EFL_delta_ratio",
                                        Param1=diff_line, Param2=cons_line)); line += 1
            new_cfg.mfe.append(_mfe_row(line, "BLNK", comment="DELTA EFL")); line += 1
            pct_line = line
            new_cfg.mfe.append(_mfe_row(line, "PROB", target=100, comment="EFL_DELTA_PCT",
                                        Param1=divi_line)); line += 1
            new_cfg.report.append({ 
                "启用": "Y",
                "标签": "EFL_DELTA_PCT",
                "MF行号": pct_line,
                "方向": "小",
                "单位": "%",
            })
            added.append("EFL_DELTA_PCT")
            log(f"焦距偏移百分比：EFFL0={efl0:.12g}，已追加 EFL_DELTA_PCT。")

    if added:
        log("标准模板动态评价项：已追加 " + ", ".join(added))
    return new_cfg


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
    field_mapping_report_path: str = ""
    tde_meta: list | None = None


def prepare_session(zmx: str, config: str, outdir: str | None = None,
                    connect: str = "extension", log=print,
                    zos_dir: str | None = None,
                    use_current_settings: bool = False,
                    current_args: dict | None = None) -> Prepared:
    """连接 Zemax、打开工作副本并完成 TSC/Save 准备。

    默认按 Excel 重建 TDE/MFE；use_current_settings=True 时保留当前 TDE/MFE，
    只从工作副本当前 MFE 自动生成 REPORT/TSC。
    """
    if use_current_settings:
        if not os.path.isfile(zmx):
            raise ValueError(f"运行前校验失败：\n- ZMX 文件不存在：{zmx}")
        cfg = None
        rp = {}
    else:
        cfg = validate_config(zmx, config)
        rp = cfg.run_params

    parent_out, out = _make_run_dir(zmx, outdir)
    log_path = os.path.join(out, "run.log")
    log = _tee_logger(log, log_path)
    log(f"结果目录: {out}")
    log(f"日志文件: {log_path}")
    used_excel_path = os.path.join(out, "used_excel.xlsx")

    standard_mode = str(rp.get("分析模式") or "").strip() == "标准模板"
    center_wave = _as_int(rp.get("中心波长号"), 0)
    center_wave_auto = standard_mode and center_wave <= 0
    comp_surface = _as_int(rp.get("后焦补偿面"), 0)
    comp_min = _num(rp.get("补偿Min"))
    comp_max = _num(rp.get("补偿Max"))
    comp_freq = _num(rp.get("补偿线对"), 34.0)
    comp_mode = str(rp.get("补偿器模式") or "无").strip()
    # 补偿器模式=无 → 完全不补偿：TSC 不写优化行/不走双 MF、不建 comp MF、不加 TDE COMP。
    # 模式≠无 → 写优化行 + comp MF（双 MF）+ TDE COMP；未填后焦补偿面时自动用像面前一面。
    comp_off = str(comp_mode).replace(" ", "").lower() in ("无", "none", "")
    comp_on = not comp_off

    sess = zos_connect.ZosSession(zos_dir=zos_dir)
    log(f"ZOS 目录: {sess.zos_dir}")
    sess.connect(mode=connect)
    if connect == "standalone":
        log("已启动 Zemax 独立实例（即将载入下面的工作副本）")
    else:
        log(f"已连接交互扩展: {sess.sys.SystemFile}")

    src_base, src_ext = os.path.splitext(os.path.basename(zmx))
    safe_src_base = _safe_name(src_base)
    if safe_src_base != src_base:
        log(f"提示：镜头文件名包含空格或特殊字符，Zemax 输出前缀将使用安全名称: {safe_src_base}")
    copy_path = os.path.join(out, f"{safe_src_base}_tol{src_ext}")

    copy = sess.open_as_copy(zmx, copy_path=copy_path)
    log(f"工作副本: {copy}")
    if use_current_settings:
        current_args = current_args or {}
        cfg = current_settings.build_config_from_current_mfe(
            sess.sys,
            num_runs=_as_int(current_args.get("num_runs"), 20),
            num_to_save=_as_int(current_args.get("num_to_save"), 0),
            comp_mode=str(current_args.get("comp_mode") or "无"),
            save_worst_best=_yes(current_args.get("save_worst_best")),
            report_filter=current_args.get("report_filter"))
        rp = cfg.run_params
        _validate_inputs(cfg, rp)
        standard_mode = False
        center_wave = _as_int(rp.get("中心波长号"), 0)
        center_wave_auto = False
        comp_surface = _as_int(rp.get("后焦补偿面"), 0)
        comp_min = _num(rp.get("补偿Min"))
        comp_max = _num(rp.get("补偿Max"))
        comp_freq = _num(rp.get("补偿线对"), 34.0)
        comp_mode = str(rp.get("补偿器模式") or "无").strip()
        comp_off = str(comp_mode).replace(" ", "").lower() in ("无", "none", "")
        comp_on = not comp_off
        has_tde_comp = current_settings.tde_has_comp(sess.sys)
        if comp_on and not has_tde_comp:
            log("当前 TDE 无 COMP 操作数，强制关闭补偿器优化（TSC 不写 OPTIMIZE 行）。")
            comp_mode = "无"
            comp_off = True
            comp_on = False
            rp["补偿器模式"] = "无"
        elif comp_on and has_tde_comp:
            comp_freq = current_settings.detect_comp_freq_from_mfe(cfg)
            if comp_freq and comp_freq > 0:
                rp["补偿线对"] = comp_freq
                log(f"当前 TDE 含 COMP，补偿 MF 线对取自 MFE MTF 操作数: {comp_freq} lp/mm")
            else:
                log("当前 TDE 含 COMP，补偿 MF 线对将使用默认值")
        log("当前设置模式：已读取当前 MFE，保留当前 TDE。")
        for msg in current_settings.summarize_config(cfg):
            log(msg)
    lens_info = sess.read_lens_info()
    if comp_on and comp_surface <= 0 and not use_current_settings:
        comp_surface = max(1, int(sess.sys.LDE.NumberOfSurfaces) - 2)
        rp["后焦补偿面"] = comp_surface
        log(f"未填后焦补偿面，已自动设置为像面前一面: {comp_surface}")
    add_comp_operand = comp_on and comp_surface > 0 and not use_current_settings
    if center_wave_auto and lens_info.primary_wave > 0:
        center_wave = lens_info.primary_wave
        rp["中心波长号"] = center_wave
        for row in cfg.mfe:
            op = str(row.get("操作数") or "").strip().upper()
            if op in ("RSCE", "GENC", "GMTT", "GMTS", "GMTA") \
                    and _as_int(row.get("Param2"), 0) <= 0:
                row["Param2"] = center_wave
        log(f"已自动使用主波长号: {center_wave}")
    _fill_auto_standard_surfaces(sess.sys, cfg, rp, log=log)

    cfg, field_mapping_result = field_mapping.process(sess.sys, cfg, rp, log=log)
    field_mapping_report_path = ""
    if field_mapping_result.enabled:
        log(f"已启用视场映射：目标 {len(field_mapping_result.targets)} 个，"
            f"阈值 {field_mapping_result.threshold:g}，插入策略={field_mapping_result.insert_strategy}")
        _log_field_mapping(field_mapping_result, log)
        field_mapping_report_path = os.path.join(out, "field_mapping.txt")
        try:
            _write_field_mapping_report(field_mapping_report_path, field_mapping_result)
            log(f"视场映射报告: {field_mapping_report_path}")
        except Exception as e:
            log(f"保存视场映射报告失败(忽略): {e}")
            field_mapping_report_path = ""
    else:
        log("视场映射：未启用")

    cfg = _append_standard_dynamic_metrics(sess.sys, cfg, rp, center_wave, log=log)

    try:
        if use_current_settings:
            current_settings.write_config_excel(used_excel_path, cfg, overwrite=True)
        else:
            excel_io.write_config_snapshot(config, used_excel_path, cfg)
        log(f"Excel 配置快照: {used_excel_path}")
    except Exception as e:
        log(f"保存 Excel 配置快照失败(忽略): {e}")

    base = os.path.splitext(os.path.basename(copy))[0]
    lens_dir = os.path.dirname(os.path.abspath(copy))

    test_wl = 0.0
    if center_wave > 0:
        try:
            test_wl = float(sess.sys.SystemData.Wavelengths
                            .GetWavelength(center_wave).Wavelength)
        except Exception as e:
            log(f"读取测试波长失败(忽略): {e}")

    if use_current_settings:
        n_tde = 0
        log("当前设置模式：跳过 TDE 重建，保留镜头文件现有 TDE。")
        n_mfe = len(cfg.mfe)
        mf_path = mfe_builder.default_mf_path(sess.sys, base)
        mfe_builder.save_mf(sess.sys, mf_path)
        log(f"当前设置模式：已保存当前 MFE: {n_mfe} 行  → {mf_path}")
    else:
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
        wave_for_mf = center_wave if center_wave > 0 else (lens_info.primary_wave or 2)
        _n_comp, comp_mf_path = mfe_builder.build_comp_mf(
            sess.sys, base, freq_lp=comp_freq, wave=wave_for_mf)
        comp_mf_name = os.path.basename(comp_mf_path)
        log(f"已生成补偿专用 MF（GMTA {comp_freq}lp/mm）→ {comp_mf_path}")

    mf_name = os.path.basename(mf_path)
    optimize_cycles = _as_int(rp.get("TSC优化周期"), 4)
    n_report, tsc_path = tsc_builder.build_and_write(
        sess.sys, cfg.report, mf_name, base, optimize_cycles=optimize_cycles,
        comp_mode=comp_mode, comp_mf_name=comp_mf_name, log=log)
    log(f"已生成 TSC: {n_report} 个 REPORT 分项  → {tsc_path}")

    tde_meta = _read_tde_meta(sess.sys)
    comp_count = sum(1 for row in tde_meta if row.get("操作数") == "COMP")
    if comp_count:
        log(f"已记录 TDE 元数据：{len(tde_meta)} 个公差操作数，其中 COMP {comp_count} 个。")

    sess.sys.Save()

    run_config_path = os.path.join(out, "run_config.json")
    _write_json(run_config_path, {
        "source_zmx": os.path.abspath(zmx),
        "config_excel": os.path.abspath(config) if config else "",
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
        "tde_meta": tde_meta,
        "run_params": rp,
        "field_mapping": field_mapping_result.to_dict(),
        "field_mapping_report": field_mapping_report_path,
    })
    log(f"运行配置快照: {run_config_path}")

    return Prepared(sess=sess, cfg=cfg, rp=rp, base=base, outdir=out,
                    tsc_path=tsc_path, mf_path=mf_path,
                    n_tde=n_tde, n_mfe=n_mfe, n_report=n_report,
                    lens_dir=lens_dir, parent_outdir=parent_out,
                    source_zmx=os.path.abspath(zmx),
                    config_path=os.path.abspath(config) if config else "",
                    log_path=log_path, run_config_path=run_config_path,
                    field_mapping_report_path=field_mapping_report_path,
                    tde_meta=tde_meta)


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
        comp_mode=str(rp.get("补偿器模式") or "无").strip(),
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
    if prep.field_mapping_report_path:
        log(f"  视场映射报告: {prep.field_mapping_report_path}")
    log("本次保存策略：")
    log(f"  蒙特卡洛次数: {spec.num_runs}")
    log(f"  保存数量: {spec.num_to_save}")
    log(f"  Worst/Best 保存: {'开启' if spec.save_best_worst else '关闭'}"
        f"（Worst={'Y' if save_worst else 'N'}, Best={'Y' if save_best else 'N'}；Zemax API 为同一开关）")


def run_montecarlo(prep: Prepared, log=print,
                   export_stats: bool | None = None):
    """通过 Zemax API 直接跑完蒙特卡洛。返回 RunResult。"""
    spec = make_runspec(prep)
    log_run_plan(prep, spec, log=log, export_stats=export_stats)
    log(f"开始公差分析：{spec.num_runs} 次蒙特卡洛（{spec.distribution}分布）…")

    def on_progress(progress: int, msg: str) -> None:
        log(f"  [{progress:>3}%] {msg}")

    return tol_runner.run(prep.sess.sys, spec, progress_cb=on_progress)
