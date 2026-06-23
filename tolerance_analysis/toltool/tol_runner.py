r"""tol_runner.py —— 配置并运行脚本式蒙特卡洛公差分析（界面解耦）。

职责（需求文档 §4.7/§7、skill §2b/§2c/§2d）：
- 配置 OpenTolerancing：SetupModeIndex=0 + CriterionIndex=15(自定义脚本)
  + CriterionScript(按 TSC 文件名定位索引) + 补偿器/分布(按中文匹配)
  + NumberOfRuns/NumberToSave + SaveTolDataFile=True/TolDataFile
  + Worst/Best 保存(IsSaveBestWorstUsed + FilePrefix；
    BestWorstOutputFolder 实测只读，输出位置由 Zemax 决定，运行后回读)。
- 核心 run() 与界面解耦：非阻塞 Run() + 轮询 IsRunning，
  通过 progress_cb 回调进度、cancel_flag 支持取消。
  既可被命令行调用，也可被 GUI 前台线程调用。

实测要点（probe 验证）：
- ListOf* 在 Python 端为 None；改用 GetXxxAt(i)+NumberOfXxx 遍历。
- CriterionScripts 列表项是文件名(如 '05304_tol.TSC')，CriterionScript 取其索引。
- 完成判据：IsRunning 变 False 且 Succeeded == True。
- Progress 在 TSC 模式可能直到结束才到 100。
- ZTD 正确写法★：运行前 SaveTolDataFile=True + TolDataFile=纯文件名
  (os.path.basename，含 .ZTD 扩展名，不带任何路径)。Zemax 官方约定：
  TolDataFile/OutputFile 只接受文件名，落盘目录强制为当前镜头文件
  (即工作副本)所在目录，由 Zemax 决定、不可自定义。传带路径的字符串
  (含绝对路径)会被判为非法文件名而静默不落盘，自报 Succeeded=True 却
  无 ZTD，造成"假成功"——副本在镜头根目录时偶能侥幸解析，副本移入子
  目录后必然失效。故只传文件名，并保证工作副本已在目标输出目录，ZTD
  即随副本落到该目录。绝不在运行后调 tol.Save(ztd)，那会写出损坏 ZTD。
  运行后按多候选路径(副本目录/lens_dir/Documents\Zemax\Tolerance)回读
  实际生成的 ZTD；要求保存 ZTD 却回读不到时判为失败，不再"假成功"。
"""

from __future__ import annotations

import os
import time
from dataclasses import dataclass, field


COMP_ALIASES = {
    "全部优化dls": ["全部优化 (dls)", "全部优化(dls)", "dls"],
    "近轴焦点": ["近轴焦点", "近轴"],
    "无": ["无"],
    "全部优化od": ["全部优化 (od)", "全部优化(od)", "od"],
}

DIST_ALIASES = {
    "正态": ["正态分布", "正态", "normal"],
    "均匀": ["均匀", "uniform"],
    "抛物线": ["抛物线", "parabolic"],
}

_INVALID_FILE_CHARS = set('<>:"/\\|?*')
_RESERVED_FILE_STEMS = {
    "CON", "PRN", "AUX", "NUL",
    "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
    "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
}


@dataclass
class RunSpec:
    tsc_name: str
    num_runs: int = 200
    num_to_save: int = 10
    comp_mode: str = "近轴焦点"
    distribution: str = "正态"
    ztd_path: str = ""
    save_best_worst: bool = True
    file_prefix: str = ""
    lens_dir: str = ""


@dataclass
class RunResult:
    succeeded: bool
    ztd_path: str
    num_runs: int
    num_cols_hint: int = 0
    message: str = ""
    bestworst_folder: str = ""
    saved_files: list = field(default_factory=list)


def _find_index(tol, count_attr, getter_attr, target, aliases=None):
    """在索引列表里按字符串匹配返回索引；找不到返回 -1。"""
    n = getattr(tol, count_attr)
    getter = getattr(tol, getter_attr)
    cand = [target]
    if aliases:
        key = str(target).strip().lower().replace(" ", "")
        for k, vs in aliases.items():
            if key == k or str(target).strip() in vs:
                cand = vs + [target]
                break
    norm = lambda s: str(s).strip().lower().replace(" ", "")
    cset = {norm(c) for c in cand}
    for i in range(n):
        item = norm(getter(i))
        if item in cset:
            return i
    for i in range(n):
        item = norm(getter(i))
        if any(norm(c) and norm(c) in item for c in cand):
            return i
    return -1


def _win_long_path(path: str) -> str:
    r"""Windows 下对超过 MAX_PATH(260) 的本地绝对路径加 \\?\ 扩展前缀。

    扩展前缀要求纯反斜杠且为绝对路径，故先规范化分隔符。已带前缀或
    非 Windows 平台原样返回。
    """
    if os.name != "nt" or not path:
        return path
    if path.startswith("\\\\?\\"):
        return path
    if len(path) < 260:
        return path
    norm = os.path.normpath(path)
    if norm.startswith("\\\\"):
        return "\\\\?\\UNC\\" + norm[2:]
    return "\\\\?\\" + norm


def _find_script_index(tol, tsc_name) -> int:
    """按 TSC 文件名(忽略大小写/扩展名)定位 CriterionScript 索引。"""
    base = os.path.basename(str(tsc_name)).lower()
    base_noext = os.path.splitext(base)[0]
    n = tol.NumberOfCriterionScripts
    for i in range(n):
        item = str(tol.GetCriterionScriptAt(i)).lower()
        if item == base or os.path.splitext(item)[0] == base_noext:
            return i
    return -1


def _ztd_file_name(path: str) -> str:
    name = os.path.basename(str(path or "").strip())
    if not name:
        raise ValueError("ZTD 文件名为空，无法保存公差数据。")
    if any(ch in _INVALID_FILE_CHARS for ch in name):
        raise ValueError(f"ZTD 文件名包含非法字符：{name}")
    if name.endswith((" ", ".")):
        raise ValueError(f"ZTD 文件名不能以空格或点结尾：{name}")
    base, ext = os.path.splitext(name)
    if not base:
        raise ValueError("ZTD 文件名为空，无法保存公差数据。")
    if base.upper() in _RESERVED_FILE_STEMS:
        raise ValueError(f"ZTD 文件名不能使用 Windows 保留名称：{base}")
    if ext.lower() == ".ztd":
        return name
    return base + ".ZTD"


def _ztd_expected_path(path: str) -> str:
    folder = os.path.dirname(os.path.abspath(path))
    return os.path.join(folder, _ztd_file_name(path))


def _safe_text(obj, name: str, default="") -> str:
    try:
        value = getattr(obj, name)
        if value is None:
            return str(default) if default != "" else "无"
        return str(value)
    except Exception as e:
        return f"<读取失败: {e}>" if default == "" else str(default)


def _tool_error_message(tol) -> str:
    try:
        value = getattr(tol, "ErrorMessage")
    except Exception as e:
        return f"读取 Zemax 错误信息失败：{e}"
    text = str(value or "").strip()
    if not text or text.lower() == "none":
        return "运行未成功"
    return text


def _zemax_tolerance_dir(zos_system) -> str:
    try:
        mf_dir = str(zos_system.MFE.MeritFunctionDirectory or "")
        if not mf_dir:
            return ""
        data_dir = os.path.dirname(mf_dir.rstrip("\\/"))
        return os.path.join(data_dir, "Tolerance")
    except Exception:
        return ""


def _fmt_tool_line(label: str, name: str, value: str) -> str:
    return f"  - {label}（{name}）：{value}"


_COMMON_TOOL_FIELDS = (
    ("NumberToSave", "保存案例数量"),
    ("IsSaveBestWorstUsed", "是否保存Worst/Best"),
    ("FilePrefix", "案例文件前缀"),
    ("SaveTolDataFile", "是否保存ZTD"),
    ("TolDataFile", "ZTD文件名"),
)


def _format_tool_fields(tol, fields: tuple[tuple[str, str], ...]) -> list[str]:
    return [_fmt_tool_line(label, name, _safe_text(tol, name)) for name, label in fields]


def _tool_setting_lines(tol) -> list[str]:
    fields = (
        ("SetupModeIndex", "运行模式索引"),
        ("CriterionIndex", "判据索引"),
        ("CriterionScript", "TSC脚本索引"),
        ("CriterionCompIndex", "工具栏补偿器索引"),
        ("MonteCarloStatisticIndex", "蒙特卡洛分布索引"),
        ("NumberOfRuns", "蒙特卡洛次数"),
        *_COMMON_TOOL_FIELDS,
    )
    return _format_tool_fields(tol, fields)


def _tool_result_lines(tol) -> list[str]:
    fields = (
        ("Succeeded", "Zemax是否判定成功"),
        ("ErrorMessage", "Zemax错误信息"),
        ("Progress", "Zemax进度"),
        *_COMMON_TOOL_FIELDS[:2],
        ("BestWorstOutputFolder", "Worst/Best输出目录"),
        *_COMMON_TOOL_FIELDS[2:],
    )
    return _format_tool_fields(tol, fields)


def configure(tol, spec: RunSpec):
    """把 RunSpec 写入已打开的 Tolerancing 工具。返回 (script_index, warnings)。"""
    warnings: list[str] = []
    tol.SetupModeIndex = 0
    tol.CriterionIndex = 15

    si = _find_script_index(tol, spec.tsc_name)
    if si < 0:
        raise RuntimeError(
            f"在 CriterionScripts 列表里找不到 TSC: {spec.tsc_name}。"
            f"请确认 .TSC 已写入 Tolerance 目录。")
    tol.CriterionScript = si

    ci = _find_index(tol, "NumberOfCriterionComps", "GetCriterionCompAt",
                     spec.comp_mode, COMP_ALIASES)
    if ci >= 0:
        tol.CriterionCompIndex = ci
        if tol.CriterionCompIndex != ci:
            warnings.append(
                f"补偿器模式未生效(期望 {spec.comp_mode}，实际索引 "
                f"{tol.CriterionCompIndex}={tol.GetCriterionCompAt(tol.CriterionCompIndex)})；"
                f"脚本模式下补偿由 TSC 的 OPTIMIZE 负责，可忽略。")

    di = _find_index(tol, "NumberOfMonteCarloStatistics",
                     "GetMonteCarloStatisticAt", spec.distribution, DIST_ALIASES)
    if di >= 0:
        tol.MonteCarloStatisticIndex = di

    tol.NumberOfRuns = int(spec.num_runs)
    tol.NumberToSave = int(spec.num_to_save)

    if spec.save_best_worst:
        tol.IsSaveBestWorstUsed = True
        if spec.file_prefix:
            tol.FilePrefix = spec.file_prefix

    if spec.ztd_path:
        tol.SaveTolDataFile = True
        tol.TolDataFile = _ztd_file_name(spec.ztd_path)

    return si, warnings


def run(zos_system, spec: RunSpec, progress_cb=None, cancel_flag=None,
        poll_interval: float = 2.0, timeout_s: float = 0.0) -> RunResult:
    """配置并运行公差分析。界面解耦：progress_cb(progress:int, msg:str)、
    cancel_flag(callable -> bool 或带 .is_set())。

    注意：扩展模式下对 tol.IsRunning 高频轮询(<1.5s)会导致 IPC 管道关闭，
    poll_interval 默认 2.0 秒，并对 IPC 偶发异常做有限重试。"""

    def emit(p, m):
        if progress_cb:
            try:
                progress_cb(p, m)
            except Exception:
                pass

    def cancelled() -> bool:
        if cancel_flag is None:
            return False
        if callable(cancel_flag):
            return bool(cancel_flag())
        if hasattr(cancel_flag, "is_set"):
            return cancel_flag.is_set()
        return bool(cancel_flag)

    def is_running(t) -> bool:
        last = None
        for _ in range(3):
            try:
                return bool(t.IsRunning)
            except Exception as e:
                last = e
                time.sleep(poll_interval)
        raise RuntimeError(f"读取运行状态失败(IPC): {last}")

    tol = zos_system.Tools.OpenTolerancing()
    try:
        emit(0, "配置公差工具…")
        _si, warns = configure(tol, spec)
        for w in warns:
            emit(0, "提示：" + w)
        emit(0, "Zemax公差工具设置：\n" + "\n".join(_tool_setting_lines(tol)))
        emit(1, f"开始运行：{spec.num_runs} 次蒙特卡洛")
        try:
            tol.Run()
        except Exception as e:
            lines = "\n".join(_tool_result_lines(tol))
            return RunResult(False, spec.ztd_path, spec.num_runs,
                             message=f"Zemax 公差工具启动失败：{type(e).__name__}: {e}\n工具状态：\n{lines}")

        start = time.time()
        running_notified = False
        while is_running(tol):
            if cancelled():
                if tol.CanCancel:
                    tol.Cancel()
                emit(0, "已取消")
                return RunResult(False, spec.ztd_path, spec.num_runs,
                                 message="用户取消")
            if timeout_s and (time.time() - start) > timeout_s:
                if tol.CanCancel:
                    tol.Cancel()
                return RunResult(False, spec.ztd_path, spec.num_runs,
                                 message="超时")
            if not running_notified:
                emit(50, "运行中…（请稍候，完成后会提示）")
                running_notified = True
            time.sleep(poll_interval)

        ok = bool(tol.Succeeded)
        result_lines = _tool_result_lines(tol)
        emit(100 if ok else 0, "Zemax公差工具结果：\n" + "\n".join(result_lines))
        bw_folder = ""
        try:
            bw_folder = str(tol.BestWorstOutputFolder or "")
        except Exception:
            pass

        actual_ztd = spec.ztd_path
        ztd_found = False
        ztd_candidates = []
        if spec.ztd_path:
            ztd_name = _ztd_file_name(spec.ztd_path)
            expected_ztd = _ztd_expected_path(spec.ztd_path)
            actual_ztd = expected_ztd
            ztd_candidates = [expected_ztd]
            if spec.lens_dir:
                ztd_candidates.append(os.path.join(spec.lens_dir, ztd_name))
            zemax_tol_dir = _zemax_tolerance_dir(zos_system)
            if zemax_tol_dir:
                ztd_candidates.append(os.path.join(zemax_tol_dir, ztd_name))
            tol_dir = os.path.join(
                os.path.expanduser("~"), "Documents", "Zemax", "Tolerance", ztd_name)
            ztd_candidates.append(tol_dir)
            seen = set()
            candidates = []
            for c in ztd_candidates:
                if not c:
                    continue
                key = os.path.normcase(os.path.abspath(c))
                if key in seen:
                    continue
                seen.add(key)
                candidates.append(c)
            ztd_candidates = candidates
            for c in ztd_candidates:
                if os.path.isfile(c):
                    actual_ztd = c
                    ztd_found = True
                    break

        zemax_status = "\n".join(result_lines)
        msg = "" if ok else (_tool_error_message(tol) + f"\nZemax工具状态：\n{zemax_status}")
        if ok and spec.ztd_path and not ztd_found:
            ok = False
            searched = "\n".join(f"  - {p}" for p in ztd_candidates)
            msg = (f"分析已运行，但未找到 ZTD 文件 "
                   f"{_ztd_file_name(spec.ztd_path)}。\n"
                   f"已检查以下路径：\n{searched}\n"
                   "ZTD 只能由 Zemax 保存到当前工作副本所在目录；"
                   "请确认工作副本已 SaveAs 到目标输出目录、目录可写，"
                   "且 OpticStudio 未弹出阻塞对话框。\n"
                   f"Zemax工具状态：\n{zemax_status}")

        emit(100, "完成" if ok else "失败")
        return RunResult(
            succeeded=ok,
            ztd_path=actual_ztd,
            num_runs=spec.num_runs,
            message=msg,
            bestworst_folder=bw_folder,
        )
    finally:
        try:
            tol.Close()
        except Exception:
            pass
