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
- ZTD 正确写法★：运行前 SaveTolDataFile=True + TolDataFile=绝对路径
  (os.path.abspath)。早期曾用纯文件名(os.path.basename)依赖 Zemax 相对
  目录解析，实测在 standalone + 跨盘(镜头在 F 盘)场景下失效：Zemax 自报
  Succeeded=True 却不落盘到任何标准目录，造成"假成功"。改绝对路径后稳定。
  绝不在运行后调 tol.Save(ztd)，那会写出损坏 ZTD。运行后按多候选路径
  (绝对路径/lens_dir/Documents\Zemax\Tolerance)回读实际生成的 ZTD；
  要求保存 ZTD 却回读不到时判为失败，不再"假成功"。
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
        tol.TolDataFile = _win_long_path(os.path.abspath(spec.ztd_path))

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
        emit(1, f"开始运行：{spec.num_runs} 次蒙特卡洛")
        tol.Run()

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
        bw_folder = ""
        try:
            bw_folder = str(tol.BestWorstOutputFolder or "")
        except Exception:
            pass

        actual_ztd = spec.ztd_path
        ztd_found = False
        if spec.ztd_path:
            ztd_name = os.path.basename(spec.ztd_path)
            candidates = [os.path.abspath(spec.ztd_path)]
            if spec.lens_dir:
                candidates.append(os.path.join(spec.lens_dir, ztd_name))
            tol_dir = os.path.join(
                os.path.expanduser("~"), "Documents", "Zemax", "Tolerance", ztd_name)
            candidates.append(tol_dir)
            seen = set()
            for c in candidates:
                if not c or c in seen:
                    continue
                seen.add(c)
                if os.path.isfile(c):
                    actual_ztd = c
                    ztd_found = True
                    break

        msg = "" if ok else (str(tol.ErrorMessage) or "运行未成功")
        if ok and spec.ztd_path and not ztd_found:
            ok = False
            searched = spec.lens_dir or os.path.dirname(spec.ztd_path)
            msg = (f"分析已运行，但未在 {searched} 找到 ZTD 文件 "
                   f"{os.path.basename(spec.ztd_path)}。"
                   "请确认 OpticStudio 未弹出阻塞对话框、目录可写，"
                   "或在公差分析对话框勾选「Save Tolerance Data File」。")

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
