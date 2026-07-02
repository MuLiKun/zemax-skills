r"""gui.py —— 公差分析图形界面（PySide6，暗色主题）。

在窗口里选 zmx / Excel 配置 / 输出目录与连接模式，点「开始分析」即可，
日志实时回显在下方文本框。业务逻辑全部复用 toltool.pipeline，
本文件只负责界面与线程调度，不重复任何公差/评价函数/TSC 逻辑。

连接模式：
  Standalone —— 程序后台自起 Zemax 实例，跑完自动释放（推荐）。
  GUI 模式   —— 连入已在 OpticStudio「交互扩展」等待的窗口，不关它。

运行：
    .\.venv\Scripts\python.exe -u tolerance_analysis\gui.py
"""

from __future__ import annotations

import ctypes
import json
import os
import sys
import tempfile

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

from PySide6 import QtCore, QtGui, QtWidgets

from toltool import pipeline
from toltool import standard_templates
from toltool import zos_connect


def _app_dir() -> str:
    """打包后返回 exe 所在目录，否则返回脚本所在目录。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return _HERE


def _default_config() -> str:
    """优先用 exe/脚本同级的模板配置作为默认，不存在则留空。"""
    for name in ("tol_config.xlsx", "tol_config_模板.xlsx"):
        p = os.path.join(_app_dir(), name)
        if os.path.isfile(p):
            return p
    return ""


DEFAULT_ZMX = ""
DEFAULT_CONFIG = _default_config()
DEFAULT_OUTDIR = _app_dir()
SETTINGS_ORG = os.environ.get("ZEMAX_TOL_SETTINGS_ORG", "ZemaxTools")
SETTINGS_APP = os.environ.get("ZEMAX_TOL_SETTINGS_APP", "ZemaxToleranceTool")


def _yes(v) -> bool:
    return str(v).strip().upper() in ("Y", "YES", "1", "TRUE", "是")


def _as_int(v, default: int) -> int:
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return default


def _resolve_ztd_config(ztd: str, fallback_config: str) -> tuple[str, str]:
    ztd_dir = os.path.dirname(os.path.abspath(ztd))
    used_excel = os.path.join(ztd_dir, "used_excel.xlsx")
    if os.path.isfile(used_excel):
        return used_excel, "ZTD 同目录 used_excel.xlsx"
    return fallback_config, "界面选择的 Excel"


def _apply_dark_titlebar(widget: QtWidgets.QWidget) -> None:
    """Windows 下把窗口标题栏改成深色（DWM 沉浸式暗色模式）。

    非 Windows 或调用失败时静默忽略，不影响界面其余部分。
    """
    if sys.platform != "win32":
        return
    try:
        hwnd = int(widget.winId())
        dwm = ctypes.windll.dwmapi
        value = ctypes.c_int(1)
        for attr in (20, 19):
            if dwm.DwmSetWindowAttribute(
                hwnd, attr, ctypes.byref(value), ctypes.sizeof(value)) == 0:
                break
    except Exception:
        pass


def _dark_message_box(parent, icon, title, text, buttons, default):
    """创建带深色标题栏的 QMessageBox 并弹出，返回用户点击的按钮。"""
    box = QtWidgets.QMessageBox(parent)
    box.setIcon(icon)
    box.setWindowTitle(title)
    box.setText(text)
    box.setStandardButtons(buttons)
    box.setDefaultButton(default)
    box.show()
    _apply_dark_titlebar(box)
    return box.exec()


def _apply_dark_palette(app: QtWidgets.QApplication) -> None:
    """暗色主题：Fusion 风格 + 深色 QPalette + 少量 QSS 微调。"""
    app.setStyle("Fusion")
    pal = QtGui.QPalette()
    bg = QtGui.QColor(30, 31, 34)
    base = QtGui.QColor(43, 45, 48)
    text = QtGui.QColor(220, 221, 222)
    disabled = QtGui.QColor(120, 121, 122)
    accent = QtGui.QColor(64, 132, 214)

    pal.setColor(QtGui.QPalette.Window, bg)
    pal.setColor(QtGui.QPalette.WindowText, text)
    pal.setColor(QtGui.QPalette.Base, base)
    pal.setColor(QtGui.QPalette.AlternateBase, bg)
    pal.setColor(QtGui.QPalette.ToolTipBase, base)
    pal.setColor(QtGui.QPalette.ToolTipText, text)
    pal.setColor(QtGui.QPalette.Text, text)
    pal.setColor(QtGui.QPalette.Button, base)
    pal.setColor(QtGui.QPalette.ButtonText, text)
    pal.setColor(QtGui.QPalette.BrightText, QtGui.QColor(255, 80, 80))
    pal.setColor(QtGui.QPalette.Highlight, accent)
    pal.setColor(QtGui.QPalette.HighlightedText, QtGui.QColor(255, 255, 255))
    pal.setColor(QtGui.QPalette.Link, accent)
    pal.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.Text, disabled)
    pal.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, disabled)
    pal.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText, disabled)
    app.setPalette(pal)

    app.setStyleSheet(
        """
        QWidget { font-size: 13px; }
        QPushButton {
            background-color: #3a3d41; border: 1px solid #4a4d51;
            border-radius: 4px; padding: 6px 14px;
        }
        QPushButton:hover { background-color: #45484d; }
        QPushButton:pressed { background-color: #2f3236; }
        QPushButton:disabled { color: #787878; border-color: #3a3d41; }
        QPushButton#run { background-color: #2f6fd0; border-color: #2f6fd0;
                          color: white; font-weight: bold; }
        QPushButton#run:hover { background-color: #3a7be0; }
        QPushButton#run:disabled { background-color: #3a3d41; color: #787878; }
        QLineEdit, QComboBox {
            background-color: #2b2d30; border: 1px solid #4a4d51;
            border-radius: 4px; padding: 5px 8px;
        }
        QComboBox QAbstractItemView {
            background-color: #2b2d30; selection-background-color: #2f6fd0;
        }
        QPlainTextEdit {
            background-color: #1b1c1e; border: 1px solid #3a3d41;
            border-radius: 4px; font-family: Consolas, "Courier New", monospace;
        }
        QLabel#hint { color: #9a9b9c; }
        """
    )


class _Worker(QtCore.QObject):
    """在子线程里跑 prepare_session + run_montecarlo，避免阻塞 UI。"""

    log = QtCore.Signal(str)
    finished = QtCore.Signal(bool, str)
    started = QtCore.Signal(int, int)
    need_zos_dir = QtCore.Signal(list)

    def __init__(self, zmx: str, config: str, outdir: str, connect: str,
                 standard_args: dict | None = None,
                 current_args: dict | None = None,
                 zos_dir: str | None = None):
        super().__init__()
        self._zmx = zmx
        self._config = config
        self._outdir = outdir
        self._connect = connect
        self._standard_args = standard_args or {}
        self._current_args = current_args or {}
        self._zos_dir = zos_dir
        self._cancel = False
        self._force = False
        self._sess = None

    def cancel(self) -> None:
        self._cancel = True

    def force_stop(self) -> None:
        """强制停止：置取消标志并立即关闭后台 Zemax 实例（仅 standalone）。"""
        self._cancel = True
        self._force = True
        sess = self._sess
        if sess is not None and self._connect == "standalone":
            try:
                sess.close()
            except Exception:
                pass

    @QtCore.Slot()
    def run(self) -> None:
        sess = None
        prep = None
        try:
            config = self._config
            if self._standard_args:
                config = standard_templates.make_temp_config(
                    self._zmx, self._outdir,
                    self._standard_args["template"],
                    self._standard_args["level"],
                    self._standard_args["num_runs"],
                    self._standard_args["num_to_save"],
                    self._standard_args["center_wave"],
                    self._standard_args["comp_mode"],
                    self._standard_args["save_worst_best"],
                    product_type=self._standard_args["product_type"])
                self.log.emit(f"标准模板配置: {config}")

            prep = pipeline.prepare_session(
                self._zmx, config, outdir=self._outdir,
                connect=self._connect, log=self.log.emit,
                zos_dir=self._zos_dir,
                use_current_settings=bool(self._current_args),
                current_args=self._current_args)
            sess = prep.sess
            self._sess = sess

            def run_log(message: str) -> None:
                pipeline.append_run_log(prep, message, log=self.log.emit)

            spec = pipeline.make_runspec(prep)
            pipeline.log_run_plan(
                prep, spec, log=run_log,
                export_stats=_yes(prep.rp.get("输出统计Excel", "N")))
            self.started.emit(int(spec.num_runs), int(spec.num_to_save))
            run_log(
                f"开始公差分析：{spec.num_runs} 次蒙特卡洛"
                f"（{spec.distribution}分布）…")

            from toltool import tol_runner
            result = tol_runner.run(
                prep.sess.sys, spec,
                progress_cb=lambda p, m: run_log(f"  [{p:>3}%] {m}"),
                cancel_flag=lambda: self._cancel)

            if self._force:
                self.finished.emit(False, "已强制停止（后台 Zemax 已关闭）")
                return
            if not result.succeeded:
                run_log("公差分析失败：" + (result.message or "未知错误"))
                self.finished.emit(False, result.message or "公差分析失败")
                return

            run_log(f"分析完成。ZTD: {result.ztd_path}")
            if result.bestworst_folder:
                run_log(f"Worst/Best 输出目录: {result.bestworst_folder}")

            if _yes(prep.rp.get("输出统计Excel", "N")):
                from toltool import ztd_reader
                report_meta = [
                    r for r in prep.cfg.report
                    if _yes(r.get("启用")) and r.get("标签")
                ]
                report_labels = [str(r.get("标签")).strip() for r in report_meta]
                num_items = len(report_labels) + 1 if report_labels else None
                comp_count = sum(
                    1 for r in (prep.tde_meta or [])
                    if str(r.get("操作数") or "").strip().upper() == "COMP")
                if num_items:
                    run_log(
                        f"ZTD 自动统计分项: {num_items + comp_count} 项（自定义脚本 + {len(report_labels)} 个 REPORT + {comp_count} 个 COMP）")
                num_runs = _as_int(prep.rp.get("蒙特卡洛次数"), int(spec.num_runs))
                run_log("正在读取 ZTD 并导出统计 Excel…")
                zres = ztd_reader.read_ztd(
                    prep.sess.sys, result.ztd_path, num_runs=num_runs,
                    report_labels=report_labels or None,
                    num_items=num_items,
                    report_meta=report_meta or None,
                    tde_meta=prep.tde_meta or None)
                if not zres.succeeded:
                    run_log("分析完成，但读取 ZTD 失败：" + zres.message)
                    self.finished.emit(False, "分析完成，但读取 ZTD 失败：" + zres.message)
                    return
                if zres.message:
                    run_log("提示：" + zres.message)
                stat_path = result.ztd_path.rsplit(".", 1)[0] + "_统计.xlsx"
                out = ztd_reader.export_excel(zres, stat_path)
                run_log(f"统计 Excel: {out}")

            self.finished.emit(True, result.ztd_path)
        except zos_connect.ZosDirNotFound as e:
            self.need_zos_dir.emit(list(e.searched))
        except Exception as e:
            msg = f"{type(e).__name__}: {e}"
            if prep is not None:
                pipeline.append_run_log(prep, "任务异常：" + msg, log=lambda _m: None)
            if self._force:
                self.finished.emit(False, "已强制停止（后台 Zemax 已关闭）")
            else:
                self.finished.emit(False, msg)
        finally:
            if sess is not None and not self._force \
                    and self._connect == "standalone":
                try:
                    sess.close()
                    self.log.emit("已释放 Zemax 独立实例。")
                except Exception as e:
                    self.log.emit(f"释放 Zemax 独立实例时出错（已忽略）: {type(e).__name__}: {e}")
            self._sess = None


class _ZtdWorker(QtCore.QObject):

    log = QtCore.Signal(str)
    finished = QtCore.Signal(bool, str)
    need_zos_dir = QtCore.Signal(list)

    def __init__(self, ztd: str, config: str, connect: str,
                 zos_dir: str | None = None):
        super().__init__()
        self._ztd = ztd
        self._config = config
        self._connect = connect
        self._zos_dir = zos_dir
        self._sess = None

    def force_stop(self) -> None:
        sess = self._sess
        if sess is not None and self._connect == "standalone":
            try:
                sess.close()
            except Exception:
                pass

    @QtCore.Slot()
    def run(self) -> None:
        sess = None
        try:
            from toltool import excel_io, ztd_reader

            config_path, config_source = _resolve_ztd_config(self._ztd, self._config)
            self.log.emit(f"读取配置 Excel: {config_path}")
            self.log.emit(f"配置来源: {config_source}")
            if config_path != self._config:
                self.log.emit(f"已忽略界面选择的 Excel: {self._config}")
            cfg = excel_io.read_config(config_path)
            report_meta = [
                r for r in cfg.report if _yes(r.get("启用")) and r.get("标签")
            ]
            report_labels = [str(r.get("标签")).strip() for r in report_meta]
            tde_meta = None
            run_config = os.path.join(os.path.dirname(os.path.abspath(self._ztd)), "run_config.json")
            if os.path.isfile(run_config):
                try:
                    with open(run_config, "r", encoding="utf-8") as f:
                        tde_meta = json.load(f).get("tde_meta") or None
                except Exception:
                    tde_meta = None
            num_runs = _as_int(cfg.run_params.get("蒙特卡洛次数"), 0)
            if num_runs <= 0:
                self.finished.emit(False, "配置中的『蒙特卡洛次数』无效，无法确定 ZTD 数据行数。")
                return

            self.log.emit(f"连接模式: {'Standalone' if self._connect == 'standalone' else 'GUI(交互扩展)'}")
            sess = zos_connect.ZosSession(zos_dir=self._zos_dir)
            self._sess = sess
            self.log.emit(f"ZOS 目录: {sess.zos_dir}")
            sess.connect(mode=self._connect)
            if self._connect == "standalone":
                self.log.emit("已启动 Zemax 独立实例（用于读取已有 ZTD）")
            else:
                self.log.emit(f"已连接交互扩展: {sess.sys.SystemFile}")

            num_items = len(report_labels) + 1 if report_labels else None
            self.log.emit(f"正在读取 ZTD: {self._ztd}")
            zres = ztd_reader.read_ztd(
                sess.sys, self._ztd, num_runs=num_runs,
                report_labels=report_labels or None,
                num_items=num_items,
                report_meta=report_meta or None,
                tde_meta=tde_meta)
            if not zres.succeeded:
                self.finished.emit(False, zres.message or "读取 ZTD 失败")
                return
            if zres.message:
                self.log.emit("提示：" + zres.message)

            ztd_dir = os.path.dirname(os.path.abspath(self._ztd))
            base = os.path.splitext(os.path.basename(self._ztd))[0]
            stat_path = os.path.join(ztd_dir, base + "_统计.xlsx")
            out = ztd_reader.export_excel(zres, stat_path)
            self.log.emit(f"统计 Excel: {out}")
            self.finished.emit(True, out)
        except zos_connect.ZosDirNotFound as e:
            self.need_zos_dir.emit(list(e.searched))
        except Exception as e:
            self.finished.emit(False, f"{type(e).__name__}: {e}")
        finally:
            if sess is not None and self._connect == "standalone":
                try:
                    sess.close()
                    self.log.emit("已释放 Zemax 独立实例。")
                except Exception as e:
                    self.log.emit(f"释放 Zemax 独立实例时出错（已忽略）: {type(e).__name__}: {e}")
            self._sess = None


class _FieldPreviewWorker(QtCore.QObject):

    log = QtCore.Signal(str)
    finished = QtCore.Signal(bool, str)
    need_zos_dir = QtCore.Signal(list)

    def __init__(self, zmx: str, run_params: dict, connect: str,
                 zos_dir: str | None = None):
        super().__init__()
        self._zmx = zmx
        self._run_params = run_params
        self._connect = connect
        self._zos_dir = zos_dir
        self._sess = None

    def force_stop(self) -> None:
        sess = self._sess
        if sess is not None and self._connect == "standalone":
            try:
                sess.close()
            except Exception:
                pass

    @QtCore.Slot()
    def run(self) -> None:
        sess = None
        tmp_dir = None
        try:
            from toltool import field_mapping

            self.log.emit(f"连接模式: {'Standalone' if self._connect == 'standalone' else 'GUI(交互扩展)'}")
            sess = zos_connect.ZosSession(zos_dir=self._zos_dir)
            self._sess = sess
            self.log.emit(f"ZOS 目录: {sess.zos_dir}")
            sess.connect(mode=self._connect)
            if self._connect == "standalone":
                self.log.emit("已启动 Zemax 独立实例（用于预览视场映射）")
                tmp_dir = tempfile.TemporaryDirectory(prefix="zemax_field_preview_")
                copy_path = os.path.join(tmp_dir.name, os.path.basename(self._zmx))
                self.log.emit(f"预览工作副本: {sess.open_as_copy(self._zmx, copy_path=copy_path)}")
            else:
                self.log.emit(f"已连接交互扩展: {sess.sys.SystemFile}")
                self.log.emit(
                    "提示：交互扩展模式的预览基于 OpticStudio 当前打开的文件，"
                    "而非界面所选 zmx；如需预览指定文件请改用 Standalone 模式。")
            result = field_mapping.preview(
                sess.sys, self._run_params,
                simulate_insert=self._connect == "standalone")
            if not result.enabled:
                self.log.emit("视场映射：未启用")
            else:
                for line in pipeline.field_mapping_report_lines(result):
                    if line:
                        self.log.emit(line)
            self.finished.emit(True, "视场映射预览完成")
        except zos_connect.ZosDirNotFound as e:
            self.need_zos_dir.emit(list(e.searched))
        except Exception as e:
            self.finished.emit(False, f"{type(e).__name__}: {e}")
        finally:
            if sess is not None and self._connect == "standalone":
                try:
                    sess.close()
                    self.log.emit("已释放 Zemax 独立实例。")
                except Exception as e:
                    self.log.emit(f"释放 Zemax 独立实例时出错（已忽略）: {type(e).__name__}: {e}")
            if tmp_dir is not None:
                tmp_dir.cleanup()
            self._sess = None


class MainWindow(QtWidgets.QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Zemax 公差分析")
        self.resize(720, 560)
        self._thread = None
        self._worker = None
        self._run_args = None
        self._ztd_args = None
        self._field_preview_args = None
        self._active_task = ""
        self._settings = QtCore.QSettings(SETTINGS_ORG, SETTINGS_APP)

        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        root = QtWidgets.QVBoxLayout(central)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(10)

        form = QtWidgets.QGridLayout()
        form.setHorizontalSpacing(8)
        form.setVerticalSpacing(8)
        root.addLayout(form)

        self.ed_zmx = QtWidgets.QLineEdit(
            self._setting("zmx", DEFAULT_ZMX))
        self.ed_config = QtWidgets.QLineEdit(
            self._setting("config", DEFAULT_CONFIG))
        self.ed_outdir = QtWidgets.QLineEdit(
            self._setting("outdir", DEFAULT_OUTDIR))
        self._add_file_row(form, 0, "待分析 zmx：", self.ed_zmx,
                           self._pick_zmx)
        self.btn_pick_config = self._add_file_row(form, 1, "Excel 配置：", self.ed_config,
                                                  self._pick_config)
        self._add_file_row(form, 2, "输出目录：", self.ed_outdir,
                           self._pick_outdir)

        form.addWidget(QtWidgets.QLabel("分析模式："), 3, 0)
        self.cb_analysis_mode = QtWidgets.QComboBox()
        self.cb_analysis_mode.addItem("高级 Excel 配置", "excel")
        self.cb_analysis_mode.addItem("普通标准模板", "standard")
        self.cb_analysis_mode.addItem("使用 Zemax 当前设置", "current")
        mode = self.cb_analysis_mode.findData(self._setting("analysis_mode", "excel"))
        if mode >= 0:
            self.cb_analysis_mode.setCurrentIndex(mode)
        self.cb_analysis_mode.currentIndexChanged.connect(self._on_analysis_mode_changed)
        form.addWidget(self.cb_analysis_mode, 3, 1, 1, 2)

        self.standard_panel = QtWidgets.QWidget()
        std = QtWidgets.QHBoxLayout(self.standard_panel)
        std.setContentsMargins(0, 0, 0, 0)
        std.setSpacing(6)
        self.cb_product_type = QtWidgets.QComboBox()
        self.cb_product_type.addItems(standard_templates.product_types())
        idx = self.cb_product_type.findText(self._setting("product_type", standard_templates.DEFAULT_PRODUCT_TYPE))
        if idx >= 0:
            self.cb_product_type.setCurrentIndex(idx)
        self.cb_std_template = QtWidgets.QComboBox()
        self._refresh_standard_templates(self._setting("standard_template", standard_templates.DEFAULT_TEMPLATE_NAME))
        self.cb_product_type.currentIndexChanged.connect(self._on_product_type_changed)
        self.cb_std_template.currentIndexChanged.connect(self._update_template_tooltip)
        self.cb_tol_level = QtWidgets.QComboBox()
        self.cb_tol_level.addItems(standard_templates.LEVEL_NAMES)
        idx = self.cb_tol_level.findText(self._setting("tolerance_level", "标准"))
        if idx >= 0:
            self.cb_tol_level.setCurrentIndex(idx)
        self.sp_runs = QtWidgets.QSpinBox()
        self.sp_runs.setRange(1, 100000)
        self.sp_runs.setValue(int(self._setting("standard_runs", "20")))
        self.sp_save = QtWidgets.QSpinBox()
        self.sp_save.setRange(0, 100000)
        self.sp_save.setValue(int(self._setting("standard_save", "0")))
        self.cb_comp = QtWidgets.QComboBox()
        self.cb_comp.addItems(["无", "全部优化DLS", "全部优化OD"])
        idx = self.cb_comp.findText(self._setting("standard_comp", "无"))
        if idx >= 0:
            self.cb_comp.setCurrentIndex(idx)
        self.chk_save_worst_best = QtWidgets.QCheckBox("保存 WC/BC")
        self.chk_save_worst_best.setChecked(_yes(self._setting("standard_save_worst_best", "N")))
        self.lb_product_type = QtWidgets.QLabel("产品")
        self.lb_std_template = QtWidgets.QLabel("模板")
        self.lb_tol_level = QtWidgets.QLabel("等级")
        self.lb_runs = QtWidgets.QLabel("MC")
        self.lb_save = QtWidgets.QLabel("保存")
        self.lb_comp = QtWidgets.QLabel("补偿")
        std.addWidget(self.lb_product_type)
        std.addWidget(self.cb_product_type)
        std.addWidget(self.lb_std_template)
        std.addWidget(self.cb_std_template, 2)
        std.addWidget(self.lb_tol_level)
        std.addWidget(self.cb_tol_level)
        std.addWidget(self.lb_runs)
        std.addWidget(self.sp_runs)
        std.addWidget(self.lb_save)
        std.addWidget(self.sp_save)
        std.addWidget(self.lb_comp)
        std.addWidget(self.cb_comp)
        std.addWidget(self.chk_save_worst_best)
        self.lb_standard_panel = QtWidgets.QLabel("标准模板：")
        form.addWidget(self.lb_standard_panel, 4, 0)
        form.addWidget(self.standard_panel, 4, 1, 1, 2)

        form.addWidget(QtWidgets.QLabel("连接模式："), 5, 0)
        self.cb_mode = QtWidgets.QComboBox()
        self.cb_mode.addItem("Standalone（程序后台挂起，推荐）", "standalone")
        self.cb_mode.addItem("GUI 模式（连入交互扩展窗口）", "extension")
        mode_index = self.cb_mode.findData(self._setting("connect", "standalone"))
        if mode_index >= 0:
            self.cb_mode.setCurrentIndex(mode_index)
        self.cb_mode.currentIndexChanged.connect(
            lambda: self._settings.setValue(
                "connect", self.cb_mode.currentData()))
        form.addWidget(self.cb_mode, 5, 1, 1, 2)

        self.ed_ztd = QtWidgets.QLineEdit(self._setting("ztd", ""))
        self.btn_pick_ztd = self._add_file_row(form, 6, "已有 ZTD：", self.ed_ztd,
                                               self._pick_ztd)

        self.ztd_panel = QtWidgets.QWidget()
        ztdbar = QtWidgets.QHBoxLayout(self.ztd_panel)
        ztdbar.setContentsMargins(0, 0, 0, 0)
        ztdbar.addStretch(1)
        self.btn_ztd = QtWidgets.QPushButton("分析已有 ZTD")
        self.btn_ztd.clicked.connect(self._on_analyze_ztd)
        ztdbar.addWidget(self.btn_ztd)
        root.addWidget(self.ztd_panel)

        btns = QtWidgets.QHBoxLayout()
        btns.addStretch(1)
        self.btn_export_std = QtWidgets.QPushButton("导出标准配置")
        self.btn_export_std.clicked.connect(self._on_export_standard_config)
        self.btn_check_config = QtWidgets.QPushButton("检查配置")
        self.btn_check_config.clicked.connect(self._on_check_config)
        self.btn_preview_fields = QtWidgets.QPushButton("预览视场映射")
        self.btn_preview_fields.clicked.connect(self._on_preview_field_mapping)
        self.btn_open_result = QtWidgets.QPushButton("打开结果目录")
        self.btn_open_result.clicked.connect(self._on_open_result_dir)
        self.btn_run = QtWidgets.QPushButton("开始分析")
        self.btn_run.setObjectName("run")
        self.btn_run.clicked.connect(self._on_run)
        self.btn_cancel = QtWidgets.QPushButton("取消")
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.clicked.connect(self._on_cancel)
        btns.addWidget(self.btn_export_std)
        btns.addWidget(self.btn_check_config)
        btns.addWidget(self.btn_preview_fields)
        btns.addWidget(self.btn_open_result)
        btns.addWidget(self.btn_cancel)
        btns.addWidget(self.btn_run)
        root.addLayout(btns)
        self._on_analysis_mode_changed()

        hint = QtWidgets.QLabel(
            "提示：GUI 模式需先在 OpticStudio 进入「编程 → 交互扩展」并等待。")
        hint.setObjectName("hint")
        root.addWidget(hint)

        logbar = QtWidgets.QHBoxLayout()
        logbar.addWidget(QtWidgets.QLabel("日志："))
        logbar.addStretch(1)
        self.btn_copy = QtWidgets.QPushButton("复制日志")
        self.btn_copy.clicked.connect(self._on_copy_log)
        self.btn_clear = QtWidgets.QPushButton("清空日志")
        self.btn_clear.clicked.connect(self._on_clear_log)
        logbar.addWidget(self.btn_copy)
        logbar.addWidget(self.btn_clear)
        root.addLayout(logbar)

        self.log_view = QtWidgets.QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setMaximumBlockCount(5000)
        root.addWidget(self.log_view, 1)

        self.lbl_runinfo = QtWidgets.QLabel("")
        self.lbl_elapsed = QtWidgets.QLabel("")
        self.statusBar().addPermanentWidget(self.lbl_elapsed)
        self.statusBar().addPermanentWidget(self.lbl_runinfo)

        self._num_runs = 0
        self._num_to_save = 0
        self._elapsed = 0
        self._timer = QtCore.QTimer(self)
        self._timer.setInterval(1000)
        self._timer.timeout.connect(self._on_tick)

        self.statusBar().showMessage("就绪")

    def _add_file_row(self, grid, row, label, edit, slot):
        lb = QtWidgets.QLabel(label)
        grid.addWidget(lb, row, 0)
        grid.addWidget(edit, row, 1)
        btn = QtWidgets.QPushButton("浏览…")
        btn.clicked.connect(slot)
        grid.addWidget(btn, row, 2)
        btn._row_label = lb
        btn._row_edit = edit
        return btn

    def _set_file_row_visible(self, btn: QtWidgets.QPushButton, visible: bool) -> None:
        btn.setVisible(visible)
        getattr(btn, "_row_label").setVisible(visible)
        getattr(btn, "_row_edit").setVisible(visible)

    def _refresh_standard_templates(self, preferred: str = "") -> None:
        product_type = self.cb_product_type.currentText() or standard_templates.DEFAULT_PRODUCT_TYPE
        names = standard_templates.template_names(product_type)
        self.cb_std_template.blockSignals(True)
        self.cb_std_template.clear()
        self.cb_std_template.addItems(names)
        idx = self.cb_std_template.findText(preferred or standard_templates.DEFAULT_TEMPLATE_NAME)
        if idx < 0:
            idx = 0
        if idx >= 0:
            self.cb_std_template.setCurrentIndex(idx)
        self.cb_std_template.blockSignals(False)
        self._update_template_tooltip()

    def _update_template_tooltip(self) -> None:
        product_type = self.cb_product_type.currentText() or standard_templates.DEFAULT_PRODUCT_TYPE
        template = self.cb_std_template.currentText() or standard_templates.DEFAULT_TEMPLATE_NAME
        try:
            tip = standard_templates.template_description(template, product_type)
        except Exception:
            tip = ""
        self.cb_std_template.setToolTip(tip)
        if hasattr(self, "lb_std_template"):
            self.lb_std_template.setToolTip(tip)

    def _on_product_type_changed(self) -> None:
        self._refresh_standard_templates(self.cb_std_template.currentText())

    def _setting(self, key: str, default: str = "") -> str:
        return str(self._settings.value(key, default) or "")

    def _remember_path(self, key: str, path: str) -> None:
        if path:
            self._settings.setValue(key, path)
            if os.path.isfile(path):
                d = os.path.dirname(path)
            elif os.path.isdir(path):
                d = path
            else:
                d = ""
            if d:
                self._settings.setValue("last_dir", d)

    def _dialog_dir(self, path: str, fallback: str = "") -> str:
        path = str(path or "").strip()
        if os.path.isfile(path):
            return os.path.dirname(path)
        if os.path.isdir(path):
            return path
        fallback = str(fallback or "").strip()
        if fallback and os.path.isdir(fallback):
            return fallback
        last_dir = self._setting("last_dir", "")
        if last_dir and os.path.isdir(last_dir):
            return last_dir
        return _app_dir()

    def _remember_form_paths(self) -> None:
        self._remember_path("zmx", self.ed_zmx.text().strip())
        self._remember_path("config", self.ed_config.text().strip())
        self._remember_path("outdir", self.ed_outdir.text().strip())
        self._remember_path("ztd", self.ed_ztd.text().strip())
        self._settings.setValue("connect", self.cb_mode.currentData())
        self._settings.setValue("analysis_mode", self.cb_analysis_mode.currentData())
        self._settings.setValue("product_type", self.cb_product_type.currentText())
        self._settings.setValue("standard_template", self.cb_std_template.currentText())
        self._settings.setValue("tolerance_level", self.cb_tol_level.currentText())
        self._settings.setValue("standard_runs", self.sp_runs.value())
        self._settings.setValue("standard_save", self.sp_save.value())
        self._settings.setValue("standard_comp", self.cb_comp.currentText())
        self._settings.setValue(
            "standard_save_worst_best",
            "Y" if self.chk_save_worst_best.isChecked() else "N")

    def _on_analysis_mode_changed(self):
        mode = self.cb_analysis_mode.currentData()
        use_excel = mode == "excel"
        use_standard = mode == "standard"
        use_current = mode == "current"
        if hasattr(self, "btn_pick_config"):
            self._set_file_row_visible(self.btn_pick_config, use_excel)
            # config 行仅高级 Excel 模式可见；可见时确保恢复可用（运行态可能临时禁用过）。
            self.ed_config.setEnabled(use_excel)
        if hasattr(self, "btn_export_std"):
            self.btn_export_std.setVisible(use_standard)
            self.btn_export_std.setEnabled(use_standard)
        if hasattr(self, "btn_check_config"):
            self.btn_check_config.setVisible(use_excel)
            self.btn_check_config.setEnabled(use_excel)
        if hasattr(self, "btn_preview_fields"):
            self.btn_preview_fields.setVisible(not use_current)
            self.btn_preview_fields.setEnabled(not use_current)
        self.standard_panel.setVisible(use_standard or use_current)
        self.standard_panel.setEnabled(use_standard or use_current)
        self.cb_product_type.setVisible(use_standard)
        self.lb_product_type.setVisible(use_standard)
        self.cb_std_template.setVisible(use_standard)
        self.lb_std_template.setVisible(use_standard)
        self.cb_tol_level.setVisible(use_standard)
        self.lb_tol_level.setVisible(use_standard)
        self.cb_product_type.setEnabled(use_standard)
        self.cb_std_template.setEnabled(use_standard)
        self.cb_tol_level.setEnabled(use_standard)
        if hasattr(self, "lb_standard_panel"):
            self.lb_standard_panel.setVisible(use_standard or use_current)
            self.lb_standard_panel.setText("标准模板：" if use_standard else "运行参数：")
        self._settings.setValue("analysis_mode", mode)

    def _pick_zmx(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "选择 zmx 文件", self._dialog_dir(self.ed_zmx.text()),
            "Zemax 镜头 (*.zmx *.zos);;所有文件 (*.*)")
        if path:
            self.ed_zmx.setText(path)
            self._remember_path("zmx", path)

    def _pick_config(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "选择 Excel 配置", self._dialog_dir(self.ed_config.text()),
            "Excel (*.xlsx);;所有文件 (*.*)")
        if path:
            self.ed_config.setText(path)
            self._remember_path("config", path)

    def _pick_outdir(self):
        path = QtWidgets.QFileDialog.getExistingDirectory(
            self, "选择输出目录", self._dialog_dir(self.ed_outdir.text()))
        if path:
            self.ed_outdir.setText(path)
            self._remember_path("outdir", path)

    def _pick_ztd(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "选择已有 ZTD 文件",
            self._dialog_dir(self.ed_ztd.text(), self.ed_outdir.text()),
            "Zemax 公差数据 (*.ztd *.ZTD);;所有文件 (*.*)")
        if path:
            self.ed_ztd.setText(path)
            self._remember_path("ztd", path)

    def _append_log(self, text: str):
        self.log_view.appendPlainText(text)

    def _standard_args_from_ui(self) -> dict:
        return {
            "product_type": self.cb_product_type.currentText(),
            "template": self.cb_std_template.currentText(),
            "level": self.cb_tol_level.currentText(),
            "num_runs": self.sp_runs.value(),
            "num_to_save": self.sp_save.value(),
            "center_wave": 0,
            "comp_mode": self.cb_comp.currentText(),
            "save_worst_best": self.chk_save_worst_best.isChecked(),
        }

    def _current_args_from_ui(self) -> dict:
        args = self._standard_args_from_ui()
        args["report_filter"] = "all"
        return args

    def _config_from_current_form(self):
        zmx = self.ed_zmx.text().strip()
        config = self.ed_config.text().strip()
        mode = self.cb_analysis_mode.currentData()
        use_standard = mode == "standard"
        use_current = mode == "current"
        if not os.path.isfile(zmx):
            raise ValueError("运行前校验失败：\n- ZMX 文件不存在：" + zmx)
        if use_current:
            return None
        if use_standard:
            if self.sp_save.value() > self.sp_runs.value():
                raise ValueError("运行前校验失败：\n- 保存数量不能大于 MC 次数")
            cfg = standard_templates.build_config(
                zmx,
                template=self.cb_std_template.currentText(),
                level=self.cb_tol_level.currentText(),
                num_runs=self.sp_runs.value(),
                num_to_save=self.sp_save.value(),
                center_wave=0,
                comp_mode=self.cb_comp.currentText(),
                save_worst_best=self.chk_save_worst_best.isChecked(),
                product_type=self.cb_product_type.currentText())
            return pipeline.validate_config_data(cfg)
        return pipeline.validate_config(zmx, config)

    def _validate_current_form(self) -> list[str]:
        mode = self.cb_analysis_mode.currentData()
        if mode == "current":
            zmx = self.ed_zmx.text().strip()
            if not os.path.isfile(zmx):
                raise ValueError("运行前校验失败：\n- ZMX 文件不存在：" + zmx)
            return [
                "当前设置模式：已校验 zmx 路径；MFE/TDE 需连接 Zemax 后运行期读取。",
            ]
        cfg = self._config_from_current_form()
        if mode == "standard":
            return [
                "标准模板配置校验通过。",
                f"评价函数有效行数: {len([r for r in cfg.mfe if str(r.get('操作数') or '').strip()])}",
                f"REPORT 启用行数: {len([r for r in cfg.report if _yes(r.get('启用'))])}",
            ]
        return [
            "Excel 配置校验通过。",
            f"配置 Excel: {self.ed_config.text().strip()}",
            f"评价函数有效行数: {len([r for r in cfg.mfe if str(r.get('操作数') or '').strip()])}",
            f"REPORT 启用行数: {len([r for r in cfg.report if _yes(r.get('启用'))])}",
        ]

    def _on_check_config(self):
        if self._thread is not None and self._thread.isRunning():
            self._warn("当前已有任务正在运行，请等待结束后再检查配置。")
            return
        try:
            lines = self._validate_current_form()
        except Exception as e:
            self._append_log("❌ 配置检查失败：" + str(e))
            self.statusBar().showMessage("配置检查失败", 3000)
            self._warn("配置检查失败：\n" + str(e))
            return
        self._remember_form_paths()
        self.log_view.clear()
        self._append_log("配置检查通过。")
        for line in lines:
            self._append_log(line)
        self.statusBar().showMessage("配置检查通过", 3000)

    def _on_preview_field_mapping(self):
        if self._thread is not None and self._thread.isRunning():
            self._warn("当前已有任务正在运行，请等待结束后再预览视场映射。")
            return
        zmx = self.ed_zmx.text().strip()
        connect = self.cb_mode.currentData()
        mode = self.cb_analysis_mode.currentData()
        if mode == "current":
            self._warn("当前设置模式的视场/MFE/REPORT 会在连接 Zemax 后按当前文件读取，暂不支持单独预览视场映射。")
            return
        try:
            cfg = self._config_from_current_form()
        except Exception as e:
            self._append_log("❌ 视场映射预览前检查失败：" + str(e))
            self.statusBar().showMessage("预览前检查失败", 3000)
            self._warn("视场映射预览前检查失败：\n" + str(e))
            return
        self._remember_form_paths()
        self.log_view.clear()
        self._append_log("开始预览视场映射。")
        self._append_log(f"待分析镜头: {zmx}")
        self._append_log(f"分析模式: {'普通标准模板' if mode == 'standard' else '高级 Excel 配置'}")
        if mode == "excel":
            self._append_log(f"配置 Excel: {self.ed_config.text().strip()}")
        # run_params 仅含标量（数字/字符串），dict() 浅拷贝即可安全跨线程传递。
        self._field_preview_args = (zmx, dict(cfg.run_params), connect)
        self._active_task = "field_preview"
        self._start_field_preview_worker(None)

    def _on_export_standard_config(self):
        if self.cb_analysis_mode.currentData() != "standard":
            self._warn("请先将分析模式切换为“普通标准模板”。")
            return
        zmx = self.ed_zmx.text().strip()
        outdir = self.ed_outdir.text().strip()
        if not os.path.isfile(zmx):
            self._warn("zmx 文件不存在：\n" + zmx)
            return
        if self.sp_save.value() > self.sp_runs.value():
            self._warn("标准模板模式下，保存数量不能大于 MC 次数。")
            return
        if not outdir:
            outdir = os.path.join(_app_dir(), "output")
            self.ed_outdir.setText(outdir)
        os.makedirs(outdir, exist_ok=True)
        args = self._standard_args_from_ui()
        default_path = standard_templates.default_config_path(zmx, outdir)
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "导出标准配置", default_path,
            "Excel (*.xlsx);;所有文件 (*.*)")
        if not path:
            return
        if os.path.splitext(path)[1].lower() != ".xlsx":
            path += ".xlsx"
        try:
            cfg = standard_templates.build_config(
                zmx, template=args["template"], level=args["level"],
                num_runs=args["num_runs"], num_to_save=args["num_to_save"],
                center_wave=args["center_wave"], comp_mode=args["comp_mode"],
                save_worst_best=args["save_worst_best"],
                product_type=args["product_type"])
            standard_templates.write_config_excel(path, cfg, overwrite=True)
        except Exception as e:
            self._warn("导出标准配置失败：\n" + str(e))
            return
        self.ed_config.setText(path)
        self._remember_path("config", path)
        self._remember_form_paths()
        self._append_log(f"已导出标准配置: {path}")
        self.statusBar().showMessage("标准配置已导出", 3000)

    def _on_run(self):
        zmx = self.ed_zmx.text().strip()
        config = self.ed_config.text().strip()
        outdir = self.ed_outdir.text().strip()
        connect = self.cb_mode.currentData()
        mode = self.cb_analysis_mode.currentData()
        use_standard = mode == "standard"
        use_current = mode == "current"

        if not os.path.isfile(zmx):
            self._warn("zmx 文件不存在：\n" + zmx)
            return
        if not (use_standard or use_current) and not os.path.isfile(config):
            self._warn("Excel 配置不存在：\n" + config)
            return
        if not outdir:
            outdir = os.path.join(_app_dir(), "output")
        os.makedirs(outdir, exist_ok=True)
        self.ed_outdir.setText(outdir)
        if (use_standard or use_current) and self.sp_save.value() > self.sp_runs.value():
            self._warn("保存数量不能大于 MC 次数。")
            return
        self._remember_form_paths()

        standard_args = self._standard_args_from_ui() if use_standard else None
        current_args = self._current_args_from_ui() if use_current else None

        self.log_view.clear()
        self._append_log(
            f"连接模式: {'Standalone' if connect == 'standalone' else 'GUI(交互扩展)'}")
        self._append_log(f"待分析镜头: {zmx}")
        if use_standard:
            self._append_log("分析模式: 普通标准模板")
            self._append_log(
                f"产品={standard_args['product_type']} 模板={standard_args['template']} 等级={standard_args['level']} "
                f"MC={standard_args['num_runs']} 保存={standard_args['num_to_save']} "
                f"补偿={standard_args['comp_mode']} "
                f"保存WC/BC={'Y' if standard_args['save_worst_best'] else 'N'}")
        elif use_current:
            self._append_log("分析模式: 使用 Zemax 当前设置")
            self._append_log(
                f"MC={current_args['num_runs']} 保存={current_args['num_to_save']} "
                f"补偿={current_args['comp_mode']} "
                f"保存WC/BC={'Y' if current_args['save_worst_best'] else 'N'}")
        else:
            self._append_log("分析模式: 高级 Excel 配置")
            self._append_log(f"配置 Excel: {config}")
        self._append_log(f"输出目录: {outdir}")

        self._run_args = (zmx, config, outdir, connect, standard_args, current_args)
        self._active_task = "tol"
        self._start_worker(None)

    def _on_analyze_ztd(self):
        if self._thread is not None and self._thread.isRunning():
            self._warn("当前已有任务正在运行，请等待结束后再分析 ZTD。")
            return

        ztd = self.ed_ztd.text().strip()
        config = self.ed_config.text().strip()
        connect = self.cb_mode.currentData()

        if not os.path.isfile(ztd):
            self._warn("ZTD 文件不存在：\n" + ztd)
            return
        if os.path.splitext(ztd)[1].lower() != ".ztd":
            self._warn("请选择 .ZTD 公差数据文件：\n" + ztd)
            return
        config_path, config_source = _resolve_ztd_config(ztd, config)
        if not os.path.isfile(config_path):
            self._warn("Excel 配置不存在：\n" + config_path)
            return
        self._remember_form_paths()

        self.log_view.clear()
        self._append_log("开始独立分析已有 ZTD。")
        self._append_log(f"ZTD 文件: {ztd}")
        self._append_log(f"配置 Excel: {config_path}")
        self._append_log(f"配置来源: {config_source}")
        if config_path != config:
            self._append_log(f"已忽略界面选择的 Excel: {config}")
        self._append_log(f"统计输出目录: {os.path.dirname(os.path.abspath(ztd))}")

        self._ztd_args = (ztd, config_path, connect)
        self._active_task = "ztd"
        self._start_ztd_worker(None)

    def _start_worker(self, zos_dir):
        if not self._run_args:
            self._warn("尚未设置运行参数，请先点击「开始分析」。")
            return
        zmx, config, outdir, connect, standard_args, current_args = self._run_args
        self._set_running(True)
        self._num_runs = 0
        self._num_to_save = 0
        self._elapsed = 0
        self.lbl_elapsed.setText("已用时间 00:00")
        self.lbl_runinfo.setText("")
        self._timer.start()
        self._thread = QtCore.QThread(self)
        self._worker = _Worker(
            zmx, config, outdir, connect,
            standard_args=standard_args,
            current_args=current_args,
            zos_dir=zos_dir)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.log.connect(self._append_log)
        self._worker.started.connect(self._on_started)
        self._worker.finished.connect(self._on_finished)
        self._worker.need_zos_dir.connect(self._on_need_zos_dir)
        self._thread.start()

    def _start_ztd_worker(self, zos_dir):
        if not self._ztd_args:
            self._warn("尚未设置 ZTD 分析参数，请先选择已有 ZTD。")
            return
        ztd, config, connect = self._ztd_args
        self._set_running(True)
        self._num_runs = 0
        self._num_to_save = 0
        self._elapsed = 0
        self.lbl_elapsed.setText("已用时间 00:00")
        self.lbl_runinfo.setText("ZTD 分析中")
        self._timer.start()
        self._thread = QtCore.QThread(self)
        self._worker = _ZtdWorker(ztd, config, connect, zos_dir=zos_dir)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.log.connect(self._append_log)
        self._worker.finished.connect(self._on_ztd_finished)
        self._worker.need_zos_dir.connect(self._on_need_zos_dir)
        self._thread.start()

    def _start_field_preview_worker(self, zos_dir):
        if not self._field_preview_args:
            self._warn("尚未设置视场映射预览参数，请先点击「预览视场映射」。")
            return
        zmx, run_params, connect = self._field_preview_args
        self._set_running(True)
        self._num_runs = 0
        self._num_to_save = 0
        self._elapsed = 0
        self.lbl_elapsed.setText("已用时间 00:00")
        self.lbl_runinfo.setText("视场映射预览中")
        self._timer.start()
        self._thread = QtCore.QThread(self)
        self._worker = _FieldPreviewWorker(zmx, run_params, connect, zos_dir=zos_dir)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.log.connect(self._append_log)
        self._worker.finished.connect(self._on_field_preview_finished)
        self._worker.need_zos_dir.connect(self._on_need_zos_dir)
        self._thread.start()

    def _on_tick(self):
        self._elapsed += 1
        m, s = divmod(self._elapsed, 60)
        self.lbl_elapsed.setText(f"已用时间 {m:02d}:{s:02d}")

    @QtCore.Slot(int, int)
    def _on_started(self, num_runs: int, num_to_save: int):
        self._num_runs = num_runs
        self._num_to_save = num_to_save
        self.lbl_runinfo.setText(
            f"运行总数 {num_runs} 次 / 保存总数 {num_to_save} 个")

    def _on_cancel(self):
        if self._worker is not None:
            self._worker.cancel()
            self.btn_cancel.setEnabled(False)
            self._append_log("正在请求取消…（等待当前蒙特卡洛步完成）")

    def _on_copy_log(self):
        text = self.log_view.toPlainText()
        QtWidgets.QApplication.clipboard().setText(text)
        self.statusBar().showMessage("日志已复制到剪贴板", 2000)

    def _on_clear_log(self):
        self.log_view.clear()
        self.statusBar().showMessage("日志已清空", 2000)

    def _set_last_result_dir(self, path: str) -> None:
        if os.path.isfile(path):
            path = os.path.dirname(path)
        if os.path.isdir(path):
            self._settings.setValue("last_result_dir", os.path.abspath(path))

    def _on_open_result_dir(self):
        path = self._setting("last_result_dir", "")
        if not path or not os.path.isdir(path):
            fallback = self.ed_outdir.text().strip()
            if os.path.isdir(fallback):
                path = fallback
        if not path or not os.path.isdir(path):
            self._warn("暂无可打开的结果目录。请先完成一次分析，或确认输出目录存在。")
            return
        QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(os.path.abspath(path)))
        self.statusBar().showMessage("已打开结果目录", 2000)

    @QtCore.Slot(bool, str)
    def _on_finished(self, ok: bool, info: str):
        self._timer.stop()
        if self._thread is not None:
            self._thread.quit()
            self._thread.wait(5000)
            self._thread = None
        self._worker = None
        self._set_running(False)
        self._active_task = ""
        if ok:
            self._set_last_result_dir(info)
            self._append_log("✅ 分析完成。")
            self.statusBar().showMessage("分析完成")
            if self._num_runs:
                self.lbl_runinfo.setText(
                    f"已完成：运行 {self._num_runs} 次 / 保存 {self._num_to_save} 个")
        else:
            self._append_log("❌ 失败：" + info)
            self.statusBar().showMessage("失败")
            self.lbl_runinfo.setText("已停止")
            self._warn("公差分析失败：\n" + info)

    @QtCore.Slot(bool, str)
    def _on_ztd_finished(self, ok: bool, info: str):
        self._timer.stop()
        if self._thread is not None:
            self._thread.quit()
            self._thread.wait(5000)
            self._thread = None
        self._worker = None
        self._set_running(False)
        self._active_task = ""
        if ok:
            self._set_last_result_dir(info)
            self._append_log("✅ ZTD 分析完成。")
            self.statusBar().showMessage("ZTD 分析完成")
            self.lbl_runinfo.setText("ZTD 分析完成")
        else:
            self._append_log("❌ ZTD 分析失败：" + info)
            self.statusBar().showMessage("ZTD 分析失败")
            self.lbl_runinfo.setText("ZTD 分析失败")
            self._warn("ZTD 分析失败：\n" + info)

    @QtCore.Slot(bool, str)
    def _on_field_preview_finished(self, ok: bool, info: str):
        self._timer.stop()
        if self._thread is not None:
            self._thread.quit()
            self._thread.wait(5000)
            self._thread = None
        self._worker = None
        self._set_running(False)
        self._active_task = ""
        if ok:
            self._append_log("✅ " + info)
            self.statusBar().showMessage("视场映射预览完成")
            self.lbl_runinfo.setText("视场映射预览完成")
        else:
            self._append_log("❌ 视场映射预览失败：" + info)
            self.statusBar().showMessage("视场映射预览失败")
            self.lbl_runinfo.setText("视场映射预览失败")
            self._warn("视场映射预览失败：\n" + info)

    @QtCore.Slot(list)
    def _on_need_zos_dir(self, searched: list):
        self._timer.stop()
        if self._thread is not None:
            self._thread.quit()
            self._thread.wait(5000)
            self._thread = None
        self._worker = None
        self._set_running(False)
        self._append_log("⚠ 未自动找到 Zemax 安装目录，请手动指定。")

        tip = "未能自动找到 Zemax（ZOS-API）安装目录。"
        if searched:
            tip += "\n\n已搜索以下位置：\n" + "\n".join(
                "  - " + str(p) for p in searched[:20])
        tip += "\n\n请选择 Zemax OpticStudio 安装目录（含 ZOSAPI_NetHelper.dll）。"

        max_attempts = 5
        for _ in range(max_attempts):
            ret = _dark_message_box(
                self, QtWidgets.QMessageBox.Warning, "未找到 Zemax", tip,
                QtWidgets.QMessageBox.Ok | QtWidgets.QMessageBox.Cancel,
                QtWidgets.QMessageBox.Ok)
            if ret != QtWidgets.QMessageBox.Ok:
                self._active_task = ""
                self._append_log("已取消指定 Zemax 目录。")
                self.statusBar().showMessage("已取消")
                return
            d = QtWidgets.QFileDialog.getExistingDirectory(
                self, "选择 Zemax OpticStudio 安装目录")
            if not d:
                continue
            if not zos_connect.is_valid_zos_dir(d):
                self._warn(
                    "该目录下未找到完整的 ZOS-API DLL，请重新选择。\n"
                    "（需含 ZOSAPI_NetHelper.dll / ZOSAPI_Interfaces.dll / ZOSAPI.dll）\n\n"
                    f"你选的是：\n{d}")
                continue
            try:
                path = zos_connect.save_zos_dir_to_config(d)
                self._append_log(f"已记住 Zemax 目录并写入：{path}")
            except Exception as e:
                self._append_log(f"写入配置失败（不影响本次运行）：{e}")
            self._append_log(f"使用 Zemax 目录重新连接：{d}")
            if self._active_task == "ztd":
                self._start_ztd_worker(d)
            elif self._active_task == "field_preview":
                self._start_field_preview_worker(d)
            else:
                self._start_worker(d)
            return

        self._append_log(f"已连续 {max_attempts} 次未选定有效目录，已停止。")
        self.statusBar().showMessage("已取消")

    def _set_running(self, running: bool):
        self.btn_run.setEnabled(not running)
        self.btn_ztd.setEnabled(not running)
        self.btn_cancel.setEnabled(running and self._active_task == "tol")
        # 运行态统一禁用这些按钮；非运行态的显隐与 enable 交给 _on_analysis_mode_changed 复原。
        for btn in (self.btn_export_std, self.btn_check_config,
                    self.btn_preview_fields, self.btn_open_result):
            btn.setEnabled(not running)
        for w in (self.ed_zmx, self.ed_config, self.ed_outdir,
                  self.ed_ztd, self.cb_mode, self.cb_analysis_mode,
                  self.standard_panel):
            w.setEnabled(not running)
        if not running:
            self._on_analysis_mode_changed()
        if running:
            if self._active_task == "ztd":
                msg = "ZTD 分析中…"
            elif self._active_task == "field_preview":
                msg = "视场映射预览中…"
            else:
                msg = "公差分析运行中…"
        else:
            msg = "就绪"
        self.statusBar().showMessage(msg)

    def _warn(self, msg: str):
        _dark_message_box(
            self, QtWidgets.QMessageBox.Warning, "提示", msg,
            QtWidgets.QMessageBox.Ok, QtWidgets.QMessageBox.Ok)

    def closeEvent(self, event: QtGui.QCloseEvent):
        if self._thread is not None and self._thread.isRunning():
            if self._active_task == "ztd":
                text = (
                    "ZTD 分析正在运行中。\n\n"
                    "关闭窗口将中断已有 ZTD 的读取与统计导出，"
                    "当前统计 Excel 可能不会生成。\n\n确定要关闭吗？")
            elif self._active_task == "field_preview":
                text = (
                    "视场映射预览正在运行中。\n\n"
                    "关闭窗口将中断预览并关闭后台 Zemax 实例。\n\n确定要关闭吗？")
            else:
                text = (
                    "公差分析正在运行中。\n\n"
                    "关闭窗口将强制中断蒙特卡洛分析并关闭后台 Zemax 实例，"
                    "未完成的结果会丢失。\n\n确定要关闭吗？")
            ret = _dark_message_box(
                self, QtWidgets.QMessageBox.Question, "确认关闭", text,
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                QtWidgets.QMessageBox.No)
            if ret != QtWidgets.QMessageBox.Yes:
                event.ignore()
                return
            if self._worker is not None and hasattr(self._worker, "force_stop"):
                self._worker.force_stop()
            self._thread.quit()
            self._thread.wait(5000)
        super().closeEvent(event)


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    _apply_dark_palette(app)
    win = MainWindow()
    win.show()
    _apply_dark_titlebar(win)
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
