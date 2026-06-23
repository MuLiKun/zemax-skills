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
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

from PySide6 import QtCore, QtGui, QtWidgets

from toltool import pipeline
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
                 zos_dir: str | None = None):
        super().__init__()
        self._zmx = zmx
        self._config = config
        self._outdir = outdir
        self._connect = connect
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
        try:
            prep = pipeline.prepare_session(
                self._zmx, self._config, outdir=self._outdir,
                connect=self._connect, log=self.log.emit,
                zos_dir=self._zos_dir)
            sess = prep.sess
            self._sess = sess

            spec = pipeline.make_runspec(prep)
            self.started.emit(int(spec.num_runs), int(spec.num_to_save))
            self.log.emit(
                f"开始公差分析：{spec.num_runs} 次蒙特卡洛"
                f"（{spec.distribution}分布）…")

            from toltool import tol_runner
            result = tol_runner.run(
                prep.sess.sys, spec,
                progress_cb=lambda p, m: self.log.emit(f"  [{p:>3}%] {m}"),
                cancel_flag=lambda: self._cancel)

            if self._force:
                self.finished.emit(False, "已强制停止（后台 Zemax 已关闭）")
                return
            if not result.succeeded:
                self.finished.emit(False, result.message or "公差分析失败")
                return

            self.log.emit(f"分析完成。ZTD: {result.ztd_path}")
            if result.bestworst_folder:
                self.log.emit(f"Worst/Best 输出目录: {result.bestworst_folder}")

            if _yes(prep.rp.get("输出统计Excel", "N")):
                from toltool import ztd_reader
                report_meta = [
                    r for r in prep.cfg.report
                    if _yes(r.get("启用")) and r.get("标签")
                ]
                report_labels = [str(r.get("标签")).strip() for r in report_meta]
                num_runs = _as_int(prep.rp.get("蒙特卡洛次数"), int(spec.num_runs))
                self.log.emit("正在读取 ZTD 并导出统计 Excel…")
                zres = ztd_reader.read_ztd(
                    prep.sess.sys, result.ztd_path, num_runs=num_runs,
                    report_labels=report_labels or None,
                    report_meta=report_meta or None)
                if not zres.succeeded:
                    self.finished.emit(False, "分析完成，但读取 ZTD 失败：" + zres.message)
                    return
                if zres.message:
                    self.log.emit("提示：" + zres.message)
                stat_path = result.ztd_path.rsplit(".", 1)[0] + "_统计.xlsx"
                out = ztd_reader.export_excel(zres, stat_path)
                self.log.emit(f"统计 Excel: {out}")

            self.finished.emit(True, result.ztd_path)
        except zos_connect.ZosDirNotFound as e:
            self.need_zos_dir.emit(list(e.searched))
        except Exception as e:
            if self._force:
                self.finished.emit(False, "已强制停止（后台 Zemax 已关闭）")
            else:
                self.finished.emit(False, f"{type(e).__name__}: {e}")
        finally:
            if sess is not None and not self._force \
                    and self._connect == "standalone":
                try:
                    sess.close()
                    self.log.emit("已释放 Zemax 独立实例。")
                except Exception:
                    pass
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

            self.log.emit(f"读取配置 Excel: {self._config}")
            cfg = excel_io.read_config(self._config)
            report_meta = [
                r for r in cfg.report if _yes(r.get("启用")) and r.get("标签")
            ]
            report_labels = [str(r.get("标签")).strip() for r in report_meta]
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

            self.log.emit(f"正在读取 ZTD: {self._ztd}")
            zres = ztd_reader.read_ztd(
                sess.sys, self._ztd, num_runs=num_runs,
                report_labels=report_labels or None,
                report_meta=report_meta or None)
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
                except Exception:
                    pass
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
        self._add_file_row(form, 1, "Excel 配置：", self.ed_config,
                           self._pick_config)
        self._add_file_row(form, 2, "输出目录：", self.ed_outdir,
                           self._pick_outdir)

        form.addWidget(QtWidgets.QLabel("连接模式："), 3, 0)
        self.cb_mode = QtWidgets.QComboBox()
        self.cb_mode.addItem("Standalone（程序后台挂起，推荐）", "standalone")
        self.cb_mode.addItem("GUI 模式（连入交互扩展窗口）", "extension")
        mode_index = self.cb_mode.findData(self._setting("connect", "standalone"))
        if mode_index >= 0:
            self.cb_mode.setCurrentIndex(mode_index)
        self.cb_mode.currentIndexChanged.connect(
            lambda: self._settings.setValue(
                "connect", self.cb_mode.currentData()))
        form.addWidget(self.cb_mode, 3, 1, 1, 2)

        self.ed_ztd = QtWidgets.QLineEdit(self._setting("ztd", ""))
        self._add_file_row(form, 4, "已有 ZTD：", self.ed_ztd,
                           self._pick_ztd)

        ztdbar = QtWidgets.QHBoxLayout()
        ztdbar.addStretch(1)
        self.btn_ztd = QtWidgets.QPushButton("分析已有 ZTD")
        self.btn_ztd.clicked.connect(self._on_analyze_ztd)
        ztdbar.addWidget(self.btn_ztd)
        root.addLayout(ztdbar)

        btns = QtWidgets.QHBoxLayout()
        btns.addStretch(1)
        self.btn_run = QtWidgets.QPushButton("开始分析")
        self.btn_run.setObjectName("run")
        self.btn_run.clicked.connect(self._on_run)
        self.btn_cancel = QtWidgets.QPushButton("取消")
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.clicked.connect(self._on_cancel)
        btns.addWidget(self.btn_cancel)
        btns.addWidget(self.btn_run)
        root.addLayout(btns)

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
        grid.addWidget(QtWidgets.QLabel(label), row, 0)
        grid.addWidget(edit, row, 1)
        btn = QtWidgets.QPushButton("浏览…")
        btn.clicked.connect(slot)
        grid.addWidget(btn, row, 2)

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

    def _on_run(self):
        zmx = self.ed_zmx.text().strip()
        config = self.ed_config.text().strip()
        outdir = self.ed_outdir.text().strip()
        connect = self.cb_mode.currentData()

        if not os.path.isfile(zmx):
            self._warn("zmx 文件不存在：\n" + zmx)
            return
        if not os.path.isfile(config):
            self._warn("Excel 配置不存在：\n" + config)
            return
        if not outdir:
            outdir = os.path.join(_HERE, "output")
        os.makedirs(outdir, exist_ok=True)
        self.ed_outdir.setText(outdir)
        self._remember_form_paths()

        self.log_view.clear()
        self._append_log(
            f"连接模式: {'Standalone' if connect == 'standalone' else 'GUI(交互扩展)'}")
        self._append_log(f"待分析镜头: {zmx}")
        self._append_log(f"配置 Excel: {config}")
        self._append_log(f"输出目录: {outdir}")

        self._run_args = (zmx, config, outdir, connect)
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
        if not os.path.isfile(config):
            self._warn("Excel 配置不存在：\n" + config)
            return
        self._remember_form_paths()

        self.log_view.clear()
        self._append_log("开始独立分析已有 ZTD。")
        self._append_log(f"ZTD 文件: {ztd}")
        self._append_log(f"配置 Excel: {config}")
        self._append_log(f"统计输出目录: {os.path.dirname(os.path.abspath(ztd))}")

        self._ztd_args = (ztd, config, connect)
        self._active_task = "ztd"
        self._start_ztd_worker(None)

    def _start_worker(self, zos_dir):
        if not self._run_args:
            self._warn("尚未设置运行参数，请先点击「开始分析」。")
            return
        zmx, config, outdir, connect = self._run_args
        self._set_running(True)
        self._num_runs = 0
        self._num_to_save = 0
        self._elapsed = 0
        self.lbl_elapsed.setText("已用时间 00:00")
        self.lbl_runinfo.setText("")
        self._timer.start()
        self._thread = QtCore.QThread(self)
        self._worker = _Worker(zmx, config, outdir, connect, zos_dir=zos_dir)
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

    @QtCore.Slot(bool, str)
    def _on_finished(self, ok: bool, info: str):
        self._timer.stop()
        if self._thread is not None:
            self._thread.quit()
            self._thread.wait()
            self._thread = None
        self._worker = None
        self._set_running(False)
        self._active_task = ""
        if ok:
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
            self._thread.wait()
            self._thread = None
        self._worker = None
        self._set_running(False)
        self._active_task = ""
        if ok:
            self._append_log("✅ ZTD 分析完成。")
            self.statusBar().showMessage("ZTD 分析完成")
            self.lbl_runinfo.setText("ZTD 分析完成")
        else:
            self._append_log("❌ ZTD 分析失败：" + info)
            self.statusBar().showMessage("ZTD 分析失败")
            self.lbl_runinfo.setText("ZTD 分析失败")
            self._warn("ZTD 分析失败：\n" + info)

    @QtCore.Slot(list)
    def _on_need_zos_dir(self, searched: list):
        self._timer.stop()
        if self._thread is not None:
            self._thread.quit()
            self._thread.wait()
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
            else:
                self._start_worker(d)
            return

        self._append_log(f"已连续 {max_attempts} 次未选定有效目录，已停止。")
        self.statusBar().showMessage("已取消")

    def _set_running(self, running: bool):
        self.btn_run.setEnabled(not running)
        self.btn_ztd.setEnabled(not running)
        self.btn_cancel.setEnabled(running and self._active_task == "tol")
        for w in (self.ed_zmx, self.ed_config, self.ed_outdir,
                  self.ed_ztd, self.cb_mode):
            w.setEnabled(not running)
        if running:
            msg = "ZTD 分析中…" if self._active_task == "ztd" else "公差分析运行中…"
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
