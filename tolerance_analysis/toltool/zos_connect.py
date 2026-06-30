"""zos_connect.py —— 连接 Zemax 并读取镜头信息。

职责（需求文档 §7）：
- 定位本机 ZOS-API DLL（注册表 / 候选目录），用 pythonnet 加载。
- 以交互扩展(ConnectAsExtension)或独立实例(CreateNewApplication)连接。
- 打开指定 zmx 并另存为工作副本（不改原文件）。
- 读取镜头基本信息（面数、半径、厚度、玻璃、视场、波长、光阑面）。

注意：所有具体路径都在运行时核实，不写死常量。
"""

from __future__ import annotations

import configparser
import glob
import os
import sys
import winreg
from dataclasses import dataclass, field
from typing import Optional


_DLL_NAMES = ("ZOSAPI_NetHelper.dll", "ZOSAPI_Interfaces.dll", "ZOSAPI.dll")

_ENV_VAR = "ZEMAX_ZOS_DIR"

_CONFIG_NAME = "zemax_config.ini"


class ZosDirNotFound(Exception):
    """无法定位含 ZOS-API DLL 的 Zemax 安装目录。

    携带 ``searched`` 列表（已尝试过的路径），供 GUI 弹窗展示给用户。
    """

    def __init__(self, searched: Optional[list[str]] = None):
        self.searched = searched or []
        msg = ("未找到含 ZOS-API DLL 的安装目录，请确认 OpticStudio 已安装。"
               "（需同时含 ZOSAPI_NetHelper.dll / ZOSAPI_Interfaces.dll / ZOSAPI.dll）")
        if self.searched:
            msg += "\n已搜索以下位置：\n" + "\n".join("  - " + p for p in self.searched)
        super().__init__(msg)


def is_valid_zos_dir(d: Optional[str]) -> bool:
    """判断目录是否同时含三个 ZOS-API DLL。"""
    return bool(d) and all(os.path.isfile(os.path.join(d, n)) for n in _DLL_NAMES)


def _config_path() -> str:
    """配置文件路径：打包后为 exe 同级，否则为项目根（zos_connect.py 上两层）。"""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(os.path.abspath(sys.executable))
    else:
        base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, _CONFIG_NAME)


def _read_config_zos_dir() -> Optional[str]:
    path = _config_path()
    if not os.path.isfile(path):
        return None
    try:
        cp = configparser.ConfigParser()
        cp.read(path, encoding="utf-8")
        val = cp.get("zemax", "zos_dir", fallback="").strip()
        return val or None
    except Exception:
        return None


def save_zos_dir_to_config(d: str) -> str:
    """把 zos_dir 写入同目录 zemax_config.ini。返回写入的文件路径。"""
    path = _config_path()
    cp = configparser.ConfigParser()
    if os.path.isfile(path):
        try:
            cp.read(path, encoding="utf-8")
        except Exception:
            pass
    if not cp.has_section("zemax"):
        cp.add_section("zemax")
    cp.set("zemax", "zos_dir", d)
    with open(path, "w", encoding="utf-8") as f:
        f.write("# Zemax OpticStudio 安装目录（含 ZOSAPI_NetHelper.dll）\n")
        f.write("# 首次自动查找失败并由弹窗选定后自动写入；也可手动填写\n")
        cp.write(f)
    return path


def _registry_candidates() -> list[str]:
    """从注册表读取候选目录。ZemaxRoot 常指向用户数据目录，
    这里同时把其本身及向上回溯/同级目录一并纳入候选。"""
    cands: list[str] = []
    for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        for sub in (r"Software\Zemax", r"Software\Ansys\Zemax OpticStudio"):
            try:
                with winreg.OpenKey(hive, sub) as key:
                    for vname in ("ZemaxRoot", "InstallDir", "InstallPath"):
                        try:
                            val, _ = winreg.QueryValueEx(key, vname)
                        except OSError:
                            continue
                        if not val:
                            continue
                        cands.append(val)
                        parent = os.path.dirname(val.rstrip("\\/"))
                        if parent:
                            cands.append(parent)
            except OSError:
                continue
    return cands


def _glob_candidates() -> list[str]:
    """扫描 C:–Z: 实际存在的盘符 × 多路径模板，glob 通配版本号。"""
    drives = [f"{chr(c)}:\\" for c in range(ord("C"), ord("Z") + 1)
              if os.path.isdir(f"{chr(c)}:\\")]
    templates = [
        r"Program Files\Ansys Zemax OpticStudio *",
        r"Program Files\Zemax OpticStudio*",
        r"Program Files\ANSYS Inc\*\Zemax*",
        r"Program Files (x86)\Ansys Zemax OpticStudio *",
        r"Program Files (x86)\Zemax OpticStudio*",
    ]
    cands: list[str] = []
    for drv in drives:
        for tpl in templates:
            cands.extend(glob.glob(os.path.join(drv, tpl)))
    return cands


def find_zos_dir(collect_searched: Optional[list[str]] = None) -> Optional[str]:
    """方案四：智能搜索含 ZOS-API DLL 的安装目录。

    顺序：注册表候选（含向上回溯）→ 多盘 glob 智能搜索。
    只接受确实同时含三个 ZOS-API DLL 的目录。找到返回路径，否则返回 None。
    若传入 collect_searched 列表，会把所有尝试过的目录写入其中（供弹窗展示）。
    """
    seen: list[str] = []

    def _try(cands: list[str]) -> Optional[str]:
        for d in cands:
            if not d:
                continue
            nd = os.path.normpath(d)
            if nd not in seen:
                seen.append(nd)
            if is_valid_zos_dir(d):
                return d
        return None

    hit = _try(_registry_candidates())
    if hit is None:
        hit = _try(_glob_candidates())
    if collect_searched is not None:
        collect_searched.extend(seen)
    return hit


def resolve_zos_dir(explicit: Optional[str] = None) -> str:
    """按优先级返回含 DLL 的安装目录，全部失败抛 ZosDirNotFound。

    优先级：
      1. 显式传入 explicit
      2. INI 配置 [zemax] zos_dir
      3. 环境变量 ZEMAX_ZOS_DIR
      4. 注册表 + 多盘智能搜索（find_zos_dir）
    """
    searched: list[str] = []

    if explicit:
        if is_valid_zos_dir(explicit):
            return explicit
        searched.append(f"[显式传入] {explicit}")

    cfg = _read_config_zos_dir()
    if cfg:
        if is_valid_zos_dir(cfg):
            return cfg
        searched.append(f"[配置文件] {cfg}")

    env = os.environ.get(_ENV_VAR, "").strip()
    if env:
        if is_valid_zos_dir(env):
            return env
        searched.append(f"[环境变量 {_ENV_VAR}] {env}")

    auto_searched: list[str] = []
    hit = find_zos_dir(collect_searched=auto_searched)
    if hit:
        return hit
    searched.extend(auto_searched)

    raise ZosDirNotFound(searched)


def _load_zosapi(zos_dir: str):
    """用 pythonnet 加载三个 ZOS-API 程序集。返回 ZOSAPI 命名空间模块。"""
    import clr  # noqa: F401  pythonnet

    helper = os.path.join(zos_dir, "ZOSAPI_NetHelper.dll")
    if not os.path.isfile(helper):
        raise FileNotFoundError(f"未找到 {helper}")
    clr.AddReference(helper)
    import ZOSAPI_NetHelper  # type: ignore

    ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize(zos_dir)
    clr.AddReference(os.path.join(zos_dir, "ZOSAPI_Interfaces.dll"))
    clr.AddReference(os.path.join(zos_dir, "ZOSAPI.dll"))
    import ZOSAPI  # type: ignore

    return ZOSAPI


@dataclass
class SurfaceInfo:
    index: int
    comment: str
    radius: float
    thickness: float
    material: str
    is_stop: bool


@dataclass
class LensInfo:
    system_file: str
    num_surfaces: int
    stop_surface: int
    surfaces: list[SurfaceInfo] = field(default_factory=list)
    wavelengths_um: list[float] = field(default_factory=list)
    primary_wave: int = 0
    fields: list[tuple[float, float]] = field(default_factory=list)
    field_type: str = ""
    aperture_value: float = 0.0
    aperture_type: str = ""


class ZosSession:
    """ZOS-API 连接会话。封装连接、打开副本、读取信息。"""

    def __init__(self, zos_dir: Optional[str] = None):
        self.zos_dir = resolve_zos_dir(zos_dir)
        self.ZOSAPI = _load_zosapi(self.zos_dir)
        self._conn = None
        self.app = None
        self.sys = None

    def connect(self, mode: str = "extension", instance: int = 1) -> None:
        ZOSAPI = self.ZOSAPI
        self._conn = ZOSAPI.ZOSAPI_Connection()
        if mode == "extension":
            self.app = self._conn.ConnectAsExtension(instance)
        elif mode == "standalone":
            self.app = self._conn.CreateNewApplication()
        else:
            raise ValueError(f"未知连接模式: {mode}")
        if self.app is None:
            raise RuntimeError(
                "连接 Zemax 失败：扩展模式需先在 OpticStudio 中开启交互扩展。"
            )
        self.sys = self.app.PrimarySystem
        if self.sys is None:
            raise RuntimeError("PrimarySystem 为空，连接异常。")

    def open_as_copy(self, zmx_path: str, copy_path: Optional[str] = None) -> str:
        """打开 zmx 并另存为工作副本，后续操作都在副本上进行。返回副本路径。"""
        if not os.path.isfile(zmx_path):
            raise FileNotFoundError(zmx_path)
        if copy_path is None:
            base, ext = os.path.splitext(zmx_path)
            copy_path = f"{base}_tol{ext}"
        self.sys.LoadFile(zmx_path, False)
        self.sys.SaveAs(copy_path)
        return copy_path

    def read_lens_info(self) -> LensInfo:
        ZOSAPI = self.ZOSAPI
        s = self.sys
        lde = s.LDE
        n = lde.NumberOfSurfaces
        stop = -1
        surfaces: list[SurfaceInfo] = []
        for i in range(n):
            surf = lde.GetSurfaceAt(i)
            is_stop = bool(surf.IsStop)
            if is_stop:
                stop = i
            surfaces.append(
                SurfaceInfo(
                    index=i,
                    comment=str(surf.Comment),
                    radius=float(surf.Radius),
                    thickness=float(surf.Thickness),
                    material=str(surf.Material),
                    is_stop=is_stop,
                )
            )

        sd = s.SystemData
        waves = sd.Wavelengths
        wl = [float(waves.GetWavelength(i).Wavelength)
              for i in range(1, waves.NumberOfWavelengths + 1)]
        
        primary = 0
        for i in range(1, waves.NumberOfWavelengths + 1):
            if waves.GetWavelength(i).IsPrimary:
                primary = i
                break

        flds = sd.Fields
        field_list = [
            (float(flds.GetField(i).X), float(flds.GetField(i).Y))
            for i in range(1, flds.NumberOfFields + 1)
        ]

        ap = sd.Aperture
        return LensInfo(
            system_file=str(s.SystemFile),
            num_surfaces=n,
            stop_surface=stop,
            surfaces=surfaces,
            wavelengths_um=wl,
            primary_wave=primary,
            fields=field_list,
            field_type=str(flds.GetFieldType()),
            aperture_value=float(ap.ApertureValue),
            aperture_type=str(ap.ApertureType),
        )

    def close(self) -> None:
        if self.app is not None:
            try:
                self.app.CloseApplication()
            except Exception:
                pass
            self.app = None
            self.sys = None
