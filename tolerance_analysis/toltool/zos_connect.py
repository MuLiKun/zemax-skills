"""zos_connect.py —— 连接 Zemax 并读取镜头信息。

职责（需求文档 §7）：
- 定位本机 ZOS-API DLL（注册表 / 候选目录），用 pythonnet 加载。
- 以交互扩展(ConnectAsExtension)或独立实例(CreateNewApplication)连接。
- 打开指定 zmx 并另存为工作副本（不改原文件）。
- 读取镜头基本信息（面数、半径、厚度、玻璃、视场、波长、光阑面）。

注意：所有具体路径都在运行时核实，不写死常量。
"""

from __future__ import annotations

import os
import sys
import winreg
from dataclasses import dataclass, field
from typing import Optional


_CANDIDATE_DIRS = [
    r"C:\Program Files\Ansys Zemax OpticStudio 2023 R1.00",
    r"C:\Program Files\Zemax OpticStudio",
]


def _has_dll(d: str) -> bool:
    return bool(d) and os.path.isfile(os.path.join(d, "ZOSAPI_NetHelper.dll"))


def find_zos_dir() -> str:
    """定位含 ZOS-API DLL 的安装目录。

    注意：注册表 Software\\Zemax 的 ZemaxRoot 通常指向用户数据目录
    （如 ~\\Documents\\Zemax），并不含 DLL。因此只接受确实存在
    ZOSAPI_NetHelper.dll 的目录。
    """
    candidates: list[str] = []
    for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        for sub in (r"Software\Zemax", r"Software\Ansys\Zemax OpticStudio"):
            try:
                with winreg.OpenKey(hive, sub) as key:
                    for vname in ("ZemaxRoot", "InstallDir", "InstallPath"):
                        try:
                            val, _ = winreg.QueryValueEx(key, vname)
                            if val:
                                candidates.append(val)
                        except OSError:
                            pass
            except OSError:
                continue
    candidates.extend(_CANDIDATE_DIRS)
    for d in candidates:
        if _has_dll(d):
            return d
    pf = os.environ.get("ProgramFiles", r"C:\Program Files")
    ansys = os.path.join(pf, "Ansys")
    if os.path.isdir(ansys):
        for name in sorted(os.listdir(ansys), reverse=True):
            d = os.path.join(ansys, name)
            if _has_dll(d):
                return d
    raise FileNotFoundError(
        "未找到含 ZOSAPI_NetHelper.dll 的 ZOS-API 安装目录，请确认 OpticStudio 已安装。"
    )


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
        self.zos_dir = zos_dir or find_zos_dir()
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
        primary = int(waves.GetWavelength(1).Wavelength and waves.NumberOfWavelengths)
        try:
            primary = int(sd.Wavelengths.GetWavelength(1).IsPrimary and 1) or 0
        except Exception:
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
