r"""main.py —— 公差分析主入口（在代码里设置地址，直接运行）。

使用方法：
  1. 在 OpticStudio 进入「编程 → 交互扩展」等待状态；
  2. 修改下面的 ZMX_FILE / CONFIG_FILE 两个地址；
  3. 直接运行本文件：
         .\.venv\Scripts\python.exe -u main.py

也可用命令行参数临时覆盖写死的地址（不传则用下面的常量）：
    .\.venv\Scripts\python.exe -u main.py --zmx "..." --config "..."

流程：连接 → 建 TDE/MFE/TSC → Save → API 直接跑完蒙特卡洛。
输出（ZTD / Worst / Best）落在 output 目录，跑完不读取 ZTD（在 OpticStudio 查看）。
"""

from __future__ import annotations

import argparse
import os
import sys

# ===== 在这里设置地址 =====================================================
ZMX_FILE = r"F:\个人文件\P3111\05304_tol.zmx"      # 待分析的 zmx 文件
CONFIG_FILE = r"F:\个人文件\P3111\tol_config_05304.xlsx"             # 输入的 Excel 配置
CONNECT_MODE = "extension"                          # extension 或 standalone
# =========================================================================

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

from toltool import pipeline


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description="公差分析主入口（地址写在文件顶部）")
    p.add_argument("--zmx", default=ZMX_FILE, help="待分析的 zmx 文件路径")
    p.add_argument("--config", default=CONFIG_FILE, help="Excel 配置文件路径")
    p.add_argument("--connect", choices=["extension", "standalone"],
                   default=CONNECT_MODE, help="连接方式，默认 extension")
    args = p.parse_args(argv)

    outdir = os.path.join(_HERE, "output")

    print(f"待分析镜头: {args.zmx}")
    print(f"配置 Excel: {args.config}")

    prep = pipeline.prepare_session(
        args.zmx, args.config, outdir=outdir, connect=args.connect)

    result = pipeline.run_montecarlo(prep)
    if not result.succeeded:
        print(f"公差分析失败: {result.message}", file=sys.stderr)
        return 1

    print(f"分析完成。ZTD: {result.ztd_path}")
    if result.bestworst_folder:
        print(f"Worst/Best 输出目录: {result.bestworst_folder}")
    print("（结果可在 OpticStudio 或上述 ZTD 中查看；本入口不读取 ZTD）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
