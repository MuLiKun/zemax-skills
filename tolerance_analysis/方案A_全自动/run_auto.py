r"""方案A —— 全自动蒙特卡洛公差分析。

程序自动：连接 → 建 TDE/MFE/TSC → Save → API 直接跑完蒙特卡洛，
中途无需用户在 OpticStudio 点 Run；进度实时打印在本终端。

输出（ZTD / Worst / Best）落在本脚本所在的「方案A_全自动\output」目录，
便于与方案B 区分。

用法（在你自己的终端前台运行，先在 OpticStudio 进入交互扩展等待状态）：
    .\.venv\Scripts\python.exe -u "方案A_全自动\run_auto.py" ^
        --zmx "F:\个人文件\P3111\05304_tol.zmx" --config "tol_config_05304.xlsx"
"""

from __future__ import annotations

import argparse
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_HERE)
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from toltool import pipeline


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description="方案A：全自动蒙特卡洛公差分析")
    p.add_argument("--zmx", required=True, help="待分析的 zmx 文件路径")
    p.add_argument("--config", required=True, help="Excel 配置文件路径")
    p.add_argument("--connect", choices=["extension", "standalone"],
                   default="extension", help="连接方式，默认 extension")
    args = p.parse_args(argv)

    outdir = os.path.join(_HERE, "output")

    prep = pipeline.prepare_session(
        args.zmx, args.config, outdir=outdir, connect=args.connect)

    result = pipeline.run_montecarlo(prep)
    if not result.succeeded:
        print(f"公差分析失败: {result.message}", file=sys.stderr)
        return 1

    print(f"分析完成。ZTD: {result.ztd_path}")
    if result.bestworst_folder:
        print(f"Worst/Best 输出目录: {result.bestworst_folder}")
    print("（结果可在 OpticStudio 或上述 ZTD 中查看；本方案不读取 ZTD）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
