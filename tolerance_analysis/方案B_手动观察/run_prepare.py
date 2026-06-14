r"""方案B —— 仅准备，手动在 OpticStudio 观察蒙特卡洛运行。

程序自动：连接 → 建 TDE/MFE/TSC → Save，然后停下，不调用 API 跑蒙卡。
你在 OpticStudio 里：公差分析(Tolerancing) 对话框 → Criterion 选 User Script
→ 下拉选下面打印的 TSC → 设 Monte Carlo 次数 → 点 Run，
即可看到原生公差分析窗口逐次试验实时滚动。

输出（你在 GUI 里设的 ZTD / Worst / Best）建议手动保存到本脚本所在的
「方案B_手动观察\output」目录，便于与方案A 区分。

用法（在你自己的终端前台运行，先在 OpticStudio 进入交互扩展等待状态）：
    .\.venv\Scripts\python.exe -u "方案B_手动观察\run_prepare.py" ^
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
    p = argparse.ArgumentParser(
        description="方案B：仅准备 TDE/MFE/TSC，由用户在 OpticStudio 手动跑")
    p.add_argument("--zmx", required=True, help="待分析的 zmx 文件路径")
    p.add_argument("--config", required=True, help="Excel 配置文件路径")
    p.add_argument("--connect", choices=["extension", "standalone"],
                   default="extension", help="连接方式，默认 extension")
    args = p.parse_args(argv)

    outdir = os.path.join(_HERE, "output")

    prep = pipeline.prepare_session(
        args.zmx, args.config, outdir=outdir, connect=args.connect)
    rp = prep.rp

    print("\n================ 准备完成，请在 OpticStudio 手动运行 ================")
    print(f"  TDE 公差条数 : {prep.n_tde}")
    print(f"  MFE 行数     : {prep.n_mfe}")
    print(f"  TSC 文件     : {prep.tsc_path}")
    print(f"  REPORT 分项  : {prep.n_report}")
    print("\n操作步骤：")
    print("  1. 公差分析 → 公差分析(Tolerancing) 打开对话框")
    print("  2. Setup → Criterion 选 User Script")
    print(f"     脚本选: {os.path.basename(prep.tsc_path)}")
    print(f"  3. Number of Runs(蒙特卡洛次数) = {rp.get('蒙特卡洛次数')}，"
          f"Number to Save = {rp.get('保存数量')}")
    print(f"     补偿器 = {rp.get('补偿器模式')}，分布 = {rp.get('统计分布')}")
    print("  4. 点 Run，观察窗口实时滚动每次试验。")
    print(f"  5. 结果建议另存到: {outdir}")
    print("====================================================================")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
