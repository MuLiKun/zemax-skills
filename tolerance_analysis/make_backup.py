from __future__ import annotations

import argparse
import json
import shutil
from datetime import datetime
from pathlib import Path


DEFAULT_ITEMS = [
    "README.md",
    "tolerance_analysis/main.py",
    "tolerance_analysis/gui.py",
    "tolerance_analysis/tol_run.py",
    "tolerance_analysis/check_stage1.py",
    "tolerance_analysis/build.ps1",
    "tolerance_analysis/gui.spec",
    "tolerance_analysis/zemax_config.ini.example",
    "tolerance_analysis/tol_config_模板.xlsx",
    "tolerance_analysis/toltool",
    "tolerance_analysis/方案A_全自动",
    "tolerance_analysis/方案B_手动观察",
    "tolerance_analysis/公差分析程序_使用说明.md",
    "tolerance_analysis/公差分析程序_需求文档.md",
    "tolerance_analysis/公差分析工具_泛化需求文档.md",
    "tolerance_analysis/公差分析工具_开发进度与测试记录.md",
]


IGNORE_DIRS = {"__pycache__", ".pytest_cache"}
IGNORE_SUFFIXES = {".pyc", ".pyo"}


def _copy_item(src: Path, dst: Path) -> None:
    if not src.exists():
        return
    if src.is_dir():
        shutil.copytree(src, dst, ignore=_ignore_names)
    else:
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)


def _ignore_names(_dir: str, names: list[str]) -> set[str]:
    ignored: set[str] = set()
    for name in names:
        path = Path(name)
        if name in IGNORE_DIRS or path.suffix.lower() in IGNORE_SUFFIXES:
            ignored.add(name)
    return ignored


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="生成一阶段本地代码/文档备份快照")
    parser.add_argument("--outdir", default=None,
                        help="备份根目录，默认仓库根目录下的 _backups")
    args = parser.parse_args(argv)

    script_dir = Path(__file__).resolve().parent
    repo_root = script_dir.parent
    backup_root = Path(args.outdir).resolve() if args.outdir else repo_root / "_backups"
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_dir = backup_root / stamp
    backup_dir.mkdir(parents=True, exist_ok=False)

    copied: list[str] = []
    skipped: list[str] = []
    for item in DEFAULT_ITEMS:
        src = repo_root / item
        dst = backup_dir / item
        if src.exists():
            _copy_item(src, dst)
            copied.append(item)
        else:
            skipped.append(item)

    manifest = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "repo_root": str(repo_root),
        "backup_dir": str(backup_dir),
        "copied": copied,
        "skipped": skipped,
    }
    (backup_dir / "backup_manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"备份完成：{backup_dir}")
    print(f"已复制：{len(copied)} 项")
    if skipped:
        print(f"已跳过不存在项：{len(skipped)} 项")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
