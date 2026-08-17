#!/usr/bin/env python3
"""設計書レビュー用スクリプトの利用者向け入口。"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path


MINIMUM_PYTHON = (3, 12)
COMMANDS = {"preflight", "prepare", "finalize"}


def usage() -> str:
    return (
        "使用方法:\n"
        "  py -3 review.py preflight\n"
        "  py -3 review.py prepare [--run-id yyyyMMddhhmm]\n"
        "  py -3 review.py finalize --run-id yyyyMMddhhmm "
        "--results work/review-runs/yyyyMMddhhmm/results.json\n"
    )


def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if sys.version_info < MINIMUM_PYTHON:
        print("エラー: Python 3.12 以上が必要です。", file=sys.stderr)
        return 2

    if not args or args[0] in {"-h", "--help"}:
        print(usage())
        return 0

    command = args.pop(0)
    if command not in COMMANDS:
        print(f"エラー: 未対応のコマンドです: {command}", file=sys.stderr)
        print(usage(), file=sys.stderr)
        return 2

    repo_root = Path(__file__).resolve().parent
    skill_root = repo_root / ".agents" / "skills" / "review-documents-orchestrator" / "scripts"
    script = skill_root / ("preflight.py" if command == "preflight" else "orchestrate.py")
    if not script.is_file():
        print(f"エラー: 統合スクリプトが見つかりません: {script}", file=sys.stderr)
        return 2

    forwarded = ([] if command == "preflight" else [command]) + args
    if not any(item == "--root" or item.startswith("--root=") for item in forwarded):
        forwarded.extend(["--root", str(repo_root)])
    completed = subprocess.run([sys.executable, str(script), *forwarded], check=False)
    return completed.returncode


if __name__ == "__main__":
    raise SystemExit(main())
