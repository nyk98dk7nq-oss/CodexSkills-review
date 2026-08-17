#!/usr/bin/env python3
"""全Skillの単体テストとリポジトリ構成テストを順番に実行する。"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def main() -> int:
    root = Path(__file__).resolve().parent
    suites = [root / "tests"]
    suites.extend(sorted((root / ".agents" / "skills").glob("*/tests")))
    failures: list[Path] = []

    for suite in suites:
        if not suite.is_dir():
            continue
        print(f"\n=== {suite.relative_to(root)} ===", flush=True)
        completed = subprocess.run(
            [sys.executable, "-m", "unittest", "discover", "-s", str(suite), "-p", "test_*.py", "-v"],
            cwd=root,
            check=False,
        )
        if completed.returncode:
            failures.append(suite)

    if failures:
        print("\n失敗したテスト:", file=sys.stderr)
        for suite in failures:
            print(f"- {suite.relative_to(root)}", file=sys.stderr)
        return 1

    print("\nすべてのテストに合格しました。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
