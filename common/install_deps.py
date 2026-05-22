#!/usr/bin/env python3
"""检测并安装全部依赖（不安装 Python）。"""

from __future__ import annotations

import importlib.util
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
REQ = ROOT / "requirements.txt"
MODULES = ("fitz", "docx", "openpyxl", "pptx")


def has_module(name: str) -> bool:
    return importlib.util.find_spec(name) is not None


def main() -> int:
    if not REQ.is_file():
        print(f"ERROR: not found {REQ}")
        return 1

    print("=" * 42)
    print(f"Python: {sys.version.split()[0]}")
    print(f"Exe: {sys.executable}")
    print("=" * 42)

    if all(has_module(m) for m in MODULES):
        print("[SKIP] All dependencies OK:")
        for n in MODULES:
            print(f"  + {n}")
        return 0

    missing = [m for m in MODULES if not has_module(m)]
    print(f"[INSTALL] Missing: {', '.join(missing)}")
    subprocess.run([sys.executable, "-m", "pip", "install", "--upgrade", "pip"], check=False)
    subprocess.run(
        [sys.executable, "-m", "pip", "install", "-r", str(REQ)],
        check=True,
    )
    still = [m for m in MODULES if not has_module(m)]
    if still:
        print(f"ERROR: still missing {still}", file=sys.stderr)
        return 1
    print("[DONE] Install OK.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
