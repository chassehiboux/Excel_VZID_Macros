from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dist-dir", default=str(ROOT / "build"))
    args = parser.parse_args()

    dist_dir = Path(args.dist_dir).resolve()
    dist_dir.mkdir(parents=True, exist_ok=True)

    pyi_root = ROOT / ".pyinstaller"
    work_dir = pyi_root / "work"
    spec_dir = pyi_root / "spec"
    work_dir.mkdir(parents=True, exist_ok=True)
    spec_dir.mkdir(parents=True, exist_ok=True)

    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--name",
        "updater",
        "--distpath",
        str(dist_dir),
        "--workpath",
        str(work_dir),
        "--specpath",
        str(spec_dir),
        str((ROOT / "updater" / "vzid_updater.py").resolve()),
    ]

    subprocess.run(command, check=True)
    print(dist_dir / "updater.exe")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
