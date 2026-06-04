from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--loader-xlam", default=str(ROOT / "build" / "LoaderVZID.xlam"))
    parser.add_argument("--main-xlam", default=str(ROOT / "build" / "MainVZID.xlam"))
    parser.add_argument("--config-template", default=str(ROOT / "config" / "config.template.json"))
    parser.add_argument("--dist-dir", default=str(ROOT / "build" / "release"))
    args = parser.parse_args()

    loader_xlam = Path(args.loader_xlam).resolve()
    main_xlam = Path(args.main_xlam).resolve()
    config_template = Path(args.config_template).resolve()
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
        "setup",
        "--distpath",
        str(dist_dir),
        "--workpath",
        str(work_dir),
        "--specpath",
        str(spec_dir),
        "--add-data",
        f"{loader_xlam};.",
        "--add-data",
        f"{main_xlam};.",
        "--add-data",
        f"{config_template};config",
        str((ROOT / "installer" / "vzid_setup.py").resolve()),
    ]

    subprocess.run(command, check=True)
    print(dist_dir / "setup.exe")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
