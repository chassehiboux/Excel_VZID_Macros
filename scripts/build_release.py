from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def read_version(version_arg: str | None) -> str:
    if version_arg:
        return version_arg.strip()
    return (ROOT / "release" / "version.txt").read_text(encoding="utf-8").strip()


def run_script(script_name: str, *args: str) -> None:
    subprocess.run([sys.executable, str(ROOT / "scripts" / script_name), *args], check=True)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--version")
    args = parser.parse_args()

    version_text = read_version(args.version)
    build_dir = ROOT / "build"
    release_dir = build_dir / "release"
    release_dir.mkdir(parents=True, exist_ok=True)

    run_script("build_addins.py", "--output-dir", str(build_dir))
    run_script("generate_manifest.py", "--version", version_text, "--output", str(release_dir / "manifest.json"))
    run_script("build_setup.py", "--dist-dir", str(release_dir))

    shutil.copy2(build_dir / "LoaderVZID.xlam", release_dir / "LoaderVZID.xlam")
    shutil.copy2(build_dir / "MainVZID.xlam", release_dir / "MainVZID.xlam")

    print(release_dir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
