from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def read_release_value(arg_value: str | None, file_name: str) -> str:
    if arg_value:
        return arg_value.strip()
    return (ROOT / "release" / file_name).read_text(encoding="utf-8").strip()


def run_script(script_name: str, *args: str) -> None:
    subprocess.run([sys.executable, str(ROOT / "scripts" / script_name), *args], check=True)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--version")
    parser.add_argument("--min-loader-version")
    parser.add_argument("--min-updater-version")
    args = parser.parse_args()

    version_text = read_release_value(args.version, "version.txt")
    min_loader_version = read_release_value(args.min_loader_version, "min-loader-version.txt")
    min_updater_version = read_release_value(args.min_updater_version, "min-updater-version.txt")
    build_dir = ROOT / "build"
    release_dir = build_dir / "release"
    release_dir.mkdir(parents=True, exist_ok=True)

    run_script("build_addins.py", "--output-dir", str(build_dir))
    run_script("build_updater.py", "--dist-dir", str(build_dir))
    run_script("build_setup.py", "--dist-dir", str(release_dir))
    run_script(
        "generate_manifest.py",
        "--version",
        version_text,
        "--min-loader-version",
        min_loader_version,
        "--min-updater-version",
        min_updater_version,
        "--setup-exe",
        str(release_dir / "setup.exe"),
        "--output",
        str(release_dir / "manifest.json"),
    )

    shutil.copy2(build_dir / "MainVZID.xlam", release_dir / "MainVZID.xlam")

    legacy_loader_release = release_dir / "LoaderVZID.xlam"
    if legacy_loader_release.exists():
        legacy_loader_release.unlink()

    print(release_dir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
