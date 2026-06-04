from __future__ import annotations

import argparse
import json
import re
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def write_text(path: Path, text: str) -> None:
    path.write_text(text, encoding="utf-8", newline="\n")


def update_json_file(path: Path, updater) -> None:
    payload = json.loads(path.read_text(encoding="utf-8"))
    updater(payload)
    write_text(path, json.dumps(payload, ensure_ascii=False, indent=2) + "\n")


def update_main_constants(path: Path, version: str) -> None:
    text = path.read_text(encoding="utf-8")
    text = re.sub(
        r'Public Const VZID_MAIN_VERSION As String = "[^"]+"',
        f'Public Const VZID_MAIN_VERSION As String = "{version}"',
        text,
    )
    text = re.sub(
        r'Public Const VZID_UPDATER_VERSION As String = "[^"]+"',
        f'Public Const VZID_UPDATER_VERSION As String = "{version}"',
        text,
    )
    write_text(path, text)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("version")
    parser.add_argument("--min-loader-version")
    parser.add_argument("--min-updater-version")
    args = parser.parse_args()

    version = args.version.strip()
    min_loader_version = (args.min_loader_version or version).strip()
    min_updater_version = (args.min_updater_version or version).strip()

    write_text(ROOT / "release" / "version.txt", version + "\n")
    write_text(ROOT / "release" / "min-loader-version.txt", min_loader_version + "\n")
    write_text(ROOT / "release" / "min-updater-version.txt", min_updater_version + "\n")

    update_json_file(
        ROOT / "config" / "config.template.json",
        lambda payload: payload.update(
            {
                "activeMainVersion": version,
                "activeUpdaterVersion": version,
            }
        ),
    )

    def update_manifest(payload: dict) -> None:
        payload["releaseVersion"] = version
        payload["minLoaderVersion"] = min_loader_version
        payload["minUpdaterVersion"] = min_updater_version
        payload["mainDownloadUrl"] = f"https://github.com/chassehiboux/Excel_VZID_Macros/releases/download/v{version}/MainVZID.xlam"
        payload["notesUrl"] = f"https://github.com/chassehiboux/Excel_VZID_Macros/releases/tag/v{version}"

    update_json_file(ROOT / "release" / "manifest.template.json", update_manifest)
    update_main_constants(ROOT / "src" / "main-vzid" / "modules" / "MainConstants.bas", version)

    print(f"release/version.txt -> {version}")
    print(f"release/min-loader-version.txt -> {min_loader_version}")
    print(f"release/min-updater-version.txt -> {min_updater_version}")
    print("config/config.template.json -> updated")
    print("release/manifest.template.json -> updated")
    print("src/main-vzid/modules/MainConstants.bas -> updated")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
