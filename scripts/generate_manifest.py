from __future__ import annotations

import argparse
import hashlib
import json
from datetime import datetime, timezone
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
REPO_URL = "https://github.com/chassehiboux/Excel_VZID_Macros"


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while True:
            chunk = handle.read(1024 * 1024)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--version", required=True)
    parser.add_argument("--min-loader-version", required=True)
    parser.add_argument("--main-xlam", default=str(ROOT / "build" / "MainVZID.xlam"))
    parser.add_argument("--output", default=str(ROOT / "build" / "manifest.json"))
    args = parser.parse_args()

    main_xlam_path = Path(args.main_xlam).resolve()
    output_path = Path(args.output).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    manifest = {
        "schemaVersion": "1",
        "releaseVersion": args.version,
        "publishedAt": datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z"),
        "minLoaderVersion": args.min_loader_version,
        "mainDownloadUrl": f"{REPO_URL}/releases/download/v{args.version}/MainVZID.xlam",
        "mainSha256": file_sha256(main_xlam_path),
        "notesUrl": f"{REPO_URL}/releases/tag/v{args.version}",
    }

    output_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(output_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
