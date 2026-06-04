from __future__ import annotations

import argparse
import base64
import json
import mimetypes
import subprocess
import sys
from pathlib import Path
from urllib import error, parse, request


ROOT = Path(__file__).resolve().parents[1]
REPO = "chassehiboux/Excel_VZID_Macros"


def git_credential() -> tuple[str, str]:
    result = subprocess.run(
        ["git", "credential", "fill"],
        input="protocol=https\nhost=github.com\n\n",
        text=True,
        capture_output=True,
        check=True,
    )

    payload: dict[str, str] = {}
    for line in result.stdout.splitlines():
        if "=" not in line:
            continue
        key, value = line.split("=", 1)
        payload[key.strip()] = value.strip()

    username = payload.get("username", "")
    password = payload.get("password", "")
    if not username or not password:
        raise RuntimeError("GitHub credentials not found via git credential fill.")
    return username, password


def api_headers() -> dict[str, str]:
    username, password = git_credential()
    auth = base64.b64encode(f"{username}:{password}".encode("utf-8")).decode("ascii")
    return {
        "Authorization": f"Basic {auth}",
        "Accept": "application/vnd.github+json",
        "User-Agent": "vzid-publish-release",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def api_json(method: str, url: str, headers: dict[str, str], payload=None, expected=(200, 201)):
    request_headers = dict(headers)
    data = None
    if payload is not None:
        data = json.dumps(payload).encode("utf-8")
        request_headers["Content-Type"] = "application/json; charset=utf-8"

    req = request.Request(url, data=data, headers=request_headers, method=method)
    try:
        with request.urlopen(req) as response:
            body = response.read()
            if response.status not in expected:
                raise RuntimeError(f"{method} {url} returned {response.status}")
            return json.loads(body.decode("utf-8")) if body else None
    except error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        if exc.code == 404 and method == "GET":
            return None
        raise RuntimeError(f"{method} {url} failed: {exc.code} {body}") from exc


def api_raw(method: str, url: str, headers: dict[str, str], data: bytes = b"", expected=(200, 201, 204)):
    req = request.Request(url, data=data, headers=headers, method=method)
    try:
        with request.urlopen(req) as response:
            body = response.read()
            if response.status not in expected:
                raise RuntimeError(f"{method} {url} returned {response.status}")
            return body
    except error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"{method} {url} failed: {exc.code} {body}") from exc


def read_version() -> str:
    return (ROOT / "release" / "version.txt").read_text(encoding="utf-8").strip()


def build_release() -> None:
    subprocess.run([sys.executable, str(ROOT / "scripts" / "build_release.py")], check=True)


def release_notes(version: str) -> str:
    return (
        f"Релиз {version}\n\n"
        "Состав релиза:\n"
        "- MainVZID.xlam\n"
        "- manifest.json\n"
        "- setup.exe\n"
    )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--skip-build", action="store_true")
    parser.add_argument("--target-commitish", default="main")
    parser.add_argument("--notes-file")
    args = parser.parse_args()

    version = read_version()
    tag = f"v{version}"

    if not args.skip_build:
        build_release()

    assets = [
        ROOT / "build" / "MainVZID.xlam",
        ROOT / "build" / "release" / "manifest.json",
        ROOT / "build" / "release" / "setup.exe",
    ]
    for asset in assets:
        if not asset.exists():
            raise FileNotFoundError(f"Missing asset: {asset}")

    notes = release_notes(version)
    if args.notes_file:
        notes = Path(args.notes_file).read_text(encoding="utf-8")

    headers = api_headers()
    release = api_json("GET", f"https://api.github.com/repos/{REPO}/releases/tags/{tag}", headers)

    if release is None:
        release = api_json(
            "POST",
            f"https://api.github.com/repos/{REPO}/releases",
            headers,
            {
                "tag_name": tag,
                "target_commitish": args.target_commitish,
                "name": tag,
                "body": notes,
                "draft": False,
                "prerelease": False,
                "generate_release_notes": False,
            },
        )
    else:
        release = api_json(
            "PATCH",
            f"https://api.github.com/repos/{REPO}/releases/{release['id']}",
            headers,
            {
                "tag_name": tag,
                "target_commitish": args.target_commitish,
                "name": tag,
                "body": notes,
                "draft": False,
                "prerelease": False,
            },
            expected=(200,),
        )

    asset_names = {asset.name for asset in assets}
    for asset in release.get("assets", []):
        if asset.get("name") not in asset_names:
            continue
        api_raw(
            "DELETE",
            f"https://api.github.com/repos/{REPO}/releases/assets/{asset['id']}",
            headers,
            expected=(204,),
        )

    upload_url = release["upload_url"].split("{", 1)[0]
    for asset in assets:
        content_type = mimetypes.guess_type(asset.name)[0] or "application/octet-stream"
        upload_headers = dict(headers)
        upload_headers["Content-Type"] = content_type
        api_raw(
            "POST",
            f"{upload_url}?name={parse.quote(asset.name)}",
            upload_headers,
            data=asset.read_bytes(),
            expected=(201,),
        )

    final_release = api_json("GET", f"https://api.github.com/repos/{REPO}/releases/tags/{tag}", headers, expected=(200,))
    print(final_release["html_url"])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
