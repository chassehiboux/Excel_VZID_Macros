from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import time
from datetime import datetime
from pathlib import Path
from urllib import request


CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", required=True)
    parser.add_argument("--target", required=True)
    parser.add_argument("--config", required=True)
    parser.add_argument("--backup-dir", required=True)
    parser.add_argument("--log", required=True)
    parser.add_argument("--version", required=True)
    parser.add_argument("--mode", default="main", choices=("main", "setup"))
    parser.add_argument("--expected-sha256", default="")
    parser.add_argument("--download-url", default="")
    parser.add_argument("--restart-excel", default="0")
    parser.add_argument("--skip-wait-for-excel", action="store_true")
    return parser.parse_args()


def write_log(log_path: Path, message: str) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    line = f"{datetime.now():%Y-%m-%d %H:%M:%S} {message}\n"
    with log_path.open("a", encoding="utf-8") as handle:
        handle.write(line)


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while True:
            chunk = handle.read(1024 * 1024)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def verify_expected_sha256(path: Path, expected_sha256: str) -> None:
    expected_sha256 = expected_sha256.strip().lower()
    if not expected_sha256:
        return

    actual_sha256 = file_sha256(path).lower()
    if actual_sha256 != expected_sha256:
        raise RuntimeError(
            f"Контрольная сумма файла обновления не совпала. Ожидалось {expected_sha256}, получено {actual_sha256}."
        )


def excel_is_running() -> bool:
    result = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
        capture_output=True,
        text=True,
        errors="ignore",
        creationflags=CREATE_NO_WINDOW,
        check=False,
    )
    return "EXCEL.EXE" in result.stdout.upper()


def wait_for_excel_close(log_path: Path) -> None:
    announced = False
    while excel_is_running():
        if not announced:
            write_log(log_path, "Ожидание полного закрытия Excel перед обновлением.")
            announced = True
        time.sleep(2)


def load_config(config_path: Path) -> dict:
    try:
        return json.loads(config_path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_config(config_path: Path, payload: dict) -> None:
    config_path.parent.mkdir(parents=True, exist_ok=True)
    config_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def update_config_success(config_path: Path, version_text: str) -> None:
    payload = load_config(config_path)
    payload["activeMainVersion"] = version_text
    payload["availableMainVersion"] = ""
    payload["availableMainDownloadUrl"] = ""
    payload["preparedMainVersion"] = ""
    payload["preparedMainPath"] = ""
    payload["lastUpdateStatus"] = "up_to_date"
    payload["lastUpdateMessage"] = f"Обновление {version_text} установлено."
    save_config(config_path, payload)


def update_config_failure(config_path: Path, message: str) -> None:
    payload = load_config(config_path)
    payload["lastUpdateStatus"] = "activation_failed"
    payload["lastUpdateMessage"] = message
    save_config(config_path, payload)


def relaunch_excel(log_path: Path) -> None:
    try:
        os.startfile("excel.exe")  # type: ignore[attr-defined]
        write_log(log_path, "Excel запущен повторно после обновления.")
    except Exception as exc:
        write_log(log_path, f"Не удалось запустить Excel повторно: {exc}")


def download_file(url: str, target_path: Path, log_path: Path, version_text: str) -> None:
    temp_path = target_path.with_name(target_path.name + ".download")
    target_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        if temp_path.exists():
            temp_path.unlink()

        req = request.Request(url, headers={"User-Agent": f"VZID-Updater/{version_text}"})
        with request.urlopen(req, timeout=60) as response, temp_path.open("wb") as handle:
            while True:
                chunk = response.read(1024 * 1024)
                if not chunk:
                    break
                handle.write(chunk)

        os.replace(temp_path, target_path)
        write_log(log_path, f"Файл скачан: {target_path}")
    except Exception:
        try:
            if temp_path.exists():
                temp_path.unlink()
        except Exception:
            pass
        raise


def install_update(args: argparse.Namespace) -> int:
    source_path = Path(args.source).resolve()
    target_path = Path(args.target).resolve()
    config_path = Path(args.config).resolve()
    backup_dir = Path(args.backup_dir).resolve()
    log_path = Path(args.log).resolve()
    expected_sha256 = args.expected_sha256.strip().lower()
    restart_excel = str(args.restart_excel).strip() in {"1", "true", "True"}

    try:
        write_log(log_path, f"Updater started in mode={args.mode} for version {args.version}.")
        if not args.skip_wait_for_excel:
            wait_for_excel_close(log_path)

        if args.mode == "setup":
            return run_setup_update(
                source_path=source_path,
                download_url=args.download_url.strip(),
                expected_sha256=expected_sha256,
                config_path=config_path,
                log_path=log_path,
                restart_excel=restart_excel,
                version_text=args.version,
            )

        if not source_path.exists():
            raise FileNotFoundError(f"Файл обновления не найден: {source_path}")

        verify_expected_sha256(source_path, expected_sha256)

        target_path.parent.mkdir(parents=True, exist_ok=True)
        backup_dir.mkdir(parents=True, exist_ok=True)

        if target_path.exists():
            backup_name = f"{target_path.stem}-{datetime.now():%Y%m%d-%H%M%S}{target_path.suffix}"
            backup_path = backup_dir / backup_name
            shutil.copy2(target_path, backup_path)
            write_log(log_path, f"Создана резервная копия: {backup_path}")

        temp_target = target_path.with_name(target_path.name + ".new")
        if temp_target.exists():
            temp_target.unlink()

        shutil.copy2(source_path, temp_target)
        os.replace(temp_target, target_path)
        write_log(log_path, f"Файл надстройки обновлен: {target_path}")

        try:
            source_path.unlink()
        except FileNotFoundError:
            pass

        update_config_success(config_path, args.version)
        write_log(log_path, f"Конфиг обновлен: активная версия {args.version}")

        if restart_excel:
            relaunch_excel(log_path)

        return 0
    except Exception as exc:
        update_config_failure(config_path, f"Не удалось применить обновление: {exc}")
        write_log(log_path, f"Обновление завершилось ошибкой: {exc}")
        return 1


def run_setup_update(
    *,
    source_path: Path,
    download_url: str,
    expected_sha256: str,
    config_path: Path,
    log_path: Path,
    restart_excel: bool,
    version_text: str,
) -> int:
    try:
        if download_url:
            write_log(log_path, f"Скачивание setup.exe из {download_url}")
            download_file(download_url, source_path, log_path, version_text)
        elif not source_path.exists():
            raise FileNotFoundError(f"Файл setup.exe не найден: {source_path}")

        verify_expected_sha256(source_path, expected_sha256)
        write_log(log_path, f"Запуск setup.exe: {source_path}")
        result = subprocess.run(
            [str(source_path), "--no-ui"],
            check=False,
            creationflags=CREATE_NO_WINDOW,
        )
        if result.returncode != 0:
            raise RuntimeError(f"setup.exe завершился с кодом {result.returncode}")

        try:
            source_path.unlink()
        except FileNotFoundError:
            pass

        write_log(log_path, "setup.exe успешно завершил обновление.")
        if restart_excel:
            relaunch_excel(log_path)
        return 0
    except Exception as exc:
        update_config_failure(config_path, f"Не удалось применить setup.exe: {exc}")
        write_log(log_path, f"Ошибка запуска setup.exe: {exc}")
        return 1


def main() -> int:
    return install_update(parse_args())


if __name__ == "__main__":
    raise SystemExit(main())
