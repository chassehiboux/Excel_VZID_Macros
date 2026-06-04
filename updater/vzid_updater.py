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


CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", required=True)
    parser.add_argument("--target", required=True)
    parser.add_argument("--config", required=True)
    parser.add_argument("--backup-dir", required=True)
    parser.add_argument("--log", required=True)
    parser.add_argument("--version", required=True)
    parser.add_argument("--expected-sha256", default="")
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


def install_update(args: argparse.Namespace) -> int:
    source_path = Path(args.source).resolve()
    target_path = Path(args.target).resolve()
    config_path = Path(args.config).resolve()
    backup_dir = Path(args.backup_dir).resolve()
    log_path = Path(args.log).resolve()
    expected_sha256 = args.expected_sha256.strip().lower()
    restart_excel = str(args.restart_excel).strip() in {"1", "true", "True"}

    try:
        write_log(log_path, f"Updater started for {target_path.name}, target version {args.version}.")
        if not args.skip_wait_for_excel:
            wait_for_excel_close(log_path)

        if not source_path.exists():
            raise FileNotFoundError(f"Файл обновления не найден: {source_path}")

        if expected_sha256:
            actual_sha256 = file_sha256(source_path).lower()
            if actual_sha256 != expected_sha256:
                raise RuntimeError("Контрольная сумма файла обновления не совпала.")

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


def main() -> int:
    return install_update(parse_args())


if __name__ == "__main__":
    raise SystemExit(main())
