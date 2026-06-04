from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
import winreg
from pathlib import Path
from tkinter import Tk, messagebox

import pythoncom
import win32com.client


APP_NAME = "VZID"
LEGACY_LOADER_FILE = "LoaderVZID.xlam"
MAIN_FILE = "MainVZID.xlam"
UPDATER_FILE = "updater.exe"
CONFIG_TEMPLATE = "config.template.json"
INSTALL_BASE_DIR: Path | None = None
NO_UI = False


def base_dir() -> Path:
    if INSTALL_BASE_DIR is not None:
        return INSTALL_BASE_DIR

    root = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
    if not root:
        root = str(Path.home())
    return Path(root) / APP_NAME


def addin_dir() -> Path:
    return base_dir() / "addin"


def updater_dir() -> Path:
    return base_dir() / "updater"


def updates_dir() -> Path:
    return base_dir() / "updates"


def backup_dir() -> Path:
    return base_dir() / "backup"


def legacy_loader_dir() -> Path:
    return base_dir() / "loader"


def legacy_versions_dir() -> Path:
    return base_dir() / "versions"


def config_dir() -> Path:
    return base_dir() / "config"


def logs_dir() -> Path:
    return base_dir() / "logs"


def config_path() -> Path:
    return config_dir() / "config.json"


def config_template_path() -> Path:
    return bundled_path(f"config/{CONFIG_TEMPLATE}")


def main_path() -> Path:
    return addin_dir() / MAIN_FILE


def updater_path() -> Path:
    return updater_dir() / UPDATER_FILE


def bundled_path(relative_path: str) -> Path:
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative_path

    repo_root = Path(__file__).resolve().parents[1]
    direct_path = repo_root / relative_path
    if direct_path.exists():
        return direct_path

    build_path = repo_root / "build" / Path(relative_path).name
    if build_path.exists():
        return build_path

    return direct_path


def ensure_dirs() -> None:
    for path in (addin_dir(), updater_dir(), updates_dir(), backup_dir(), config_dir(), logs_dir()):
        path.mkdir(parents=True, exist_ok=True)


def load_json_file(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_json_file(path: Path, payload: dict) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def clear_prepared_updates() -> None:
    for prepared_file in updates_dir().glob("MainVZID*.xlam"):
        try:
            prepared_file.unlink()
        except FileNotFoundError:
            pass


def cleanup_legacy_layout() -> None:
    legacy_files = [
        legacy_loader_dir() / LEGACY_LOADER_FILE,
        legacy_versions_dir() / "current" / MAIN_FILE,
    ]

    for legacy_file in legacy_files:
        try:
            legacy_file.unlink()
        except FileNotFoundError:
            pass

    for pending_file in (legacy_versions_dir() / "pending").glob("MainVZID*.xlam"):
        try:
            pending_file.unlink()
        except FileNotFoundError:
            pass


def sync_config() -> None:
    template = load_json_file(config_template_path())
    current = load_json_file(config_path()) if config_path().exists() else {}

    merged = dict(template)
    if isinstance(current, dict):
        merged.update(current)

    for obsolete_key in ("activeLoaderVersion", "pendingMainVersion", "pendingMainPath"):
        merged.pop(obsolete_key, None)

    active_main_version = str(template.get("activeMainVersion", "0.0.0"))
    active_updater_version = str(template.get("activeUpdaterVersion", "0.0.0"))

    merged["schemaVersion"] = str(template.get("schemaVersion", "2"))
    merged["activeMainVersion"] = active_main_version
    merged["activeUpdaterVersion"] = active_updater_version
    merged["availableMainVersion"] = ""
    merged["availableMainDownloadUrl"] = ""
    merged["preparedMainVersion"] = ""
    merged["preparedMainPath"] = ""
    merged["lastUpdateCheckAt"] = ""
    merged["lastUpdateStatus"] = "up_to_date"
    merged["lastUpdateMessage"] = f"Установлена версия {active_main_version} через setup.exe."

    save_json_file(config_path(), merged)


def copy_assets() -> None:
    shutil.copy2(bundled_path(MAIN_FILE), main_path())
    shutil.copy2(bundled_path(f"updater/{UPDATER_FILE}"), updater_path())
    if not config_path().exists():
        shutil.copy2(config_template_path(), config_path())

    clear_prepared_updates()
    cleanup_legacy_layout()
    sync_config()


def register_main_addin() -> None:
    version = detect_excel_version()
    register_main_addin_in_registry(version, main_path())


def detect_excel_version() -> str:
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        return str(excel.Version)
    finally:
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


def register_main_addin_in_registry(excel_version: str, addin_path: Path) -> None:
    options_key_path = f"Software\\Microsoft\\Office\\{excel_version}\\Excel\\Options"
    value_to_store = f'"{addin_path.resolve()}"'

    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, options_key_path) as key:
        existing_values: list[str] = []
        existing_names: list[str] = []
        index = 0

        while index < 64:
            value_name = "OPEN" if index == 0 else f"OPEN{index}"
            try:
                current_value, _ = winreg.QueryValueEx(key, value_name)
            except FileNotFoundError:
                index += 1
                continue

            existing_names.append(value_name)
            existing_values.append(str(current_value))
            index += 1

        for value_name in existing_names:
            try:
                winreg.DeleteValue(key, value_name)
            except FileNotFoundError:
                pass

        preserved_values = []
        for current_value in existing_values:
            candidate_name = startup_entry_basename(current_value)
            if candidate_name in {LEGACY_LOADER_FILE.lower(), MAIN_FILE.lower()}:
                continue
            preserved_values.append(current_value)

        preserved_values.append(value_to_store)

        for index, current_value in enumerate(preserved_values):
            value_name = "OPEN" if index == 0 else f"OPEN{index}"
            winreg.SetValueEx(key, value_name, 0, winreg.REG_SZ, current_value)


def startup_entry_basename(value: str) -> str:
    normalized_value = str(value).strip()
    if normalized_value.count('"') >= 2:
        first_quote = normalized_value.find('"')
        last_quote = normalized_value.rfind('"')
        if last_quote > first_quote:
            normalized_value = normalized_value[first_quote + 1:last_quote]
    else:
        normalized_value = normalized_value.strip('"')

    return Path(normalized_value).name.lower()


def show_info(title: str, text: str) -> None:
    if NO_UI:
        print(f"{title}: {text}")
        return

    root = Tk()
    root.withdraw()
    messagebox.showinfo(title, text)
    root.destroy()


def show_error(title: str, text: str) -> None:
    if NO_UI:
        print(f"{title}: {text}")
        return

    root = Tk()
    root.withdraw()
    messagebox.showerror(title, text)
    root.destroy()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--target-root")
    parser.add_argument("--skip-register", action="store_true")
    parser.add_argument("--no-ui", action="store_true")
    return parser.parse_args()


def main() -> int:
    global INSTALL_BASE_DIR
    global NO_UI

    args = parse_args()
    NO_UI = args.no_ui
    if args.target_root:
        INSTALL_BASE_DIR = Path(args.target_root).resolve()

    try:
        ensure_dirs()
        copy_assets()
        if not args.skip_register:
            register_main_addin()

        show_info(
            "VZID Setup",
            "Надстройка установлена для текущего пользователя.\n\n"
            "Если Excel сейчас открыт, полностью закройте все его окна и откройте Excel заново.",
        )
        return 0
    except OSError as exc:  # pragma: no cover
        if getattr(exc, "winerror", None) == 32:
            show_error(
                "VZID Setup",
                "Не удалось заменить файлы надстройки, потому что Excel все еще держит их открытыми.\n\n"
                "Полностью закройте все окна Excel и запустите setup.exe снова.",
            )
            return 1

        show_error("VZID Setup", f"Установка завершилась ошибкой:\n{exc}")
        return 1
    except Exception as exc:  # pragma: no cover
        show_error("VZID Setup", f"Установка завершилась ошибкой:\n{exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
