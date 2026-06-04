# Установщик новой надстройки VZID

## Цель

`setup.exe` должен установить продукт без прав администратора и без использования `XLSTART` как механизма самоподмены.

Сборка `setup.exe` делается через пользовательскую цепочку:

- `python`;
- `pip`;
- `PyInstaller`.

## Что должен сделать установщик

1. Создать папки:
   `%AppData%\Microsoft\Excel\LocalCache\VZID\addin\`
   `%AppData%\Microsoft\Excel\LocalCache\VZID\updater\`
   `%AppData%\Microsoft\Excel\LocalCache\VZID\updates\`
   `%AppData%\Microsoft\Excel\LocalCache\VZID\backup\`
   `%AppData%\Microsoft\Excel\LocalCache\VZID\config\`
   `%AppData%\Microsoft\Excel\LocalCache\VZID\logs\`
2. Положить рабочую надстройку в `%AppData%\Microsoft\Excel\LocalCache\VZID\addin\MainVZID.xlam`.
3. Положить локальный `updater.exe` в `%AppData%\Microsoft\Excel\LocalCache\VZID\updater\updater.exe`.
4. Положить стартовый `config.json` в `%AppData%\Microsoft\Excel\LocalCache\VZID\config\config.json`.
5. Подключить `MainVZID.xlam` как обычную пользовательскую надстройку Excel.

Как это делается в MVP:

- установщик пишет путь к `MainVZID.xlam` в пользовательский раздел Excel:
  `HKCU\Software\Microsoft\Office\<версия>\Excel\Options`
- используется свободный слот `OPEN` / `OPEN1` / `OPEN2` и так далее;
- за счёт этого Excel сам подхватывает надстройку при следующем полном запуске;
- старые записи для `LoaderVZID.xlam` и `MainVZID.xlam` предварительно вычищаются;
- старые файлы `LoaderVZID.xlam` в `XLSTART` и старые `MainVZID.xlam` в `LocalCache` удаляются;
- старые надстройки дополнительно отключаются в Excel, чтобы не мешать новой установке.

## Ограничения MVP

- Обновление самого `updater.exe` не автоматизируется в первом прототипе.
- Основной сценарий обновления касается `MainVZID.xlam`.
- При необходимости новый `updater.exe` ставится повторным запуском `setup.exe`.

## Команды сборки

```powershell
python -m pip install --user -r requirements-build.txt
python scripts/build_release.py
```
