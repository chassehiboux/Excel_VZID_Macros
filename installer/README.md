# Установщик новой надстройки VZID

## Цель

`setup.exe` должен установить продукт без прав администратора и без использования `XLSTART` как механизма самоподмены.

## Что должен сделать установщик

1. Создать папки:
   `%LocalAppData%\VZID\loader\`
   `%LocalAppData%\VZID\versions\current\`
   `%LocalAppData%\VZID\versions\pending\`
   `%LocalAppData%\VZID\config\`
   `%LocalAppData%\VZID\logs\`
2. Положить загрузчик в `%LocalAppData%\VZID\loader\LoaderVZID.xlam`.
3. Положить стартовую рабочую надстройку в `%LocalAppData%\VZID\versions\current\MainVZID.xlam`.
4. Положить стартовый `config.json` в `%LocalAppData%\VZID\config\config.json`.
5. Подключить `LoaderVZID.xlam` как обычную пользовательскую надстройку Excel.

## Ограничения MVP

- Обновление самого загрузчика не автоматизируется в первом прототипе.
- Основной сценарий обновления касается `MainVZID.xlam`.
- При необходимости новый загрузчик ставится повторным запуском `setup.exe`.
