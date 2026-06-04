# AGENTS.md

## О проекте

Это репозиторий новой Excel-надстройки `MainVZID`.

Главное правило:

- исходник надстройки лежит в репозитории;
- итоговый `.xlam` в `build/` или в `%LocalAppData%\VZID\addin\` не считается исходником.

## Где что править

- вкладка ленты: `src/main-vzid/customui/customUI14.xml`
- callbacks ленты: `src/main-vzid/modules/RibbonVZID.bas`
- привязка кнопки к макросу: `src/main-vzid/modules/CommandRegistry.bas`
- права по логинам: `src/main-vzid/modules/AccessPolicy.bas`
- стартовый конфиг прав и настроек: `config/config.template.json`
- `ThisWorkbook`: `src/main-vzid/workbook/ThisWorkbook.cls`
- формы: `src/main-vzid/forms/`

## Правила доработки

- Если пользователь просит новую фичу, изменение логики, изменение интерфейса или архитектурное решение, сначала согласовать с ним все важные моменты простым человеческим языком.
- Объяснять все так, как будто пользователь вообще не знает Excel/VBA/Office RibbonX.
- Если есть решение лучше, чем предложил пользователь, обязательно предложить его и коротко объяснить, почему оно лучше.
- После согласования не останавливаться на плане: сразу вносить доработку.
- Не тратить токены на лишние промежуточные рассуждения, повторные проверки и длинные тестовые прогоны без прямой пользы.
- Если пользователь не просит отдельный этап тестирования, делать только минимальные точечные проверки, нужные для сборки и релиза.

## Правила работы с ribbon

- Не считать `Office RibbonX Editor` основным местом разработки.
- Основной способ менять ленту: править `src/main-vzid/customui/customUI14.xml`.
- `Office RibbonX Editor` использовать в основном для просмотра и разовой инспекции уже собранного `.xlam`.
- Новая кнопка обычно требует изменения сразу в 3 местах:
  - `customUI14.xml`
  - `CommandRegistry.bas`
  - `AccessPolicy.bas` и `config.template.json`, если нужны права

## Локальная сборка

Тестовая локальная сборка:

```powershell
python scripts/build_addins.py --output-dir build_manual_check
```

Файл для ручной проверки:

- `build_manual_check/MainVZID.xlam`

## Правильный релиз

Если нужно выпустить новую версию, использовать такой порядок:

1. Поднять версию:

```powershell
python scripts/set_release_version.py 0.2.3
```

2. При необходимости собрать тестовый `.xlam`:

```powershell
python scripts/build_addins.py --output-dir build_manual_check
```

3. Закоммитить и запушить код в `main`.

4. Выпустить релиз:

```powershell
python scripts/publish_release.py
```

Что делает `publish_release.py`:

- собирает `MainVZID.xlam`
- собирает `manifest.json`
- собирает `setup.exe`
- создает или обновляет GitHub Release
- загружает в релиз:
  - `MainVZID.xlam`
  - `manifest.json`
  - `setup.exe`

## Что не делать

- Не редактировать установленный `MainVZID.xlam` как главный исходник.
- Не редактировать `build/MainVZID.xlam` как главный исходник.
- Не выпускать релиз вручную через случайные действия, если можно использовать `scripts/publish_release.py`.
- Не навязывать пользователю сложные технические термины без простого объяснения.
