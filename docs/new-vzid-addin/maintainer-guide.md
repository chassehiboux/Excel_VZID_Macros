# Гайд по доработке и выпуску MainVZID

## Главное правило

Исходник надстройки теперь не сам `.xlam`-файл.

Исходник теперь лежит в репозитории:

- XML вкладки: `src/main-vzid/customui/customUI14.xml`
- VBA-модули: `src/main-vzid/modules/`
- `ThisWorkbook`: `src/main-vzid/workbook/ThisWorkbook.cls`
- формы: `src/main-vzid/forms/`
- права по логинам: `config/config.template.json` и `AccessPolicy.bas`

Собранный `MainVZID.xlam` в `build/` или установленный в `%AppData%\Microsoft\Excel\LocalCache\VZID\addin\` больше не надо считать исходником. Это только результат сборки.

## Как теперь правильно мыслить

Старая схема была такая:

- есть живая `Dev_MainVZID.xlam`;
- ты правишь ее руками в Excel;
- потом раскатываешь готовый `.xlam`.

Новая схема такая:

- ты правишь файлы в репозитории;
- из них скрипт собирает новый `MainVZID.xlam`;
- этот файл тестируется;
- потом из него же собирается `setup.exe` и публикуется GitHub Release.

Если коротко:

- руками редактируем репозиторий;
- руками не редактируем итоговый `build/MainVZID.xlam`, если хотим, чтобы изменения не потерялись.

## Как дорабатывать вкладку в ленте

### Где лежит вкладка

Вкладка описана в файле:

- `src/main-vzid/customui/customUI14.xml`

Сейчас там уже есть:

- вкладка `VZID`;
- dropdown выбора региона с подписью сверху;
- блок `Формирование сопроводительных документов`;
- блок `Действия`;
- блок обновления.
- ручная раскладка по колонкам через `box`, а не автоматический перенос Excel.

### Что за что отвечает

Есть 4 слоя:

1. `customUI14.xml`
   Тут описано, какие кнопки вообще видны в ленте.
2. `RibbonVZID.bas`
   Тут описано, что Excel делает при нажатии и как он запрашивает состояние кнопок и видимость специальных блоков.
3. `CommandRegistry.bas`
   Тут описано, какой `tag` запускает какой макрос в каком регионе.
4. `AccessPolicy.bas` + `config.template.json`
   Тут описано, кому кнопка доступна.

## Как добавить новую кнопку

Допустим, ты хочешь добавить кнопку `Ручной отчет`.

### Шаг 1. Добавь кнопку в XML

Файл:

- `src/main-vzid/customui/customUI14.xml`

Пример строки:

```xml
<button
  id="btnManualProcReport"
  label="Ручной отчет"
  tag="MANUAL_PROC_REPORT"
  onAction="RibbonVZID_RunCommand"
  getEnabled="RibbonVZID_GetCommandEnabled"
  getVisible="RibbonVZID_GetCommandVisible" />
```

Что здесь значит:

- `id`
  внутренний уникальный идентификатор кнопки;
- `label`
  текст на кнопке;
- `tag`
  главный ключ команды, по нему дальше ищется макрос и право доступа;
- `onAction="RibbonVZID_RunCommand"`
  значит: нажимать будем через общий роутер;
- `getEnabled="RibbonVZID_GetCommandEnabled"`
  значит: Excel будет спрашивать, активна кнопка или нет.
- `getVisible="RibbonVZID_GetCommandVisible"`
  значит: Excel будет спрашивать, должна ли кнопка вообще показываться для выбранного региона.
- `screentip` и `supertip`
  это короткая и расширенная подсказка, которую пользователь видит при наведении на кнопку.

### Шаг 2. Привяжи кнопку к макросу

Файл:

- `src/main-vzid/modules/CommandRegistry.bas`

Добавь новый `Case`:

```vb
Case "MANUAL_PROC_REPORT"
    Select Case UCase$(regionId)
        Case "KGN"
            CommandRegistry_GetMacroName = "CommandHandlers_ShowManualProcReportKGN"
    End Select
```

Что это значит:

- если выбран регион `KGN`, то кнопка с `tag="MANUAL_PROC_REPORT"` запускает макрос `ManualProcReport`;
- если выбран регион `KGN`, то кнопка с `tag="MANUAL_PROC_REPORT"` запускает обертку `CommandHandlers_ShowManualProcReportKGN`, а уже она открывает нужную форму;
- если для другого региона запись не добавлена, кнопка там будет неактивной.

## Как назначить макрос на кнопку

В этой архитектуре макрос на кнопку назначается не в Office RibbonX Editor, а в `CommandRegistry.bas`.

То есть логика такая:

- в XML кнопка получает `tag`;
- в `CommandRegistry.bas` этот `tag` привязывается к имени макроса;
- `CommandDispatcher.bas` потом вызывает нужный макрос через `Application.Run`.

### Пример полной связки

#### В XML

```xml
<button id="btnDocPackets" label="Реестр пакетов в Excel" tag="DOC_PACKETS" onAction="RibbonVZID_RunCommand" getEnabled="RibbonVZID_GetCommandEnabled" getVisible="RibbonVZID_GetCommandVisible" />
```

#### В `CommandRegistry.bas`

```vb
Case "DOC_PACKETS"
    Select Case UCase$(regionId)
        Case "KGN", "TMN", "EKB", "CHLB"
            CommandRegistry_GetMacroName = "Doc_Packets"
    End Select
```

Итог:

- пользователь жмет кнопку;
- Excel вызывает `RibbonVZID_RunCommand`;
- код берет `control.Tag`;
- по `tag` и текущему региону ищется имя макроса;
- макрос запускается.

## Как выдать или ограничить права

### Где это хранится

Права лежат в двух местах:

1. `src/main-vzid/modules/AccessPolicy.bas`
2. `config/config.template.json`

### Шаг 1. Добавь ключ права в `AccessPolicy.bas`

Если команда новая, нужно сказать, какой JSON-ключ отвечает за доступ.

Пример:

```vb
Case "MANUAL_PROC_REPORT"
    AccessPolicy_CommandKey = "commandAccessManualProcReportCsv"
```

### Шаг 2. Добавь это поле в `config.template.json`

Пример:

```json
"commandAccessManualProcReportCsv": "*"
```

Что это значит:

- кнопку сможет нажимать любой пользователь;
- если нужен узкий доступ, вместо `*` ставишь один логин или список логинов через `;`.

### Как задаются права

- `"*"`: доступ всем;
- `"dzuikevich"`: доступ только одному логину;
- `"user1;user2;user3"`: доступ нескольким;
- `fullAccessUsersCsv`: логины с полным доступом ко всем кнопкам.

### Важно

Сейчас логика сделана так:

- если команда не поддерживается для региона, кнопка неактивна;
- если у пользователя нет прав, кнопка тоже неактивна;
- кнопка не скрывается, а именно показывается серой.

Это соответствует твоему требованию: показывать, но делать неактивными.

Отдельное правило для Тюмени:

- кнопки `Выбрать ОСП и Адрес` и `Горячая клавиша` видны только в регионе `Тюмень`;
- в остальных регионах они полностью скрыты, а не просто серые.

## Как теперь устроена раскладка

- лента не полагается на автоматический перенос кнопок Excel;
- внутри групп используются `box boxStyle="horizontal"` и `box boxStyle="vertical"`;
- это позволяет вручную задавать колонки;
- между колонками ставится `separator`, чтобы визуально отделить блоки кнопок.

Важно:

- нормального скролла внутри группы RibbonX нет;
- если кнопок станет слишком много, правильный путь не "добавить прокрутку", а пересобрать раскладку, сократить подписи или разбить команды на отдельные группы.

## Как пользоваться Office RibbonX Editor

### Самое важное

В этом проекте **Office RibbonX Editor не является главным местом редактирования**.

Главное место редактирования:

- `src/main-vzid/customui/customUI14.xml`

Почему:

- при сборке скрипт все равно заново вшивает XML в `.xlam`;
- если править ribbon прямо внутри готового `.xlam`, то при следующей сборке изменения перезапишутся.

### Для чего RibbonX Editor реально полезен

- открыть уже собранный `.xlam` и убедиться, что `customUI14.xml` внутри есть;
- быстро визуально проверить XML;
- вручную поправить XML для одноразового эксперимента;
- посмотреть структуру пакета `.xlam`.

### Как открыть файл в Office RibbonX Editor

1. Установи Office RibbonX Editor.
2. Запусти его.
3. Нажми `File -> Open`.
4. Выбери файл:
   - для локальной проверки: `build_manual_check/MainVZID.xlam`
   - для проверки установленной надстройки: `%AppData%\Microsoft\Excel\LocalCache\VZID\addin\MainVZID.xlam`
5. В левой панели найди `customUI/customUI14.xml`.
6. Открой его.

### Что там можно делать

- смотреть XML вкладки;
- менять подписи кнопок;
- добавлять/удалять кнопки;
- сохранять изменения в сам `.xlam`.

### Но правильный рабочий процесс такой

1. Сначала правишь `src/main-vzid/customui/customUI14.xml` в репозитории.
2. Потом запускаешь сборку.
3. Потом, если нужно, открываешь уже собранный `.xlam` в RibbonX Editor и проверяешь, что XML попал внутрь как надо.

То есть:

- редактирование: в репозитории;
- инспекция: в RibbonX Editor.

## Что важно про формы и кодировку

- старые VBA-экспорты модулей и форм могут быть в `cp1251`, это нормально;
- новые служебные файлы проекта лучше держать в `UTF-8`;
- `.frm` и `.frx` всегда кладутся парой;
- старые формы `frmVZID_KGN`, `frmVZID_TMN`, `frmVZID_EKB`, `frmVZID_CHLB` не импортируются в новый `MainVZID.xlam`, потому что их роль теперь выполняет вкладка ленты.

## Как правильно дорабатывать надстройку сейчас

### Новый аналог старой `Dev_MainVZID`

Раньше у тебя была одна живая `Dev_MainVZID.xlam`.

Теперь аналогом для теста будет:

- `build_manual_check/MainVZID.xlam`

Это тестовая сборка, которую ты открываешь руками в Excel и проверяешь.

### Правильный цикл работы

1. Правишь исходники в репозитории.
2. Собираешь тестовый `.xlam`.
3. Открываешь его в Excel.
4. Проверяешь ленту, макросы, формы, права.
5. Если все хорошо, поднимаешь версию.
6. Собираешь релиз.
7. Публикуешь релиз в GitHub.

## Как добавить или изменить VBA-код

### Если меняешь существующий модуль

Правь файл прямо в репозитории:

- `.bas` в `src/main-vzid/modules/`
- `.cls` в `src/main-vzid/workbook/` или `src/main-vzid/modules/`
- `.frm` и `.frx` в `src/main-vzid/forms/`

### Если переносишь код из старой надстройки

Делай так:

1. Открой старую надстройку в Excel.
2. Нажми `Alt+F11`.
3. В Project Explorer найди нужный модуль.
4. Правый клик по модулю.
5. `Export File...`
6. Сохрани:
   - модуль `.bas` в `src/main-vzid/modules/`
   - форму `.frm` и `.frx` в `src/main-vzid/forms/`
7. Пересобери `MainVZID.xlam`.

### Чего не надо делать

- не редактируй установленный `%AppData%\Microsoft\Excel\LocalCache\VZID\addin\MainVZID.xlam` как основной исходник;
- не редактируй `build/MainVZID.xlam` как основной исходник;
- не пытайся жить в логике “один рабочий xlam, из него потом все вытолкну”.

Иначе ты потеряешь изменения при следующей сборке.

## Как локально собрать тестовую версию

Команда:

```powershell
python scripts/build_addins.py --output-dir build_manual_check
```

После этого открой:

- `build_manual_check/MainVZID.xlam`

Именно этот файл используй как локальный тестовый аналог старой `Dev_MainVZID`.

## Как выпустить новую версию

### Шаг 1. Поднять номер версии

Команда:

```powershell
python scripts/set_release_version.py 0.2.3
```

Что делает этот скрипт:

- обновляет `release/version.txt`;
- обновляет минимальные версии в `release/`;
- обновляет `config/config.template.json`;
- обновляет `release/manifest.template.json`;
- обновляет `src/main-vzid/modules/MainConstants.bas`.

### Шаг 2. Собрать и проверить тестовый `.xlam`

```powershell
python scripts/build_addins.py --output-dir build_manual_check
```

Открываешь:

- `build_manual_check/MainVZID.xlam`

Если все работает, идешь дальше.

### Шаг 3. Закоммитить код и запушить `main`

Это важно:

- сначала в репозиторий должен попасть сам код;
- только потом есть смысл публиковать release-артефакты.

### Шаг 4. Опубликовать релиз

Команда:

```powershell
python scripts/publish_release.py
```

Что делает этот скрипт:

1. запускает `scripts/build_release.py`;
2. собирает:
   - `build/MainVZID.xlam`
   - `build/release/manifest.json`
   - `build/release/setup.exe`
3. через GitHub API создает или обновляет релиз `v<версия>`;
4. заливает в релиз:
   - `MainVZID.xlam`
   - `manifest.json`
   - `setup.exe`

### Важно про GitHub

Для `publish_release.py` у тебя должны работать git-учетные данные для GitHub на этой машине, потому что скрипт берет токен через:

- `git credential fill`

Если `git push origin main` у тебя работает, то обычно этого уже достаточно.

## Минимальный рабочий сценарий без лишней магии

Если совсем коротко, то теперь твой процесс такой:

1. Правишь код и XML в репозитории.
2. Запускаешь:

```powershell
python scripts/build_addins.py --output-dir build_manual_check
```

3. Открываешь:

- `build_manual_check/MainVZID.xlam`

4. Проверяешь руками.
5. Поднимаешь версию:

```powershell
python scripts/set_release_version.py 0.2.3
```

6. Коммитишь и пушишь код.
7. Публикуешь релиз:

```powershell
python scripts/publish_release.py
```

## Если хочешь совсем простой практический ответ

### Добавить кнопку

- правишь `src/main-vzid/customui/customUI14.xml`

### Назначить на нее макрос

- правишь `src/main-vzid/modules/CommandRegistry.bas`

### Назначить права

- правишь `src/main-vzid/modules/AccessPolicy.bas`
- правишь `config/config.template.json`

### Собрать тест

```powershell
python scripts/build_addins.py --output-dir build_manual_check
```

### Открыть тест

- `build_manual_check/MainVZID.xlam`

### Выпустить релиз

```powershell
python scripts/set_release_version.py 0.2.3
python scripts/publish_release.py
```
