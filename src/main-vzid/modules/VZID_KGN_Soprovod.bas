Attribute VB_Name = "VZID_KGN_Soprovod"
Option Explicit

' --- Word constants (для позднего связывания) ---
Private Const wdFormatXMLDocument As Long = 12
Private Const wdFormatOriginalFormatting As Long = 16
Private Const wdFindContinue As Long = 1
Private Const wdReplaceAll As Long = 2
Private Const wdStory As Long = 6
Private Const wdAutoFitFixed As Long = 0
Private Const wdAutoFitContent As Long = 1      ' ??? добавлено
Private Const wdAutoFitWindow As Long = 2
Private Const wdPreferredWidthPercent As Long = 2
Private Const wdLineStyleSingle As Long = 1
Private Const wdLineStyleNone As Long = 0
Private Const wdAlignParagraphLeft As Long = 0
Private Const wdAlignParagraphCenter As Long = 1
Private Const wdColorWhite As Long = 16777215
Private Const wdNoHighlight As Long = 0
Private Const wdPreferredWidthPoints As Long = 3
Private Const wdPageBreak As Long = 7
Private Const wdCollapseEnd As Long = 0
Private Const wdSectionBreakNextPage As Long = 2
Private Const wdHeaderFooterPrimary As Long = 1
Private Const wdHeaderFooterFirstPage As Long = 2
Private Const wdHeaderFooterEvenPages As Long = 3


' === Точка входа ==========================================================
Sub Make_Cover_Letters()
    Dim ws As Worksheet, wsRes As Worksheet
    Dim lastRow As Long, r As Long, blockRows As Long
    Dim templPath As String, vzyskName As String
    Dim baseFolder As String, outRoot As String
    Dim dictPerRecipient As Object
    Dim dictFilesPerRecipient As Object ' ключ -> Collection путей файлов
    Dim wdApp As Object, ownedWord As Boolean
    Dim today As Date, y As String, d As String
    Dim totalToForm As Long, formedCount As Long
    Dim representative As String
    Dim outBase As String                   ' введённый исходящий номер
    Dim addNum As String, addr As String
    Dim docSerial As Long                   ' сквозной счётчик документов
    Dim colCount As Long

    ' === Проверка наличия листа "Реестр пакетов"
    If Not SheetExists("Реестр пакетов") Then
        MsgBox "Не найден лист ""Реестр пакетов"" в текущей книге." & vbCrLf & _
               "Пожалуйста, сформируйте лист с помощью кнопки (Сформировать лист ""Реестр пакетов"" в Excel) и повторите попытку.", _
               vbExclamation, "Ошибка: отсутствует лист 'Реестр пакетов'"
        Exit Sub
    End If

    ' --- Лист
    Set ws = ActiveWorkbook.Worksheets("Реестр пакетов")
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    If lastRow < 3 Then MsgBox "На листе «Реестр пакетов» нет данных.", vbExclamation: Exit Sub

    ' --- Выбор взыскателя/шаблона (Курган)
    Dim okSel As Boolean
    okSel = KGN_SelectVzyskAndTemplate(False, vzyskName, templPath)
    If Not okSel Then Exit Sub
    If Not SafeFileExists(templPath) Then MsgBox "Не найден шаблон Word:" & vbCrLf & templPath, vbCritical: Exit Sub

    ' --- Ввод
    outBase = InputBox("Введите исходящий номер (напр., 119960):", "Исходящий номер", "")
    representative = InputBox("Введите Фамилия И.О. для поля {Представитель}:", "Представитель", "")
    addNum = InputBox("Введите добавочный номер (если есть). Будет подставлен в {доб}.", "Добавочный номер", "")
    docSerial = 1

    ' --- Предскан
    Dim missRec As Long, missInfo As Long, rowsMissRec As Collection, rowsMissInfo As Collection
    Dim resp As VbMsgBoxResult, totalBlocksAll As Long, msg As String
    totalBlocksAll = CountBlocks(ws, False)          ' < исправлено
    PreScanMissing ws, False, missRec, missInfo, rowsMissRec, rowsMissInfo
    If (missRec > 0) Or (missInfo > 0) Then
        If missRec > 0 Then msg = msg & "По " & missRec & " из " & totalBlocksAll & " документов отсутствует Получатель" & vbCrLf
        If missInfo > 0 Then msg = msg & "По " & missInfo & " из " & totalBlocksAll & " документов отсутствует Инфо" & vbCrLf
        msg = msg & vbCrLf & "Да – сформировать по тем, где все данные есть" & vbCrLf & "Нет – не формировать ничего"
        resp = MsgBox(msg, vbQuestion + vbYesNo, "Проверка заполненности")
        If resp = vbNo Then HighlightMissing ws, rowsMissRec, rowsMissInfo: Exit Sub
    End If

    ' --- Папка (уникальное имя)
    baseFolder = PickFolder("Куда сохранить документы?")
    If Len(baseFolder) = 0 Then Exit Sub
    today = Date: y = CStr(Year(today)): d = Format$(today, "dd.mm.yyyy")
    outRoot = KGN_UniqueSubfolderPath(baseFolder, SanitizeFileName(d & " " & vzyskName))
    EnsureFolder outRoot

    ' --- Лист результатов
    Set wsRes = CopyAsResultSheet(ws): PrepareResultHeaders wsRes

    ' --- Word
    Set dictPerRecipient = CreateObject("Scripting.Dictionary")
    Set dictFilesPerRecipient = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set wdApp = GetObject(Class:="Word.Application")
    ownedWord = False
    If wdApp Is Nothing Then Set wdApp = CreateObject("Word.Application"): ownedWord = True
    On Error GoTo 0
    If wdApp Is Nothing Then MsgBox "Не удалось запустить Word.", vbCritical: Exit Sub
    wdApp.Visible = False
    Application.ScreenUpdating = False

    totalToForm = CountBlocksReady(ws, False)
    formedCount = 0
    Application.StatusBar = "Формирование… 0/" & totalToForm


    ' --- Основной цикл
    r = 3
    Do While r <= lastRow
        Dim cellA As Range, recipient As String
        Dim rngInfo As Range, infoText As String
        Dim okToMake As Boolean, reason As String
        Dim doc As Object, outName As String, outPath As String, outFolder As String

        Set cellA = ws.Cells(r, "A")
        blockRows = 1
        If cellA.MergeCells Then
            If cellA.Address <> cellA.MergeArea.Cells(1, 1).Address Then r = r + 1: GoTo NextR
            blockRows = cellA.MergeArea.rows.count
        End If

        recipient = Trim(GetMergedTopValue(ws.Cells(r, "C")))
        addr = Trim(GetMergedTopValue(ws.Cells(r, "B")))
        Set rngInfo = ws.Cells(r, "D").MergeArea

        okToMake = True: reason = ""
        If rngInfo Is Nothing Or LenB(CStr(rngInfo.Cells(1, 1).value)) = 0 Then okToMake = False: reason = reason & "пустой Инфо; "
        If Len(recipient) = 0 Then okToMake = False: reason = reason & "пустой Получатель; "

        If okToMake Then
            Dim idx As Long
            If Not dictPerRecipient.exists(recipient) Then dictPerRecipient.Add recipient, 0
            idx = dictPerRecipient(recipient) + 1
            dictPerRecipient(recipient) = idx

            outFolder = AppendPath(outRoot, SanitizeFileName(recipient))
            EnsureFolder outFolder
            outName = SanitizeFileName(vzyskName & " - " & recipient & " - " & CStr(idx) & ".docx")
            outPath = AppendPath(outFolder, outName)

            ' записываем путь в dictFilesPerRecipient
            Dim coll As Collection
            If dictFilesPerRecipient.exists(recipient) Then
                Set coll = dictFilesPerRecipient(recipient)
            Else
                Set coll = New Collection
                dictFilesPerRecipient.Add recipient, coll
            End If
            coll.Add outPath

            Dim tplPathResolved As String
            tplPathResolved = ResolveTemplateForWord(templPath)
            Set doc = wdApp.Documents.Add(Template:=tplPathResolved)


            ' Плейсхолдеры
            WordReplaceEverywhere doc, "{год}", y
            WordReplaceEverywhere doc, "{дата}", d
            WordReplaceEverywhere doc, "{Получатель}", recipient
            WordReplaceEverywhere doc, "{Представитель}", representative

            ' {исх2} — база; {исх} — база + сквозной счётчик
            WordReplaceEverywhere doc, "{исх2}", outBase
            If Len(outBase) > 0 Then
                WordReplaceEverywhere doc, "{исх}", outBase & "/" & CStr(docSerial)
            Else
                WordReplaceEverywhere doc, "{исх}", ""
            End If

            WordReplaceEverywhere doc, "{доб}", addNum

            ' {Адрес}
            If Len(addr) > 0 Then WordReplaceEverywhere doc, "{Адрес}", addr _
                              Else WordReplaceEverywhere doc, "{Адрес}", ""

            ' {кол} — кол-во строк в объединённой ячейке блока
            colCount = blockRows
            WordReplaceEverywhere doc, "{кол}", CStr(colCount)

            ' {Инфо}
            infoText = CStr(rngInfo.Cells(1, 1).value)
            If Not ReplacePlaceholderWithPaste(wdApp, doc, "{Инфо}") Then doc.Activate
            InsertInfoSingleCell wdApp, infoText, doc

            doc.SaveAs2 fileName:=outPath, FileFormat:=wdFormatXMLDocument
            doc.Close SaveChanges:=False
            Set doc = Nothing

            formedCount = formedCount + 1
            docSerial = docSerial + 1                 ' < сквозная нумерация
            Application.StatusBar = "Формирование… " & formedCount & "/" & totalToForm
            WriteResult wsRes, r, blockRows, "Сформировано", outName
        Else
            WriteResult wsRes, r, blockRows, "Не сформировано", "Ошибка: " & reason
        End If

        r = r + blockRows
NextR:
    Loop

    ' --- Создание общих файлов по папкам (по каждому получателю)
    Dim k As Variant
    For Each k In dictFilesPerRecipient.keys
        Dim filesColl As Collection
        Set filesColl = dictFilesPerRecipient(k)
        If filesColl.count > 0 Then
            Dim firstFile As String, combinedPath As String
            firstFile = filesColl(1)
            combinedPath = AppendPath(AppendPath(outRoot, SanitizeFileName(k)), SanitizeFileName(vzyskName & " - " & k & " - Общий.docx"))
            CreateCombinedDocFromList wdApp, filesColl, combinedPath
        End If
    Next k

    Application.StatusBar = "Формирование завершено: " & formedCount & "/" & totalToForm
    MsgBox "Готово. Сформировано: " & formedCount & " из " & totalToForm & "." & vbCrLf & _
           "Папка: " & outRoot & vbCrLf & "Результаты помечены на листе: " & wsRes.name, vbInformation

    ' После MsgBox открываем папку с документами
    shell "explorer.exe """ & outRoot & """", vbNormalFocus

CLEANUP:
    On Error Resume Next
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    If ownedWord And Not wdApp Is Nothing Then wdApp.Quit
    Application.StatusBar = False
    Set wdApp = Nothing
End Sub

' === Предскан: отсутствие Получатель/Инфо ================================
Private Sub PreScanMissing(ByVal ws As Worksheet, ByVal only3plus As Boolean, _
                           ByRef missRec As Long, ByRef missInfo As Long, _
                           ByRef rowsMissRec As Collection, ByRef rowsMissInfo As Collection)
    Dim r As Long, lastRow As Long, cellA As Range, blockRows As Long
    Dim recipient As String, rngInfo As Range
    Set rowsMissRec = New Collection
    Set rowsMissInfo = New Collection
    missRec = 0: missInfo = 0

    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    r = 3
    Do While r <= lastRow
        Set cellA = ws.Cells(r, "A")
        blockRows = 1
        If cellA.MergeCells Then
            If cellA.Address <> cellA.MergeArea.Cells(1, 1).Address Then
                r = r + 1
                GoTo NextR
            End If
            blockRows = cellA.MergeArea.rows.count
        End If

        If (Not only3plus) Or (blockRows >= 3) Then
            recipient = Trim(GetMergedTopValue(ws.Cells(r, "C")))
            Set rngInfo = ws.Cells(r, "D").MergeArea
            If Len(recipient) = 0 Then
                missRec = missRec + 1
                rowsMissRec.Add r
            End If
            If (rngInfo Is Nothing) Or LenB(CStr(rngInfo.Cells(1, 1).value)) = 0 Then
                missInfo = missInfo + 1
                rowsMissInfo.Add r
            End If
        End If

        r = r + blockRows
NextR:
    Loop
End Sub

Private Sub HighlightMissing(ByVal ws As Worksheet, ByVal rowsMissRec As Collection, ByVal rowsMissInfo As Collection)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long, rTop As Variant, blockRows As Long
    For i = 1 To rowsMissRec.count: dict(rowsMissRec(i)) = True: Next
    For i = 1 To rowsMissInfo.count: dict(rowsMissInfo(i)) = True: Next
    For Each rTop In dict.keys
        If ws.Cells(CLng(rTop), "A").MergeCells Then
            blockRows = ws.Cells(CLng(rTop), "A").MergeArea.rows.count
        Else
            blockRows = 1
        End If
        With ws.Range(ws.Cells(CLng(rTop), "A"), ws.Cells(CLng(rTop) + blockRows - 1, "D")).Font
            .Color = vbRed
        End With
    Next rTop
End Sub

' === Подсчёт пакетов =====================================================
Private Function CountBlocks(ByVal ws As Worksheet, ByVal only3plus As Boolean) As Long
    Dim r As Long, lastRow As Long, cellA As Range, blockRows As Long, total As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    r = 3: total = 0
    Do While r <= lastRow
        Set cellA = ws.Cells(r, "A")
        blockRows = 1
        If cellA.MergeCells Then
            If cellA.Address <> cellA.MergeArea.Cells(1, 1).Address Then
                r = r + 1
                GoTo NextR
            End If
            blockRows = cellA.MergeArea.rows.count
        End If
        If (Not only3plus) Or (blockRows >= 3) Then total = total + 1
        r = r + blockRows
NextR:
    Loop
    CountBlocks = total
End Function

Private Function CountBlocksReady(ByVal ws As Worksheet, ByVal only3plus As Boolean) As Long
    Dim r As Long, lastRow As Long, cellA As Range, blockRows As Long, total As Long
    Dim recipient As String, rngInfo As Range
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    r = 3: total = 0
    Do While r <= lastRow
        Set cellA = ws.Cells(r, "A")
        blockRows = 1
        If cellA.MergeCells Then
            If cellA.Address <> cellA.MergeArea.Cells(1, 1).Address Then
                r = r + 1
                GoTo NextR
            End If
            blockRows = cellA.MergeArea.rows.count
        End If

        If (Not only3plus) Or (blockRows >= 3) Then
            recipient = Trim(GetMergedTopValue(ws.Cells(r, "C")))
            Set rngInfo = ws.Cells(r, "D").MergeArea
            If Len(recipient) > 0 _
               And (Not (rngInfo Is Nothing)) _
               And LenB(CStr(rngInfo.Cells(1, 1).value)) > 0 Then
                total = total + 1
            End If
        End If

        r = r + blockRows
NextR:
    Loop
    CountBlocksReady = total
End Function

' === Копирование листа -> «Результат формирования[_n]» ===================
Private Function CopyAsResultSheet(ByVal ws As Worksheet) As Worksheet
    Dim nm As String
    nm = UniqueSheetName("Результат формирования")
    ws.Copy After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.count)
    Set CopyAsResultSheet = ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.count)
    On Error Resume Next
    CopyAsResultSheet.name = nm
    On Error GoTo 0
End Function

Private Sub PrepareResultHeaders(ByVal wsRes As Worksheet)
    wsRes.Range("E2").value = "Результат"
    wsRes.Range("F2").value = "Описание"
    wsRes.Range("E2:F2").HorizontalAlignment = xlCenter
    wsRes.Range("E2:F2").VerticalAlignment = xlCenter
    wsRes.Columns("E").ColumnWidth = 16
    wsRes.Columns("F").ColumnWidth = 60
End Sub

' === Запись результата ===================================================
Private Sub WriteResult(ByVal wsRes As Worksheet, ByVal topRow As Long, ByVal blockRows As Long, _
                        ByVal resultText As String, ByVal descr As String)
    Dim rngE As Range, rngF As Range
    Set rngE = wsRes.Range(wsRes.Cells(topRow, "E"), wsRes.Cells(topRow + blockRows - 1, "E"))
    Set rngF = wsRes.Range(wsRes.Cells(topRow, "F"), wsRes.Cells(topRow + blockRows - 1, "F"))
    On Error Resume Next
    If rngE.MergeCells Then rngE.UnMerge
    If rngF.MergeCells Then rngF.UnMerge
    On Error GoTo 0
    rngE.Merge: rngF.Merge
    With rngE
        .value = resultText
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With rngF
        .value = descr
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With wsRes.Range(wsRes.Cells(topRow, "E"), wsRes.Cells(topRow + blockRows - 1, "F")).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub

' === Утилиты Excel/FS ===================================================
Private Function PickFolder(Optional ByVal title As String = "Выбор папки") As String
    Dim fd As FileDialog
    On Error GoTo FAIL
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .title = title
        .AllowMultiSelect = False
        If .Show = -1 Then PickFolder = .SelectedItems(1)
    End With
    Exit Function
FAIL:
    PickFolder = ""
End Function

Private Function AppendPath(base As String, tail As String) As String
    If Right$(base, 1) = "\" Or Right$(base, 1) = "/" Then
        AppendPath = base & tail
    Else
        AppendPath = base & "\" & tail
    End If
End Function

Private Sub EnsureFolder(ByVal path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(path) Then fso.CreateFolder path
End Sub

Private Function SanitizeFileName(ByVal s As String) As String
    Dim badChars As Variant, ch As Variant
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In badChars
        s = Replace$(s, CStr(ch), "_")
    Next ch
    Do While Len(s) > 0 And (Right$(s, 1) = " " Or Right$(s, 1) = ".")
        s = Left$(s, Len(s) - 1)
    Loop
    If Len(s) = 0 Then s = "_"
    SanitizeFileName = s
End Function

Private Function GetMergedTopValue(ByVal anyCell As Range) As String
    If anyCell.MergeCells Then
        GetMergedTopValue = CStr(anyCell.MergeArea.Cells(1, 1).value)
    Else
        GetMergedTopValue = CStr(anyCell.value)
    End If
End Function

' === Word-хелперы =======================================================
Private Sub WordReplaceAll(ByVal oDoc As Object, ByVal findText As String, ByVal replText As String)
    With oDoc.Content.Find
        .ClearFormatting: .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replText
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
End Sub

' === Замена плейсхолдера во ВСЕХ частях документа (тело, колонтитулы, фреймы) ===
Private Sub WordReplaceEverywhere(ByVal oDoc As Object, ByVal findText As String, ByVal replText As String)
    Dim rngStory As Object
    Dim rng As Object

    For Each rngStory In oDoc.StoryRanges
        Set rng = rngStory
        Do While Not rng Is Nothing
            With rng.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = findText
                .Replacement.Text = replText
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
            Set rng = rng.NextStoryRange
        Loop
    Next rngStory
End Sub

Private Function ReplacePlaceholderWithPaste(ByVal wdApp As Object, ByVal doc As Object, ByVal placeholder As String) As Boolean
    doc.Activate
    With wdApp.Selection
        .HomeKey wdStory
        With .Find
            .ClearFormatting: .Replacement.ClearFormatting
            .Text = placeholder
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            If .Execute Then
                Dim rng As Object
                Set rng = wdApp.Selection.Range
                rng.Text = ""
                wdApp.Selection.SetRange Start:=rng.End, End:=rng.End
                ReplacePlaceholderWithPaste = True
                Exit Function
            End If
        End With
    End With
    ReplacePlaceholderWithPaste = False
End Function

' === {Инфо}: одна ячейка 17,5 см, просто текст, TNR 12 пт =================
Private Sub InsertInfoSingleCell(ByVal wdApp As Object, ByVal infoText As String, ByVal doc As Object)
    On Error Resume Next

    ' Нормализуем переносы строк
    Dim t As String
    t = Replace(infoText, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    t = Replace(t, vbLf, vbCr) ' Word-стиль

    ' Создаём таблицу 1x1
    Dim tbl As Object
    Set tbl = doc.Tables.Add(wdApp.Selection.Range, 1, 1)

    ' Фиксируем общую ширину 17,5 см
    Dim targetW As Single: targetW = wdApp.CentimetersToPoints(17.5!)
    tbl.AllowAutoFit = False
    tbl.AutoFitBehavior 0             ' wdAutoFitFixed
    tbl.PreferredWidthType = wdPreferredWidthPoints
    tbl.PreferredWidth = targetW
    tbl.Columns(1).width = targetW
    tbl.rows.LeftIndent = 0
    tbl.LeftPadding = 0: tbl.RightPadding = 0

    ' Рамка как на скрине (при необходимости убери)
    With tbl.Borders
        .OutsideLineStyle = 1         ' wdLineStyleSingle
        .InsideLineStyle = 0
    End With

    ' Вставляем текст
    With tbl.cell(1, 1).Range
        .Text = t
        .Font.name = "Times New Roman"
        .Font.Size = 12
        With .ParagraphFormat
            .Alignment = 0            ' wdAlignParagraphLeft
            .LeftIndent = 0
            .RightIndent = 0
            .SpaceBefore = 0
            .SpaceAfter = 0
            .LineSpacingRule = 0      ' одинарный
        End With
    End With
End Sub

' ширина полосы набора (points)
Private Function GetContentWidthPoints(ByVal doc As Object) As Single
    Dim w As Single
    w = doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin
    If w <= 0 Then w = 468 ' ~16.5 см
    GetContentWidthPoints = w
End Function

' === Создать общий документ из списка файлов (с сохранением форматирования и колонтитулов) ===
Private Sub CreateCombinedDocFromList(ByVal wdApp As Object, ByVal filesColl As Collection, ByVal destPath As String)
    On Error GoTo ErrHandler
    If filesColl.count <= 1 Then
        ' По требованию: если в папке всего 0 или 1 файл — общий файл НЕ создаём.
        Exit Sub
    End If

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim i As Long

    ' Открыть первый файл как базовый комбинируемый документ (ReadOnly:=False)
    Dim baseDoc As Object
    Set baseDoc = wdApp.Documents.Open(fileName:=filesColl(1), ReadOnly:=False, AddToRecentFiles:=False, Visible:=False)

    ' Для каждого следующего файла: открыть, вставить секционный разрыв, вставить содержимое через FormattedText,
    ' скопировать заголовки/футеры (и снять LinkToPrevious), затем закрыть источник
    Dim srcDoc As Object
    Dim lastSection As Object
    For i = 2 To filesColl.count
        Set srcDoc = wdApp.Documents.Open(fileName:=filesColl(i), ReadOnly:=True, AddToRecentFiles:=False, Visible:=False)

        ' Переместить селекцию в конец baseDoc
        baseDoc.Activate
        With wdApp.Selection
            .EndKey Unit:=wdStory ' wdStory = 6, переместиться в конец документа
        End With

        ' Вставить секционный разрыв "Следующая страница"
        wdApp.Selection.InsertBreak Type:=wdSectionBreakNextPage

        ' После разрыва селекция находится в новой секции: вставляем содержимое с сохранением форматирования
        wdApp.Selection.Range.FormattedText = srcDoc.Content.FormattedText

        ' Скопировать заголовки/футеры из srcDoc.Sections(1) в последнюю секцию baseDoc
        Set lastSection = baseDoc.Sections(baseDoc.Sections.count)
        On Error Resume Next
        Dim hfType As Variant
        For Each hfType In Array(wdHeaderFooterPrimary, wdHeaderFooterFirstPage, wdHeaderFooterEvenPages)
            ' Снять связь с предыдущей секцией
            lastSection.headers(hfType).LinkToPrevious = False
            lastSection.Footers(hfType).LinkToPrevious = False

            ' Скопировать, если есть содержимое в srcDoc
            If srcDoc.Sections.count >= 1 Then
                On Error Resume Next
                lastSection.headers(hfType).Range.FormattedText = srcDoc.Sections(1).headers(hfType).Range.FormattedText
                lastSection.Footers(hfType).Range.FormattedText = srcDoc.Sections(1).Footers(hfType).Range.FormattedText
                On Error GoTo 0
            End If
        Next
        On Error GoTo 0

        srcDoc.Close SaveChanges:=False
        Set srcDoc = Nothing
    Next i

    ' Сохраняем комбинированный файл
    baseDoc.SaveAs2 fileName:=destPath, FileFormat:=wdFormatXMLDocument
    baseDoc.Close SaveChanges:=False
    Set baseDoc = Nothing
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not baseDoc Is Nothing Then baseDoc.Close SaveChanges:=False
    If Not srcDoc Is Nothing Then srcDoc.Close SaveChanges:=False
End Sub

' === Разное ==============================================================

Private Function UniqueSheetName(baseName As String) As String
    Dim nm As String, i As Long
    nm = baseName
    i = 2
    Do While SheetExists(nm)
        nm = baseName & "_" & i
        i = i + 1
    Loop
    UniqueSheetName = nm
End Function

Private Function SheetExists(nm As String) As Boolean
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ActiveWorkbook.Worksheets(nm)
    SheetExists = Not sh Is Nothing
    On Error GoTo 0
End Function

Private Function SafeFileExists(ByVal p As String) As Boolean
    On Error Resume Next
    SafeFileExists = CreateObject("Scripting.FileSystemObject").FileExists(p)
End Function

' === ResolveTemplateForWord - копировать шаблон в %AppData%\Microsoft\Templates один раз за запуск ===
Private Function ResolveTemplateForWord(ByVal netPath As String) As String
    ' Возвращает путь к локальной копии шаблона в %AppData%\Microsoft\Templates
    ' Копирует шаблон каждый раз при запуске, перезаписывая старую копию
    On Error GoTo ErrHandler

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim appData As String
    appData = Environ$("AppData")
    If Right$(appData, 1) <> "\" Then appData = appData & "\"
    Dim templatesFolder As String
    templatesFolder = appData & "Microsoft\Templates\"
    
    ' Убедимся, что папка существует
    If Not fso.FolderExists(templatesFolder) Then fso.CreateFolder templatesFolder

    ' Проверка исходного файла
    If Len(Trim(netPath)) = 0 Or Not fso.FileExists(netPath) Then
        ResolveTemplateForWord = netPath
        Exit Function
    End If

    ' Локальный путь шаблона
    Dim localName As String
    localName = fso.GetFileName(netPath) ' сохраняем исходное имя
    Dim localPath As String
    localPath = templatesFolder & localName

    ' Копируем шаблон из сети, перезаписывая старую локальную копию
    fso.CopyFile netPath, localPath, True

    ResolveTemplateForWord = localPath
    Exit Function

ErrHandler:
    ' Если что-то пошло не так — возвращаем оригинальный путь
    On Error Resume Next
    ResolveTemplateForWord = netPath
End Function

' === Опциональная утилита: удалить старые локальные копии шаблонов ===
' Удаляет файлы, содержащие baseName и старше daysOld дней в %AppData%\Microsoft\Templates
Private Sub CleanOldLocalTemplates(ByVal baseName As String, ByVal daysOld As Long)
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim appData As String: appData = Environ$("AppData")
    If Right$(appData, 1) <> "\" Then appData = appData & "\"
    Dim templatesFolder As String: templatesFolder = appData & "Microsoft\Templates\"
    If Not fso.FolderExists(templatesFolder) Then Exit Sub

    Dim fld As Object: Set fld = fso.GetFolder(templatesFolder)
    Dim f As Object
    Dim cutoff As Date: cutoff = DateAdd("d", -CLng(daysOld), Date)
    For Each f In fld.files
        If InStr(1, f.name, baseName, vbTextCompare) > 0 Then
            If f.DateLastModified < cutoff Then
                On Error Resume Next
                f.Delete True
                On Error Resume Next
            End If
        End If
    Next
End Sub

' === Дополнительная версия для формирования в режиме WSS (по каждому получателю) ===
Sub Make_Cover_Letters_ByRecipient()
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long, blockRows As Long
    Dim recipient As String, infoText As String, addr As String
    Dim templPath As String, vzyskName As String
    Dim representative As String
    Dim baseFolder As String, outRoot As String, outPath As String
    Dim wdApp As Object, doc As Object, ownedWord As Boolean
    Dim today As Date, y As String, d As String
    Dim dictFirstRow As Object, order As Collection
    Dim k As Variant, total As Long, done As Long
    Dim addNum As String
    Dim colCount As Long

    ' === Проверка наличия листа "Реестр пакетов"
    If Not SheetExists("Реестр пакетов") Then
        MsgBox "Не найден лист ""Реестр пакетов"" в текущей книге." & vbCrLf & _
               "Пожалуйста, убедитесь, что в книге есть лист с именем ""Реестр пакетов"" и повторите попытку.", _
               vbExclamation, "Ошибка: отсутствует лист 'Реестр пакетов'"
        Exit Sub
    End If

    ' Лист
    Set ws = ActiveWorkbook.Worksheets("Реестр пакетов")
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    If lastRow < 3 Then MsgBox "Нет данных на листе «Реестр пакетов».", vbExclamation: Exit Sub

    ' Первые строки по каждому получателю
    Set dictFirstRow = CreateObject("Scripting.Dictionary")
    Set order = New Collection
    r = 3
    Do While r <= lastRow
        blockRows = 1
        If ws.Cells(r, "A").MergeCells Then
            If ws.Cells(r, "A").Address <> ws.Cells(r, "A").MergeArea.Cells(1, 1).Address Then r = r + 1: GoTo NextR
            blockRows = ws.Cells(r, "A").MergeArea.rows.count
        End If
        recipient = Trim(GetMergedTopValue(ws.Cells(r, "C")))
        If Len(recipient) > 0 Then
            If Not dictFirstRow.exists(recipient) Then
                dictFirstRow.Add recipient, r
                order.Add recipient
            End If
        End If
        r = r + blockRows
NextR:
    Loop
    If order.count = 0 Then MsgBox "Не найдено ни одного получателя.", vbExclamation: Exit Sub

    ' Выбор WSS-шаблона
    Dim okSel As Boolean
    okSel = KGN_SelectVzyskAndTemplate(True, vzyskName, templPath)
    If Not okSel Then Exit Sub
    If Not SafeFileExists(templPath) Then MsgBox "Не найден шаблон Word:" & vbCrLf & templPath, vbCritical: Exit Sub

    ' Остальные вводы
    representative = InputBox("Введите Фамилия И.О. для поля {Представитель}:", "Представитель", "")
    addNum = InputBox("Введите добавочный номер (если есть). Будет подставлен в {доб}.", "Добавочный номер", "")

    baseFolder = PickFolder("Куда сохранить документы?")
    If Len(baseFolder) = 0 Then Exit Sub
    today = Date: y = CStr(Year(today)): d = Format$(today, "dd.mm.yyyy")
    outRoot = KGN_UniqueSubfolderPath(baseFolder, SanitizeFileName(d & " " & vzyskName & " для загрузки в WSS"))
    EnsureFolder outRoot

    ' Word
    On Error Resume Next
    Set wdApp = GetObject(Class:="Word.Application")
    If wdApp Is Nothing Then Set wdApp = CreateObject("Word.Application"): ownedWord = True
    On Error GoTo 0
    If wdApp Is Nothing Then MsgBox "Не удалось запустить Word.", vbCritical: Exit Sub
    wdApp.Visible = False

    ' Формирование
    total = order.count: done = 0
    Application.StatusBar = "Формирование… 0/" & total

    Dim rTop As Long

    For Each k In order
        rTop = CLng(dictFirstRow(CStr(k)))
        recipient = CStr(k)
        addr = Trim(GetMergedTopValue(ws.Cells(rTop, "B")))
        infoText = CStr(ws.Cells(rTop, "D").MergeArea.Cells(1, 1).value)
        If LenB(infoText) = 0 Then GoTo SkipThis

        Dim tplPathResolved As String
        tplPathResolved = ResolveTemplateForWord(templPath) ' сетевой путь

        Set doc = wdApp.Documents.Add(Template:=tplPathResolved)

        WordReplaceEverywhere doc, "{год}", y
        WordReplaceEverywhere doc, "{дата}", d
        WordReplaceEverywhere doc, "{Получатель}", recipient
        WordReplaceEverywhere doc, "{Представитель}", representative

        ' В этом режиме {исх}/{исх2} не используются — стираем
        WordReplaceEverywhere doc, "{исх}", ""
        WordReplaceEverywhere doc, "{исх2}", ""

        WordReplaceEverywhere doc, "{доб}", addNum

        ' {Адрес}
        If Len(addr) > 0 Then WordReplaceEverywhere doc, "{Адрес}", addr _
                          Else WordReplaceEverywhere doc, "{Адрес}", ""

        ' {кол} — размер MergeArea столбца D у этой первой строки
        If ws.Cells(rTop, "D").MergeCells Then
            colCount = ws.Cells(rTop, "D").MergeArea.rows.count
        Else
            colCount = 1
        End If
        WordReplaceEverywhere doc, "{кол}", CStr(colCount)

        ' {Инфо}
        If Not ReplacePlaceholderWithPaste(wdApp, doc, "{Инфо}") Then doc.Range(doc.Content.End - 1, doc.Content.End - 1).Select
        InsertInfoSingleCell wdApp, infoText, doc

        outPath = AppendPath(outRoot, SanitizeFileName(vzyskName & " - " & recipient & ".docx"))
        doc.SaveAs2 fileName:=outPath, FileFormat:=wdFormatXMLDocument
        doc.Close SaveChanges:=False
        Set doc = Nothing

        done = done + 1
        Application.StatusBar = "Формирование… " & done & "/" & total
SkipThis:
    Next k

    ' ВНИМАНИЕ: по вашему запросу — общий файл НЕ создаётся в этом режиме

    If ownedWord Then wdApp.Quit
    Application.StatusBar = False
    MsgBox "Готово. Сформировано документов: " & done & " из " & total & vbCrLf & outRoot, vbInformation
    ' После MsgBox открываем папку с документами
    shell "explorer.exe """ & outRoot & """", vbNormalFocus

End Sub

' === Выбор взыскателя/шаблона (Курган) ======================================
Private Function KGN_SelectVzyskAndTemplate(ByVal isWSS As Boolean, _
                                            ByRef vzyskName As String, _
                                            ByRef templPath As String) As Boolean
    Dim root As String
    Dim names() As String, filesBase() As String, filesWSS() As String
    Dim q As Variant, idx As Long

    ' Папка Кургана
    root = "\\corp.vostok-electra.ru\Kgn\Отделы\Отдел взыскания по исполнительным документам\Зуйкевич Данил Иванович\Excel\ВЗИД\"

    ' Взыскатели
    names = Split("АО ЭК Восток|Водный союз|КГК|СКС|Чистый город", "|")

    ' Имена файлов шаблонов (обычные)
    ReDim filesBase(0 To 4)
    filesBase(0) = "Сопровод Восток.docx"
    filesBase(1) = "Сопровод ВС.docx"
    filesBase(2) = "Сопровод КГК.docx"
    filesBase(3) = "Сопровод СКС.docx"
    filesBase(4) = "Сопровод ЧГ.docx"

    ' Имена файлов шаблонов (WSS)
    ReDim filesWSS(0 To 4)
    filesWSS(0) = "Сопровод Восток для WSS.docx"
    filesWSS(1) = "Сопровод ВС для WSS.docx"
    filesWSS(2) = "Сопровод КГК для WSS.docx"
    filesWSS(3) = "Сопровод СКС для WSS.docx"
    filesWSS(4) = "Сопровод ЧГ для WSS.docx"

    ' Выбор взыскателя
    q = Application.InputBox( _
        prompt:="По какому взыскателю формировать?" & vbCrLf & _
                "1 – " & names(0) & vbCrLf & _
                "2 – " & names(1) & vbCrLf & _
                "3 – " & names(2) & vbCrLf & _
                "4 – " & names(3) & vbCrLf & _
                "5 – " & names(4), _
        title:="Выбор взыскателя (Курган)", Type:=1)
    If VarType(q) = vbBoolean And q = False Then Exit Function

    idx = CLng(q) - 1
    If idx < 0 Or idx > 4 Then
        MsgBox "Введите число 1..5.", vbExclamation
        Exit Function
    End If

    vzyskName = names(idx)
    If isWSS Then
        templPath = root & filesWSS(idx)
    Else
        templPath = root & filesBase(idx)
    End If

    KGN_SelectVzyskAndTemplate = True
End Function

' Возвращает путь к подпапке folderName внутри baseFolder.
' Если уже существует, добавляет " - 1", " - 2", ...
Private Function KGN_UniqueSubfolderPath(ByVal baseFolder As String, ByVal folderName As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim basePath As String, n As Long, p As String
    If Right$(baseFolder, 1) = "\" Or Right$(baseFolder, 1) = "/" Then
        basePath = baseFolder
    Else
        basePath = baseFolder & "\"
    End If
    p = basePath & folderName
    If Not fso.FolderExists(p) Then KGN_UniqueSubfolderPath = p: Exit Function
    n = 1
    Do
        p = basePath & folderName & " - " & n
        If Not fso.FolderExists(p) Then KGN_UniqueSubfolderPath = p: Exit Function
        n = n + 1
    Loop
End Function

Private Sub LogTplDiag(ByVal p As String)
    On Error Resume Next
    Application.StatusBar = "Template=" & p & " | Len=" & Len(p) & _
                            " | Temp=" & Environ$("TEMP")
End Sub


