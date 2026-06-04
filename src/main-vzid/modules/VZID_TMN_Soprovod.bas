Attribute VB_Name = "VZID_TMN_Soprovod"
Option Explicit

' --- Word constants (для позднего связывания) ---
Private Const wdFormatXMLDocument As Long = 12
Private Const wdFindContinue As Long = 1
Private Const wdReplaceAll As Long = 2
Private Const wdStory As Long = 6
Private Const wdAutoFitFixed As Long = 0
Private Const wdPreferredWidthPoints As Long = 3
Private Const wdLineStyleSingle As Long = 1
Private Const wdLineStyleNone As Long = 0
Private Const wdAlignParagraphLeft As Long = 0
Private Const wdSectionBreakNextPage As Long = 2
Private Const wdHeaderFooterPrimary As Long = 1
Private Const wdHeaderFooterFirstPage As Long = 2
Private Const wdHeaderFooterEvenPages As Long = 3

' === ВХОД №1: пакетно по всем блокам (Тюмень) ==============================
Public Sub TMN_Make_Cover_Letters()
    Dim ws As Worksheet, wsRes As Worksheet
    Dim lastRow As Long, r As Long, blockRows As Long
    Dim templPath As String, vzyskName As String
    Dim tplResolved As String
    Dim baseFolder As String, outRoot As String
    Dim dictPerRecipient As Object
    Dim dictFilesPerRecipient As Object ' ключ: Получатель -> Collection путей файлов
    Dim wdApp As Object, ownedWord As Boolean
    Dim today As Date, y As String, d As String
    Dim totalToForm As Long, formedCount As Long
    Dim representative As String
    Dim addNum As String
    Dim addr As String

    ' === Проверка наличия листа "Реестр пакетов"
    If Not TMN_SheetExists("Реестр пакетов") Then
        MsgBox "Не найден лист ""Реестр пакетов"" в текущей книге." & vbCrLf & _
               "Пожалуйста, сформируйте лист с помощью кнопки Сформировать ""Реестр пакетов"" в Excel и повторите попытку.", _
               vbExclamation, "Ошибка: отсутствует лист 'Реестр пакетов'"
        Exit Sub
    End If

    Set ws = ActiveWorkbook.Worksheets("Реестр пакетов")
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    If lastRow < 3 Then MsgBox "На листе «Реестр пакетов» нет данных.", vbExclamation: Exit Sub

    If Not TMN_SelectTyumenVzysk(False, vzyskName, templPath) Then Exit Sub
    If Not TMN_SafeFileExists(templPath) Then MsgBox "Не найден шаблон Word:" & vbCrLf & templPath, vbCritical: Exit Sub

    ' Копируем шаблон в %AppData%\Microsoft\Templates и используем локальную копию (обновляется каждый запуск)
    tplResolved = TMN_ResolveTemplateToUserTemplates(templPath)

    Dim missRec As Long, missInfo As Long, rowsMissRec As Collection, rowsMissInfo As Collection
    Dim msg As String, resp As VbMsgBoxResult, totalBlocksAll As Long
    totalBlocksAll = TMN_CountBlocks(ws)
    TMN_PreScanMissing ws, missRec, missInfo, rowsMissRec, rowsMissInfo
    If (missRec > 0) Or (missInfo > 0) Then
        If missRec > 0 Then msg = msg & "По " & missRec & " из " & totalBlocksAll & " документам отсутствует Получатель" & vbCrLf
        If missInfo > 0 Then msg = msg & "По " & missInfo & " из " & totalBlocksAll & " документам отсутствует Инфо" & vbCrLf
        msg = msg & vbCrLf & "Да – сформировать по тем, где все данные есть" & vbCrLf & "Нет – не формировать ничего"
        resp = MsgBox(msg, vbQuestion + vbYesNo, "Проверка заполненности")
        If resp = vbNo Then
            TMN_HighlightMissing ws, rowsMissRec, rowsMissInfo
            Exit Sub
        End If
    End If

    representative = InputBox("Ваше Фамилия И.О. для поля {Представитель} подставлено автоматически. Если требуется другое, то введите его вручную:", "Представитель", "")
    addNum = InputBox("Ваш добавочный номер для поля {доб} подставлен автоматически. Если трубуется другой, то введите его вручную:", "Добавочный номер", "")

    baseFolder = TMN_PickFolder("Куда сохранить документы?")
    If Len(baseFolder) = 0 Then Exit Sub
    today = Date: y = CStr(Year(today)): d = Format$(today, "dd.mm.yyyy")
    outRoot = TMN_UniqueSubfolderPath(baseFolder, TMN_SanitizeFileName(d & " " & vzyskName))
    TMN_EnsureFolder outRoot
    Set wsRes = TMN_CopyAsResultSheet(ws): TMN_PrepareResultHeaders wsRes

    Set dictPerRecipient = CreateObject("Scripting.Dictionary")
    Set dictFilesPerRecipient = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Set wdApp = GetObject(Class:="Word.Application")
    ownedWord = False
    If wdApp Is Nothing Then
        Set wdApp = CreateObject("Word.Application")
        ownedWord = True
    End If
    On Error GoTo 0
    If wdApp Is Nothing Then MsgBox "Не удалось запустить Word.", vbCritical: Exit Sub
    wdApp.Visible = False
    Application.ScreenUpdating = False

    totalToForm = TMN_CountBlocksReady(ws)
    formedCount = 0
    Application.StatusBar = "Формирование… 0/" & totalToForm

    r = 3
    Do While r <= lastRow
        Dim cellA As Range, recipient As String
        Dim rngInfo As Range, infoText As String
        Dim okToMake As Boolean, reason As String
        Dim doc As Object, outName As String, outPath As String, outFolder As String
        Dim colCount As Long

        Set cellA = ws.Cells(r, "A")
        blockRows = 1
        If cellA.MergeCells Then
            If cellA.Address <> cellA.MergeArea.Cells(1, 1).Address Then r = r + 1: GoTo NextR
            blockRows = cellA.MergeArea.rows.count
        End If

        recipient = Trim(TMN_GetMergedTopValue(ws.Cells(r, "C")))
        addr = Trim(TMN_GetMergedTopValue(ws.Cells(r, "B")))
        Set rngInfo = ws.Cells(r, "D").MergeArea

        okToMake = True: reason = ""
        If rngInfo Is Nothing Or LenB(CStr(rngInfo.Cells(1, 1).value)) = 0 Then okToMake = False: reason = reason & "пустой Инфо; "
        If Len(recipient) = 0 Then okToMake = False: reason = reason & "пустой Получатель; "

        If okToMake Then
            Dim idx As Long
            If Not dictPerRecipient.exists(recipient) Then dictPerRecipient.Add recipient, 0
            idx = dictPerRecipient(recipient) + 1
            dictPerRecipient(recipient) = idx

            outFolder = TMN_AppendPath(outRoot, TMN_SanitizeFileName(recipient))
            TMN_EnsureFolder outFolder

            outName = TMN_SanitizeFileName(vzyskName & " - " & recipient & " - " & CStr(idx) & ".docx")
            outPath = TMN_AppendPath(outFolder, outName)

            ' учёт пути для последующего объединения
            Dim coll As Collection
            If dictFilesPerRecipient.exists(recipient) Then
                Set coll = dictFilesPerRecipient(recipient)
            Else
                Set coll = New Collection
                dictFilesPerRecipient.Add recipient, coll
            End If
            coll.Add outPath

            ' Используем локальную копию шаблона
            Set doc = wdApp.Documents.Add(Template:=tplResolved)
            doc.Activate

            TMN_WordReplaceEverywhere doc, "{год}", y
            TMN_WordReplaceEverywhere doc, "{дата}", d
            TMN_WordReplaceEverywhere doc, "{Получатель}", recipient
            TMN_WordReplaceEverywhere doc, "{Представитель}", representative
            TMN_WordReplaceEverywhere doc, "{исх2}", ""    ' Тюмень не использует исходящий
            TMN_WordReplaceEverywhere doc, "{исх}", ""     ' Тюмень не использует исходящий
            TMN_WordReplaceEverywhere doc, "{доб}", addNum

            If Len(addr) > 0 Then TMN_WordReplaceEverywhere doc, "{Адрес}", addr Else TMN_WordReplaceEverywhere doc, "{Адрес}", ""

            colCount = blockRows
            TMN_WordReplaceEverywhere doc, "{кол}", CStr(colCount)

            infoText = CStr(rngInfo.Cells(1, 1).value)
            If Not TMN_ReplacePlaceholderWithPaste(wdApp, doc, "{Инфо}") Then doc.Activate
            TMN_InsertInfoSingleCell wdApp, infoText, doc

            doc.SaveAs2 fileName:=outPath, FileFormat:=wdFormatXMLDocument
            doc.Close SaveChanges:=False
            Set doc = Nothing

            formedCount = formedCount + 1
            Application.StatusBar = "Формирование… " & formedCount & "/" & totalToForm
            TMN_WriteResult wsRes, r, blockRows, "Сформировано", outName
        Else
            TMN_WriteResult wsRes, r, blockRows, "Не сформировано", "Ошибка: " & reason
        End If

        r = r + blockRows
NextR:
    Loop

    ' --- Создание общих файлов по папкам (по каждому получателю)
    Dim k As Variant
    For Each k In dictFilesPerRecipient.keys
        Dim filesColl As Collection
        Set filesColl = dictFilesPerRecipient(k)
        If filesColl.count > 1 Then ' <<< только если более 1 файла
            Dim combinedPath As String
            combinedPath = TMN_AppendPath(TMN_AppendPath(outRoot, TMN_SanitizeFileName(CStr(k))), _
                                          TMN_SanitizeFileName(vzyskName & " - " & CStr(k) & " - Общий.docx"))
            CreateCombinedDocFromList wdApp, filesColl, combinedPath
        End If
    Next k

    Application.StatusBar = "Формирование завершено: " & formedCount & "/" & totalToForm
    MsgBox "Готово. Сформировано: " & formedCount & " из " & totalToForm & "." & vbCrLf & _
           "Папка: " & outRoot & vbCrLf & "Результаты помечены на листе: " & wsRes.name, vbInformation

    ' Открываем папку с результатами
    On Error Resume Next
    shell "explorer.exe """ & outRoot & """", vbNormalFocus

CLEANUP:
    On Error Resume Next
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    If ownedWord And Not wdApp Is Nothing Then wdApp.Quit
    Application.StatusBar = False
    Set wdApp = Nothing
End Sub

' === ВХОД №2: по одному документу на получателя (первый пакет), WSS ========
Public Sub TMN_Make_Cover_Letters_ByRecipient()
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long, blockRows As Long
    Dim recipient As String, infoText As String, addr As String
    Dim templPath As String, vzyskName As String
    Dim tplResolved As String
    Dim representative As String
    Dim baseFolder As String, outRoot As String, outPath As String
    Dim wdApp As Object, doc As Object, ownedWord As Boolean
    Dim today As Date, y As String, d As String
    Dim dictFirstRow As Object, order As Collection
    Dim k As Variant, total As Long, done As Long
    Dim addNum As String

    ' === Проверка наличия листа "Реестр пакетов"
    If Not TMN_SheetExists("Реестр пакетов") Then
        MsgBox "Не найден лист ""Реестр пакетов"" в текущей книге." & vbCrLf & _
               "Пожалуйста, сформируйте лист с помощью кнопки Сформировать ""Реестр пакетов"" в Excel и повторите попытку.", _
               vbExclamation, "Ошибка: отсутствует лист 'Реестр пакетов'"
        Exit Sub
    End If

    ' Лист
    Set ws = ActiveWorkbook.Worksheets("Реестр пакетов")
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    If lastRow < 3 Then MsgBox "Нет данных на листе «Реестр пакетов».", vbExclamation: Exit Sub

    ' Первое вхождение каждого получателя
    Set dictFirstRow = CreateObject("Scripting.Dictionary")
    Set order = New Collection
    r = 3
    Do While r <= lastRow
        blockRows = 1
        If ws.Cells(r, "A").MergeCells Then
            If ws.Cells(r, "A").Address <> ws.Cells(r, "A").MergeArea.Cells(1, 1).Address Then r = r + 1: GoTo NextR
            blockRows = ws.Cells(r, "A").MergeArea.rows.count
        End If
        recipient = Trim(TMN_GetMergedTopValue(ws.Cells(r, "C")))
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

    ' Шаблон WSS для Тюмени
    If Not TMN_SelectTyumenVzysk(True, vzyskName, templPath) Then Exit Sub
    If Not TMN_SafeFileExists(templPath) Then MsgBox "Не найден шаблон Word:" & vbCrLf & templPath, vbCritical: Exit Sub

    ' Копируем шаблон в %AppData%\Microsoft\Templates и используем локальную копию
    tplResolved = TMN_ResolveTemplateToUserTemplates(templPath)

    ' Прочее
    representative = InputBox("Ваше Фамилия И.О. для поля {Представитель} подставлено автоматически. Если требуется другое, то введите его вручную:", "Представитель", "")
    addNum = InputBox("Ваш добавочный номер для поля {доб} подставлен автоматически. Если трубуется другой, то введите его вручную:", "Добавочный номер", "")

    baseFolder = TMN_PickFolder("Куда сохранить документы?")
    If Len(baseFolder) = 0 Then Exit Sub
    today = Date: y = CStr(Year(today)): d = Format$(today, "dd.mm.yyyy")
    outRoot = TMN_AppendPath(baseFolder, TMN_SanitizeFileName(d & " " & vzyskName))
    TMN_EnsureFolder outRoot

    ' Word
    On Error Resume Next
    Set wdApp = GetObject(Class:="Word.Application")
    If wdApp Is Nothing Then
        Set wdApp = CreateObject("Word.Application")
        ownedWord = True
    End If
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
        addr = Trim(TMN_GetMergedTopValue(ws.Cells(rTop, "B")))                    ' <-- АДРЕС
        infoText = CStr(ws.Cells(rTop, "D").MergeArea.Cells(1, 1).value)
        If LenB(infoText) = 0 Then GoTo SkipThis

        ' Используем локальную копию шаблона
        Set doc = wdApp.Documents.Add(Template:=tplResolved)

        TMN_WordReplaceEverywhere doc, "{год}", y
        TMN_WordReplaceEverywhere doc, "{дата}", d
        TMN_WordReplaceEverywhere doc, "{Получатель}", recipient
        TMN_WordReplaceEverywhere doc, "{Представитель}", representative
        TMN_WordReplaceEverywhere doc, "{исх}", ""     ' Тюмень не использует
        TMN_WordReplaceEverywhere doc, "{исх2}", ""    ' Тюмень не использует
        TMN_WordReplaceEverywhere doc, "{доб}", addNum

        ' {Адрес}
        If Len(addr) > 0 Then
            TMN_WordReplaceEverywhere doc, "{Адрес}", addr
        Else
            TMN_WordReplaceEverywhere doc, "{Адрес}", ""
        End If

        ' {кол} — размер блока (по первому пакету он не нужен, но поле может быть в шаблоне)
        TMN_WordReplaceEverywhere doc, "{кол}", "1"

        If Not TMN_ReplacePlaceholderWithPaste(wdApp, doc, "{Инфо}") Then
            doc.Range(doc.Content.End - 1, doc.Content.End - 1).Select
        End If
        TMN_InsertInfoSingleCell wdApp, infoText, doc

        outPath = TMN_AppendPath(outRoot, TMN_SanitizeFileName(vzyskName & " - " & recipient & ".docx"))
        doc.SaveAs2 fileName:=outPath, FileFormat:=wdFormatXMLDocument
        doc.Close SaveChanges:=False
        Set doc = Nothing

        done = done + 1
        Application.StatusBar = "Формирование… " & done & "/" & total
SkipThis:
    Next k

    If ownedWord Then wdApp.Quit
    Application.StatusBar = False
    MsgBox "Готово. Сформировано документов: " & done & " из " & total & vbCrLf & outRoot, vbInformation

    ' Открываем папку с результатами
    On Error Resume Next
    shell "explorer.exe """ & outRoot & """", vbNormalFocus
End Sub

' === Выбор взыскателя (Тюмень) ============================================
Private Function TMN_SelectTyumenVzysk(ByVal isWSS As Boolean, _
                                       ByRef vzyskName As String, _
                                       ByRef templPath As String) As Boolean
    Dim root As String
    Dim names() As String, filesBase() As String, filesWSS() As String
    Dim choice As Variant, idx As Long

    root = "\\corp.vostok-electra.ru\Tmn\Общая\ОВЗИД\Зуйкевич\ВЗИД\"
    names = Split("АО ЭК Восток|АО СУЭНКО|ООО ТЭО|ООО ТКС", "|")

    ReDim filesBase(0 To 3)
    filesBase(0) = "Восток Тюмень.docx"
    filesBase(1) = "СУЭНКО.docx"
    filesBase(2) = "ТЭО.docx"
    filesBase(3) = "ТКС.docx"

    ReDim filesWSS(0 To 3)
    filesWSS(0) = "Восток Тюмень WSS.docx"
    filesWSS(1) = "СУЭНКО WSS.docx"
    filesWSS(2) = "ТЭО WSS.docx"
    filesWSS(3) = "ТКС WSS.docx"

    choice = Application.InputBox( _
        prompt:="По какому взыскателю формировать?" & vbCrLf & _
                "1 – " & names(0) & vbCrLf & _
                "2 – " & names(1) & vbCrLf & _
                "3 – " & names(2) & vbCrLf & _
                "4 – " & names(3), _
        title:="Выбор взыскателя (Тюмень)", Type:=1)
    If VarType(choice) = vbBoolean And choice = False Then Exit Function
    idx = CLng(choice) - 1
    If idx < 0 Or idx > 3 Then
        MsgBox "Введите число 1..4.", vbExclamation
        Exit Function
    End If

    vzyskName = names(idx)
    templPath = root & IIf(isWSS, filesWSS(idx), filesBase(idx))
    TMN_SelectTyumenVzysk = True
End Function

' === Предскан: отсутствие Получатель/Инфо =================================
Private Sub TMN_PreScanMissing(ByVal ws As Worksheet, _
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
                r = r + 1: GoTo NextR
            End If
            blockRows = cellA.MergeArea.rows.count
        End If

        recipient = Trim(TMN_GetMergedTopValue(ws.Cells(r, "C")))
        Set rngInfo = ws.Cells(r, "D").MergeArea
        If Len(recipient) = 0 Then
            missRec = missRec + 1
            rowsMissRec.Add r
        End If
        If (rngInfo Is Nothing) Or LenB(CStr(rngInfo.Cells(1, 1).value)) = 0 Then
            missInfo = missInfo + 1
            rowsMissInfo.Add r
        End If

        r = r + blockRows
NextR:
    Loop
End Sub

Private Sub TMN_HighlightMissing(ByVal ws As Worksheet, ByVal rowsMissRec As Collection, ByVal rowsMissInfo As Collection)
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

' === Подсчёты =============================================================
Private Function TMN_CountBlocks(ByVal ws As Worksheet) As Long
    Dim r As Long, lastRow As Long, cellA As Range, blockRows As Long, total As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    r = 3: total = 0
    Do While r <= lastRow
        Set cellA = ws.Cells(r, "A")
        blockRows = 1
        If cellA.MergeCells Then
            If cellA.Address <> cellA.MergeArea.Cells(1, 1).Address Then
                r = r + 1: GoTo NextR
            End If
            blockRows = cellA.MergeArea.rows.count
        End If
        total = total + 1
        r = r + blockRows
NextR:
    Loop
    TMN_CountBlocks = total
End Function

Private Function TMN_CountBlocksReady(ByVal ws As Worksheet) As Long
    Dim r As Long, lastRow As Long, cellA As Range, blockRows As Long, total As Long
    Dim recipient As String, rngInfo As Range
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    r = 3: total = 0
    Do While r <= lastRow
        Set cellA = ws.Cells(r, "A")
        blockRows = 1
        If cellA.MergeCells Then
            If cellA.Address <> cellA.MergeArea.Cells(1, 1).Address Then
                r = r + 1: GoTo NextR
            End If
            blockRows = cellA.MergeArea.rows.count
        End If
        recipient = Trim(TMN_GetMergedTopValue(ws.Cells(r, "C")))
        Set rngInfo = ws.Cells(r, "D").MergeArea
        If Len(recipient) > 0 And (Not (rngInfo Is Nothing)) And LenB(CStr(rngInfo.Cells(1, 1).value)) > 0 Then
            total = total + 1
        End If
        r = r + blockRows
NextR:
    Loop
    TMN_CountBlocksReady = total
End Function

' === Копия листа результатов ==============================================
Private Function TMN_CopyAsResultSheet(ByVal ws As Worksheet) As Worksheet
    Dim nm As String
    nm = TMN_UniqueSheetName("Результат формирования")
    ws.Copy After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.count)
    Set TMN_CopyAsResultSheet = ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.count)
    On Error Resume Next
    TMN_CopyAsResultSheet.name = nm
    On Error GoTo 0
End Function

Private Sub TMN_PrepareResultHeaders(ByVal wsRes As Worksheet)
    wsRes.Range("E2").value = "Результат"
    wsRes.Range("F2").value = "Описание"
    wsRes.Range("E2:F2").HorizontalAlignment = xlCenter
    wsRes.Range("E2:F2").VerticalAlignment = xlCenter
    wsRes.Columns("E").ColumnWidth = 16
    wsRes.Columns("F").ColumnWidth = 60
End Sub

Private Sub TMN_WriteResult(ByVal wsRes As Worksheet, ByVal topRow As Long, ByVal blockRows As Long, _
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

' === Утилиты Excel/FS =====================================================
Private Function TMN_PickFolder(Optional ByVal title As String = "Выбор папки") As String
    Dim fd As FileDialog
    On Error GoTo FAIL
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .title = title
        .AllowMultiSelect = False
        If .Show = -1 Then TMN_PickFolder = .SelectedItems(1)
    End With
    Exit Function
FAIL:
    TMN_PickFolder = ""
End Function

Private Function TMN_AppendPath(base As String, tail As String) As String
    If Right$(base, 1) = "\" Or Right$(base, 1) = "/" Then
        TMN_AppendPath = base & tail
    Else
        TMN_AppendPath = base & "\" & tail
    End If
End Function

Private Sub TMN_EnsureFolder(ByVal path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(path) Then fso.CreateFolder path
End Sub

Private Function TMN_SanitizeFileName(ByVal s As String) As String
    Dim badChars As Variant, ch As Variant
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In badChars
        s = Replace$(s, CStr(ch), "_")
    Next ch
    Do While Len(s) > 0 And (Right$(s, 1) = " " Or Right$(s, 1) = ".")
        s = Left$(s, Len(s) - 1)
    Loop
    If Len(s) = 0 Then s = "_"
    TMN_SanitizeFileName = s
End Function

Private Function TMN_GetMergedTopValue(ByVal anyCell As Range) As String
    If anyCell.MergeCells Then
        TMN_GetMergedTopValue = CStr(anyCell.MergeArea.Cells(1, 1).value)
    Else
        TMN_GetMergedTopValue = CStr(anyCell.value)
    End If
End Function

' === Word-хелперы =========================================================
Private Sub TMN_WordReplaceEverywhere(ByVal oDoc As Object, ByVal findText As String, ByVal replText As String)
    Dim rngStory As Object, rng As Object
    For Each rngStory In oDoc.StoryRanges
        Set rng = rngStory
        Do While Not rng Is Nothing
            With rng.Find
                .ClearFormatting: .Replacement.ClearFormatting
                .Text = findText
                .Replacement.Text = replText
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
            Set rng = rng.NextStoryRange
        Loop
    Next rngStory
End Sub

Private Function TMN_ReplacePlaceholderWithPaste(ByVal wdApp As Object, ByVal doc As Object, ByVal placeholder As String) As Boolean
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
                TMN_ReplacePlaceholderWithPaste = True
                Exit Function
            End If
        End With
    End With
    TMN_ReplacePlaceholderWithPaste = False
End Function

' === {Инфо}: одна ячейка 17,5 см, TNR 12 пт ===============================
Private Sub TMN_InsertInfoSingleCell(ByVal wdApp As Object, ByVal infoText As String, ByVal doc As Object)
    On Error Resume Next
    Dim t As String
    t = Replace(infoText, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    t = Replace(t, vbLf, vbCr) ' стиль Word

    Dim tbl As Object
    Set tbl = doc.Tables.Add(wdApp.Selection.Range, 1, 1)

    Dim targetW As Single: targetW = wdApp.CentimetersToPoints(17.5!)
    tbl.AllowAutoFit = False
    tbl.AutoFitBehavior wdAutoFitFixed
    tbl.PreferredWidthType = wdPreferredWidthPoints
    tbl.PreferredWidth = targetW
    tbl.Columns(1).width = targetW
    tbl.rows.LeftIndent = 0
    tbl.LeftPadding = 0: tbl.RightPadding = 0

    With tbl.Borders
        .OutsideLineStyle = wdLineStyleSingle
        .InsideLineStyle = wdLineStyleNone
    End With

    With tbl.cell(1, 1).Range
        .Text = t
        .Font.name = "Times New Roman"
        .Font.Size = 12
        With .ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            .LeftIndent = 0
            .RightIndent = 0
            .SpaceBefore = 0
            .SpaceAfter = 0
            .LineSpacingRule = 0
        End With
    End With
End Sub

' === Разное ===============================================================
Private Function TMN_UniqueSheetName(baseName As String) As String
    Dim nm As String, i As Long
    nm = baseName: i = 2
    Do While TMN_SheetExists(nm)
        nm = baseName & "_" & i: i = i + 1
    Loop
    TMN_UniqueSheetName = nm
End Function

Private Function TMN_SheetExists(nm As String) As Boolean
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ActiveWorkbook.Worksheets(nm)
    TMN_SheetExists = Not sh Is Nothing
    On Error GoTo 0
End Function

Private Function TMN_SafeFileExists(ByVal p As String) As Boolean
    On Error Resume Next
    TMN_SafeFileExists = CreateObject("Scripting.FileSystemObject").FileExists(p)
End Function

Private Function TMN_UniqueSubfolderPath(ByVal baseFolder As String, ByVal folderName As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim basePath As String, n As Long, p As String
    If Right$(baseFolder, 1) = "\" Or Right$(baseFolder, 1) = "/" Then
        basePath = baseFolder
    Else
        basePath = baseFolder & "\"
    End If
    p = basePath & folderName
    If Not fso.FolderExists(p) Then TMN_UniqueSubfolderPath = p: Exit Function
    n = 1
    Do
        p = basePath & folderName & " - " & n
        If Not fso.FolderExists(p) Then TMN_UniqueSubfolderPath = p: Exit Function
        n = n + 1
    Loop
End Function

' === Шаблон: копия в %AppData%\Microsoft\Templates (обновление каждый запуск) ===
Private Function TMN_ResolveTemplateToUserTemplates(ByVal netPath As String) As String
    ' Возвращает путь к локальной копии шаблона в %AppData%\Microsoft\Templates
    ' Копирует/перезаписывает шаблон при каждом запуске, чтобы учитывать изменения в сетевом шаблоне
    On Error GoTo ErrHandler

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim appData As String
    appData = Environ$("AppData")
    If Right$(appData, 1) <> "\" Then appData = appData & "\"
    Dim templatesFolder As String
    templatesFolder = appData & "Microsoft\Templates\"

    ' Убедимся, что папка существует
    If Not fso.FolderExists(templatesFolder) Then
        On Error Resume Next
        fso.CreateFolder templatesFolder
        On Error GoTo ErrHandler
    End If

    ' Проверка исходного файла
    If Len(Trim$(netPath)) = 0 Or Not fso.FileExists(netPath) Then
        TMN_ResolveTemplateToUserTemplates = netPath
        Exit Function
    End If

    ' Локальный путь: сохраняем исходное имя файла
    Dim localName As String
    localName = fso.GetFileName(netPath)
    Dim localPath As String
    localPath = templatesFolder & localName

    ' Копируем с перезаписью (обновляем каждый запуск)
    fso.CopyFile netPath, localPath, True

    TMN_ResolveTemplateToUserTemplates = localPath
    Exit Function

ErrHandler:
    On Error Resume Next
    ' В случае проблем возвращаем исходный путь — Word сам покажет понятную ошибку
    TMN_ResolveTemplateToUserTemplates = netPath
End Function

' === Создать общий документ из списка файлов (с сохранением форматирования и колонтитулов) ===
Private Sub CreateCombinedDocFromList(ByVal wdApp As Object, ByVal filesColl As Collection, ByVal destPath As String)
    On Error GoTo ErrHandler
    ' Если 0 или 1 файл — общий НЕ создаём (по требованию)
    If filesColl.count <= 1 Then Exit Sub

    Dim i As Long
    Dim baseDoc As Object, srcDoc As Object, lastSection As Object

    ' Открыть первый файл как базовый комбинируемый документ (ReadOnly:=False)
    Set baseDoc = wdApp.Documents.Open(fileName:=filesColl(1), ReadOnly:=False, AddToRecentFiles:=False, Visible:=False)

    ' Для каждого следующего файла: вставить секционный разрыв и содержимое, скопировать колонтитулы
    For i = 2 To filesColl.count
        Set srcDoc = wdApp.Documents.Open(fileName:=filesColl(i), ReadOnly:=True, AddToRecentFiles:=False, Visible:=False)

        baseDoc.Activate
        With wdApp.Selection
            .EndKey Unit:=wdStory
        End With
        wdApp.Selection.InsertBreak Type:=wdSectionBreakNextPage
        wdApp.Selection.Range.FormattedText = srcDoc.Content.FormattedText

        Set lastSection = baseDoc.Sections(baseDoc.Sections.count)
        On Error Resume Next
        Dim hfType As Variant
        For Each hfType In Array(wdHeaderFooterPrimary, wdHeaderFooterFirstPage, wdHeaderFooterEvenPages)
            lastSection.headers(hfType).LinkToPrevious = False
            lastSection.Footers(hfType).LinkToPrevious = False
            lastSection.headers(hfType).Range.FormattedText = srcDoc.Sections(1).headers(hfType).Range.FormattedText
            lastSection.Footers(hfType).Range.FormattedText = srcDoc.Sections(1).Footers(hfType).Range.FormattedText
        Next hfType
        On Error GoTo ErrHandler

        srcDoc.Close SaveChanges:=False
        Set srcDoc = Nothing
    Next i

    baseDoc.SaveAs2 fileName:=destPath, FileFormat:=wdFormatXMLDocument
    baseDoc.Close SaveChanges:=False
    Set baseDoc = Nothing
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not baseDoc Is Nothing Then baseDoc.Close SaveChanges:=False
    If Not srcDoc Is Nothing Then srcDoc.Close SaveChanges:=False
End Sub


