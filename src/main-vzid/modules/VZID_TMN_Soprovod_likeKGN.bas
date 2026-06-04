Attribute VB_Name = "VZID_TMN_Soprovod_likeKGN"
Option Explicit

' ==========================================================================================
' === CONFIGURATION CONSTANTS (НАСТРОЙКИ ПУТЕЙ И ФАЙЛОВ) _likeKGN
' ==========================================================================================
Private Const TMN_ROOT_PATH_likeKGN As String = "\\corp.vostok-electra.ru\Tmn\Общая\ОВЗИД\Зуйкевич\ВЗИД\"

' Имена взыскателей (для меню выбора) через разделитель "|"
Private Const TMN_VZYSK_NAMES_likeKGN As String = "АО ЭК Восток|АО СУЭНКО|ООО ТЭО|ООО ТКС"

' Имена файлов шаблонов (ОБЫЧНЫЕ) через разделитель "|"
Private Const TMN_FILES_STD_likeKGN As String = "Сопровод Восток Тюмень.docx|Сопровод СУЭНКО.docx|Сопровод ТЭО.docx|Сопровод ТКС.docx"

' Имена файлов шаблонов (WSS) через разделитель "|"
Private Const TMN_FILES_WSS_likeKGN As String = "Сопровод Восток Тюмень для WSS.docx|Сопровод СУЭНКО для WSS.docx|Сопровод ТЭО для WSS.docx|Сопровод ТКС для WSS.docx"
' ==========================================================================================

' --- Word constants (для позднего связывания) ---
Private Const wdFormatXMLDocument_likeKGN As Long = 12
Private Const wdFormatOriginalFormatting_likeKGN As Long = 16
Private Const wdFindContinue_likeKGN As Long = 1
Private Const wdReplaceAll_likeKGN As Long = 2
Private Const wdStory_likeKGN As Long = 6
Private Const wdAutoFitFixed_likeKGN As Long = 0
Private Const wdAutoFitContent_likeKGN As Long = 1
Private Const wdPreferredWidthPercent_likeKGN As Long = 2
Private Const wdLineStyleSingle_likeKGN As Long = 1
Private Const wdLineStyleNone_likeKGN As Long = 0
Private Const wdAlignParagraphLeft_likeKGN As Long = 0
Private Const wdAlignParagraphCenter_likeKGN As Long = 1
Private Const wdNoHighlight_likeKGN As Long = 0
Private Const wdPreferredWidthPoints_likeKGN As Long = 3
Private Const wdSectionBreakNextPage_likeKGN As Long = 2
Private Const wdHeaderFooterPrimary_likeKGN As Long = 1
Private Const wdHeaderFooterFirstPage_likeKGN As Long = 2
Private Const wdHeaderFooterEvenPages_likeKGN As Long = 3


' ==========================================================================================
' === ТОЧКА ВХОДА 1: Пакетное формирование (Все строки) + Склейка файлов
' ==========================================================================================
Sub TMN_Make_Cover_Letters_likeKGN()
    Dim ws As Worksheet, wsRes As Worksheet
    Dim lastRow As Long, r As Long, blockRows As Long
    Dim templPath As String, vzyskName As String
    Dim baseFolder As String, outRoot As String
    Dim dictPerRecipient As Object
    Dim dictFilesPerRecipient As Object ' ключ -> Collection путей файлов
    Dim wdApp As Object, ownedWord As Boolean
    Dim today As Date, y As String, d As String
    Dim totalToForm As Long, formedCount As Long
    Dim representative As String, defaultRep As String
    Dim outBase As String
    Dim addNum As String, addr As String, defaultAddNum As String
    Dim docSerial As Long
    Dim colCount As Long

    ' --- Проверка наличия листа "Реестр пакетов"
    If Not TMN_SheetExists_likeKGN("Реестр пакетов") Then
        MsgBox "Не найден лист ""Реестр пакетов"" в текущей книге." & vbCrLf & _
               "Пожалуйста, сформируйте лист и повторите попытку.", _
               vbExclamation, "Ошибка"
        Exit Sub
    End If

    ' --- Лист
    Set ws = ActiveWorkbook.Worksheets("Реестр пакетов")
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    If lastRow < 3 Then MsgBox "На листе «Реестр пакетов» нет данных.", vbExclamation: Exit Sub

    ' --- Выбор взыскателя/шаблона (Тюмень)
    Dim okSel As Boolean
    okSel = TMN_SelectVzyskAndTemplate_likeKGN(False, vzyskName, templPath)
    If Not okSel Then Exit Sub
    If Not TMN_SafeFileExists_likeKGN(templPath) Then MsgBox "Не найден шаблон Word:" & vbCrLf & templPath, vbCritical: Exit Sub

    ' --- Получение данных из AD (автозаполнение) ---
    TMN_GetUserInfoFromAD_likeKGN defaultRep, defaultAddNum

    ' --- Ввод данных ---
    outBase = InputBox("Введите исходящий номер (напр., 119960)." & vbCrLf & "Оставьте пустым, если номер не нужен.", "Исходящий номер", "")
    
    ' Подставляем найденные значения в InputBox как Default
    representative = InputBox("Введите Фамилия И.О. для поля {Представитель}:", "Представитель", defaultRep)
    addNum = InputBox("Введите добавочный номер (если есть). Будет подставлен в {доб}.", "Добавочный номер", defaultAddNum)
    
    docSerial = 1

    ' --- Предскан ошибок
    Dim missRec As Long, missInfo As Long, rowsMissRec As Collection, rowsMissInfo As Collection
    Dim resp As VbMsgBoxResult, totalBlocksAll As Long, msg As String
    totalBlocksAll = TMN_CountBlocks_likeKGN(ws, False)
    TMN_PreScanMissing_likeKGN ws, False, missRec, missInfo, rowsMissRec, rowsMissInfo
    If (missRec > 0) Or (missInfo > 0) Then
        If missRec > 0 Then msg = msg & "По " & missRec & " из " & totalBlocksAll & " документов отсутствует Получатель" & vbCrLf
        If missInfo > 0 Then msg = msg & "По " & missInfo & " из " & totalBlocksAll & " документов отсутствует Инфо" & vbCrLf
        msg = msg & vbCrLf & "Да – сформировать по тем, где все данные есть" & vbCrLf & "Нет – не формировать ничего"
        resp = MsgBox(msg, vbQuestion + vbYesNo, "Проверка заполненности")
        If resp = vbNo Then TMN_HighlightMissing_likeKGN ws, rowsMissRec, rowsMissInfo: Exit Sub
    End If

    ' --- Папка
    baseFolder = TMN_PickFolder_likeKGN("Куда сохранить документы?")
    If Len(baseFolder) = 0 Then Exit Sub
    today = Date: y = CStr(Year(today)): d = Format$(today, "dd.mm.yyyy")
    outRoot = TMN_UniqueSubfolderPath_likeKGN(baseFolder, TMN_SanitizeFileName_likeKGN(d & " " & vzyskName))
    TMN_EnsureFolder_likeKGN outRoot

    ' --- Лист результатов
    Set wsRes = TMN_CopyAsResultSheet_likeKGN(ws): TMN_PrepareResultHeaders_likeKGN wsRes

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

    totalToForm = TMN_CountBlocksReady_likeKGN(ws, False)
    formedCount = 0
    Application.StatusBar = "Формирование… 0/" & totalToForm

    ' --- Подготовка локальной копии шаблона
    Dim tplPathResolved As String
    tplPathResolved = TMN_ResolveTemplateForWord_likeKGN(templPath)

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

        recipient = Trim(TMN_GetMergedTopValue_likeKGN(ws.Cells(r, "C")))
        addr = Trim(TMN_GetMergedTopValue_likeKGN(ws.Cells(r, "B")))
        Set rngInfo = ws.Cells(r, "D").MergeArea

        okToMake = True: reason = ""
        If rngInfo Is Nothing Or LenB(CStr(rngInfo.Cells(1, 1).value)) = 0 Then okToMake = False: reason = reason & "пустой Инфо; "
        If Len(recipient) = 0 Then okToMake = False: reason = reason & "пустой Получатель; "

        If okToMake Then
            Dim idx As Long
            If Not dictPerRecipient.exists(recipient) Then dictPerRecipient.Add recipient, 0
            idx = dictPerRecipient(recipient) + 1
            dictPerRecipient(recipient) = idx

            outFolder = TMN_AppendPath_likeKGN(outRoot, TMN_SanitizeFileName_likeKGN(recipient))
            TMN_EnsureFolder_likeKGN outFolder
            outName = TMN_SanitizeFileName_likeKGN(vzyskName & " - " & recipient & " - " & CStr(idx) & ".docx")
            outPath = TMN_AppendPath_likeKGN(outFolder, outName)

            ' Сохраняем путь для последующей склейки
            Dim coll As Collection
            If dictFilesPerRecipient.exists(recipient) Then
                Set coll = dictFilesPerRecipient(recipient)
            Else
                Set coll = New Collection
                dictFilesPerRecipient.Add recipient, coll
            End If
            coll.Add outPath

            Set doc = wdApp.Documents.Add(Template:=tplPathResolved)

            ' Плейсхолдеры
            TMN_WordReplaceEverywhere_likeKGN doc, "{год}", y
            TMN_WordReplaceEverywhere_likeKGN doc, "{дата}", d
            TMN_WordReplaceEverywhere_likeKGN doc, "{Получатель}", recipient
            TMN_WordReplaceEverywhere_likeKGN doc, "{Представитель}", representative

            ' Логика Исходящего (как в KGN)
            TMN_WordReplaceEverywhere_likeKGN doc, "{исх2}", outBase
            If Len(outBase) > 0 Then
                TMN_WordReplaceEverywhere_likeKGN doc, "{исх}", outBase & "/" & CStr(docSerial)
            Else
                TMN_WordReplaceEverywhere_likeKGN doc, "{исх}", ""
            End If

            TMN_WordReplaceEverywhere_likeKGN doc, "{доб}", addNum

            ' {Адрес}
            If Len(addr) > 0 Then TMN_WordReplaceEverywhere_likeKGN doc, "{Адрес}", addr _
                              Else TMN_WordReplaceEverywhere_likeKGN doc, "{Адрес}", ""

            ' {кол}
            colCount = blockRows
            TMN_WordReplaceEverywhere_likeKGN doc, "{кол}", CStr(colCount)

            ' {Инфо}
            infoText = CStr(rngInfo.Cells(1, 1).value)
            If Not TMN_ReplacePlaceholderWithPaste_likeKGN(wdApp, doc, "{Инфо}") Then doc.Activate
            TMN_InsertInfoSingleCell_likeKGN wdApp, infoText, doc

            doc.SaveAs2 fileName:=outPath, FileFormat:=wdFormatXMLDocument_likeKGN
            doc.Close SaveChanges:=False
            Set doc = Nothing

            formedCount = formedCount + 1
            docSerial = docSerial + 1
            Application.StatusBar = "Формирование… " & formedCount & "/" & totalToForm
            TMN_WriteResult_likeKGN wsRes, r, blockRows, "Сформировано", outName
        Else
            TMN_WriteResult_likeKGN wsRes, r, blockRows, "Не сформировано", "Ошибка: " & reason
        End If

        r = r + blockRows
NextR:
    Loop

    ' --- Создание общих файлов (Logic from KGN)
    Dim k As Variant
    For Each k In dictFilesPerRecipient.keys
        Dim filesColl As Collection
        Set filesColl = dictFilesPerRecipient(k)
        If filesColl.count > 0 Then
            Dim combinedPath As String
            combinedPath = TMN_AppendPath_likeKGN(TMN_AppendPath_likeKGN(outRoot, TMN_SanitizeFileName_likeKGN(CStr(k))), TMN_SanitizeFileName_likeKGN(vzyskName & " - " & CStr(k) & " - Общий.docx"))
            TMN_CreateCombinedDocFromList_likeKGN wdApp, filesColl, combinedPath
        End If
    Next k

    Application.StatusBar = "Формирование завершено: " & formedCount & "/" & totalToForm
    MsgBox "Готово. Сформировано: " & formedCount & " из " & totalToForm & "." & vbCrLf & _
           "Папка: " & outRoot & vbCrLf & "Результаты помечены на листе: " & wsRes.name, vbInformation

    shell "explorer.exe """ & outRoot & """", vbNormalFocus

CLEANUP:
    On Error Resume Next
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    If ownedWord And Not wdApp Is Nothing Then wdApp.Quit
    Application.StatusBar = False
    Set wdApp = Nothing
End Sub


' ==========================================================================================
' === ТОЧКА ВХОДА 2: Режим WSS (1 файл на получателя, без {исх})
' ==========================================================================================
Sub TMN_Make_Cover_Letters_ByRecipient_likeKGN()
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long, blockRows As Long
    Dim recipient As String, infoText As String, addr As String
    Dim templPath As String, vzyskName As String
    Dim representative As String, defaultRep As String
    Dim baseFolder As String, outRoot As String, outPath As String
    Dim wdApp As Object, doc As Object, ownedWord As Boolean
    Dim today As Date, y As String, d As String
    Dim dictFirstRow As Object, order As Collection
    Dim k As Variant, total As Long, done As Long
    Dim addNum As String, defaultAddNum As String
    Dim colCount As Long

    If Not TMN_SheetExists_likeKGN("Реестр пакетов") Then
        MsgBox "Не найден лист ""Реестр пакетов"".", vbExclamation
        Exit Sub
    End If

    Set ws = ActiveWorkbook.Worksheets("Реестр пакетов")
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    If lastRow < 3 Then MsgBox "Нет данных на листе.", vbExclamation: Exit Sub

    ' Сбор уникальных получателей (берем первую строку появления)
    Set dictFirstRow = CreateObject("Scripting.Dictionary")
    Set order = New Collection
    r = 3
    Do While r <= lastRow
        blockRows = 1
        If ws.Cells(r, "A").MergeCells Then
            If ws.Cells(r, "A").Address <> ws.Cells(r, "A").MergeArea.Cells(1, 1).Address Then r = r + 1: GoTo NextR
            blockRows = ws.Cells(r, "A").MergeArea.rows.count
        End If
        recipient = Trim(TMN_GetMergedTopValue_likeKGN(ws.Cells(r, "C")))
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

    ' Выбор WSS шаблона (True)
    Dim okSel As Boolean
    okSel = TMN_SelectVzyskAndTemplate_likeKGN(True, vzyskName, templPath)
    If Not okSel Then Exit Sub
    If Not TMN_SafeFileExists_likeKGN(templPath) Then MsgBox "Не найден шаблон Word:" & vbCrLf & templPath, vbCritical: Exit Sub

    ' --- Получение данных из AD (автозаполнение) ---
    TMN_GetUserInfoFromAD_likeKGN defaultRep, defaultAddNum

    ' --- Ввод данных ---
    representative = InputBox("Введите Фамилия И.О. для поля {Представитель}:", "Представитель", defaultRep)
    addNum = InputBox("Введите добавочный номер (если есть).", "Добавочный номер", defaultAddNum)

    baseFolder = TMN_PickFolder_likeKGN("Куда сохранить документы?")
    If Len(baseFolder) = 0 Then Exit Sub
    today = Date: y = CStr(Year(today)): d = Format$(today, "dd.mm.yyyy")
    outRoot = TMN_UniqueSubfolderPath_likeKGN(baseFolder, TMN_SanitizeFileName_likeKGN(d & " " & vzyskName & " для загрузки в WSS"))
    TMN_EnsureFolder_likeKGN outRoot

    On Error Resume Next
    Set wdApp = GetObject(Class:="Word.Application")
    If wdApp Is Nothing Then Set wdApp = CreateObject("Word.Application"): ownedWord = True
    On Error GoTo 0
    If wdApp Is Nothing Then MsgBox "Не удалось запустить Word.", vbCritical: Exit Sub
    wdApp.Visible = False

    total = order.count: done = 0
    Application.StatusBar = "Формирование… 0/" & total

    ' Подготовка шаблона
    Dim tplPathResolved As String
    tplPathResolved = TMN_ResolveTemplateForWord_likeKGN(templPath)

    Dim rTop As Long

    For Each k In order
        rTop = CLng(dictFirstRow(CStr(k)))
        recipient = CStr(k)
        addr = Trim(TMN_GetMergedTopValue_likeKGN(ws.Cells(rTop, "B")))
        infoText = CStr(ws.Cells(rTop, "D").MergeArea.Cells(1, 1).value)
        If LenB(infoText) = 0 Then GoTo SkipThis

        Set doc = wdApp.Documents.Add(Template:=tplPathResolved)

        TMN_WordReplaceEverywhere_likeKGN doc, "{год}", y
        TMN_WordReplaceEverywhere_likeKGN doc, "{дата}", d
        TMN_WordReplaceEverywhere_likeKGN doc, "{Получатель}", recipient
        TMN_WordReplaceEverywhere_likeKGN doc, "{Представитель}", representative

        ' В режиме WSS очищаем исходящие
        TMN_WordReplaceEverywhere_likeKGN doc, "{исх}", ""
        TMN_WordReplaceEverywhere_likeKGN doc, "{исх2}", ""
        TMN_WordReplaceEverywhere_likeKGN doc, "{доб}", addNum

        If Len(addr) > 0 Then TMN_WordReplaceEverywhere_likeKGN doc, "{Адрес}", addr _
                          Else TMN_WordReplaceEverywhere_likeKGN doc, "{Адрес}", ""

        ' {кол} - берем высоту объединенной ячейки столбца D первой строки
        If ws.Cells(rTop, "D").MergeCells Then
            colCount = ws.Cells(rTop, "D").MergeArea.rows.count
        Else
            colCount = 1
        End If
        TMN_WordReplaceEverywhere_likeKGN doc, "{кол}", CStr(colCount)

        If Not TMN_ReplacePlaceholderWithPaste_likeKGN(wdApp, doc, "{Инфо}") Then doc.Range(doc.Content.End - 1, doc.Content.End - 1).Select
        TMN_InsertInfoSingleCell_likeKGN wdApp, infoText, doc

        outPath = TMN_AppendPath_likeKGN(outRoot, TMN_SanitizeFileName_likeKGN(vzyskName & " - " & recipient & ".docx"))
        doc.SaveAs2 fileName:=outPath, FileFormat:=wdFormatXMLDocument_likeKGN
        doc.Close SaveChanges:=False
        Set doc = Nothing

        done = done + 1
        Application.StatusBar = "Формирование… " & done & "/" & total
SkipThis:
    Next k

    If ownedWord Then wdApp.Quit
    Application.StatusBar = False
    MsgBox "Готово. Сформировано документов: " & done & " из " & total & vbCrLf & outRoot, vbInformation
    shell "explorer.exe """ & outRoot & """", vbNormalFocus
End Sub


' ==========================================================================================
' === ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (ВЫБОР, ПРОВЕРКИ, WORD, AD) _likeKGN
' ==========================================================================================

' --- Получение данных пользователя из AD (Исправлено получение Отчества) ---
Private Sub TMN_GetUserInfoFromAD_likeKGN(ByRef outFIO As String, ByRef outExt As String)
    On Error GoTo Handler
    Dim objSysInfo As Object
    Dim objConnection As Object, objCommand As Object, objRecordSet As Object
    Dim strDN As String
    Dim strSurname As String, strGiven As String, strMiddle As String, strDisplayName As String
    Dim strPhone As String
    
    ' Получаем имя текущего пользователя
    Set objSysInfo = CreateObject("ADSystemInfo")
    strDN = objSysInfo.userName
    
    ' Подключаемся к AD
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
    Set objCommand.ActiveConnection = objConnection
    
    ' Добавили displayName в запрос
    objCommand.CommandText = "SELECT sn, givenName, middleName, displayName, telephoneNumber FROM 'LDAP://" & strDN & "'"
    
    Set objRecordSet = objCommand.Execute
    
    If Not objRecordSet.EOF Then
        If Not IsNull(objRecordSet.fields("sn").value) Then strSurname = Trim(objRecordSet.fields("sn").value)
        If Not IsNull(objRecordSet.fields("givenName").value) Then strGiven = Trim(objRecordSet.fields("givenName").value)
        If Not IsNull(objRecordSet.fields("middleName").value) Then strMiddle = Trim(objRecordSet.fields("middleName").value)
        If Not IsNull(objRecordSet.fields("displayName").value) Then strDisplayName = Trim(objRecordSet.fields("displayName").value)
        If Not IsNull(objRecordSet.fields("telephoneNumber").value) Then strPhone = objRecordSet.fields("telephoneNumber").value
    End If
    
    ' === ЛОГИКА ВОССТАНОВЛЕНИЯ ОТЧЕСТВА ===
    ' Если прямого поля middleName нет, пытаемся найти его в displayName
    If Len(strMiddle) = 0 And Len(strDisplayName) > 0 Then
        Dim parts() As String
        Dim part As Variant
        parts = Split(strDisplayName, " ")
        
        For Each part In parts
            Dim s As String
            s = Trim(part)
            ' Если слово не совпадает ни с Фамилией, ни с Именем - считаем его Отчеством
            If Len(s) > 0 And _
               StrComp(s, strSurname, vbTextCompare) <> 0 And _
               StrComp(s, strGiven, vbTextCompare) <> 0 Then
                strMiddle = s
                Exit For
            End If
        Next
    End If
    ' ======================================
    
    ' Формирование "Фамилия И.О."
    Dim fio As String
    fio = strSurname
    
    If Len(fio) > 0 And Len(strGiven) > 0 Then
        ' Добавляем инициал имени
        fio = fio & " " & Left(strGiven, 1) & "."
        ' Добавляем инициал отчества (если нашли)
        If Len(strMiddle) > 0 Then
            fio = fio & Left(strMiddle, 1) & "."
        End If
    End If
    outFIO = fio
    
    ' Парсинг добавочного номера
    Dim extNum As String
    Dim p As Long
    Dim tempStr As String
    
    tempStr = LCase(strPhone)
    p = InStr(1, tempStr, "доб.")
    
    If p > 0 Then
        tempStr = Mid(tempStr, p + 4) ' пропускаем "доб."
        tempStr = Trim(tempStr)
        If Len(tempStr) >= 4 Then
            extNum = Left(tempStr, 4)
        Else
            extNum = tempStr
        End If
    Else
        extNum = ""
    End If
    
    outExt = extNum
    
    objConnection.Close
    Exit Sub
    
Handler:
    outFIO = ""
    outExt = ""
    Err.Clear
End Sub

' --- Выбор взыскателя и шаблона (использует константы из шапки) ---
Private Function TMN_SelectVzyskAndTemplate_likeKGN(ByVal isWSS As Boolean, _
                                            ByRef vzyskName As String, _
                                            ByRef templPath As String) As Boolean
    Dim names() As String, filesBase() As String, filesWSS() As String
    Dim q As Variant, idx As Long
    Dim promptStr As String, i As Long

    names = Split(TMN_VZYSK_NAMES_likeKGN, "|")
    filesBase = Split(TMN_FILES_STD_likeKGN, "|")
    filesWSS = Split(TMN_FILES_WSS_likeKGN, "|")

    promptStr = "По какому взыскателю формировать?" & vbCrLf
    For i = LBound(names) To UBound(names)
        promptStr = promptStr & (i + 1) & " – " & names(i) & vbCrLf
    Next i

    q = Application.InputBox(promptStr, "Выбор взыскателя (Тюмень)", Type:=1)
    If VarType(q) = vbBoolean And q = False Then Exit Function

    idx = CLng(q) - 1
    If idx < LBound(names) Or idx > UBound(names) Then
        MsgBox "Неверный номер.", vbExclamation
        Exit Function
    End If

    vzyskName = names(idx)
    If isWSS Then
        templPath = TMN_ROOT_PATH_likeKGN & filesWSS(idx)
    Else
        templPath = TMN_ROOT_PATH_likeKGN & filesBase(idx)
    End If
    TMN_SelectVzyskAndTemplate_likeKGN = True
End Function

' --- Подсчет и Предскан ---
Private Function TMN_CountBlocks_likeKGN(ByVal ws As Worksheet, ByVal only3plus As Boolean) As Long
    Dim r As Long, lastRow As Long, cellA As Range, blockRows As Long, total As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    r = 3: total = 0
    Do While r <= lastRow
        Set cellA = ws.Cells(r, "A")
        blockRows = 1
        If cellA.MergeCells Then
            If cellA.Address <> cellA.MergeArea.Cells(1, 1).Address Then r = r + 1: GoTo NextR
            blockRows = cellA.MergeArea.rows.count
        End If
        If (Not only3plus) Or (blockRows >= 3) Then total = total + 1
        r = r + blockRows
NextR:
    Loop
    TMN_CountBlocks_likeKGN = total
End Function

Private Function TMN_CountBlocksReady_likeKGN(ByVal ws As Worksheet, ByVal only3plus As Boolean) As Long
    Dim r As Long, lastRow As Long, cellA As Range, blockRows As Long, total As Long
    Dim recipient As String, rngInfo As Range
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    r = 3: total = 0
    Do While r <= lastRow
        Set cellA = ws.Cells(r, "A")
        blockRows = 1
        If cellA.MergeCells Then
            If cellA.Address <> cellA.MergeArea.Cells(1, 1).Address Then r = r + 1: GoTo NextR
            blockRows = cellA.MergeArea.rows.count
        End If
        If (Not only3plus) Or (blockRows >= 3) Then
            recipient = Trim(TMN_GetMergedTopValue_likeKGN(ws.Cells(r, "C")))
            Set rngInfo = ws.Cells(r, "D").MergeArea
            If Len(recipient) > 0 And (Not (rngInfo Is Nothing)) And LenB(CStr(rngInfo.Cells(1, 1).value)) > 0 Then
                total = total + 1
            End If
        End If
        r = r + blockRows
NextR:
    Loop
    TMN_CountBlocksReady_likeKGN = total
End Function

Private Sub TMN_PreScanMissing_likeKGN(ByVal ws As Worksheet, ByVal only3plus As Boolean, _
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
            If cellA.Address <> cellA.MergeArea.Cells(1, 1).Address Then r = r + 1: GoTo NextR
            blockRows = cellA.MergeArea.rows.count
        End If
        If (Not only3plus) Or (blockRows >= 3) Then
            recipient = Trim(TMN_GetMergedTopValue_likeKGN(ws.Cells(r, "C")))
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

Private Sub TMN_HighlightMissing_likeKGN(ByVal ws As Worksheet, ByVal rowsMissRec As Collection, ByVal rowsMissInfo As Collection)
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

' --- Лист результатов ---
Private Function TMN_CopyAsResultSheet_likeKGN(ByVal ws As Worksheet) As Worksheet
    Dim nm As String
    nm = TMN_UniqueSheetName_likeKGN("Результат формирования")
    ws.Copy After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.count)
    Set TMN_CopyAsResultSheet_likeKGN = ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.count)
    On Error Resume Next
    TMN_CopyAsResultSheet_likeKGN.name = nm
    On Error GoTo 0
End Function

Private Sub TMN_PrepareResultHeaders_likeKGN(ByVal wsRes As Worksheet)
    wsRes.Range("E2").value = "Результат"
    wsRes.Range("F2").value = "Описание"
    wsRes.Range("E2:F2").HorizontalAlignment = xlCenter
    wsRes.Range("E2:F2").VerticalAlignment = xlCenter
    wsRes.Columns("E").ColumnWidth = 16
    wsRes.Columns("F").ColumnWidth = 60
End Sub

Private Sub TMN_WriteResult_likeKGN(ByVal wsRes As Worksheet, ByVal topRow As Long, ByVal blockRows As Long, _
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

' --- Утилиты Excel/FS ---
Private Function TMN_PickFolder_likeKGN(Optional ByVal title As String = "Выбор папки") As String
    Dim fd As FileDialog
    On Error GoTo FAIL
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .title = title
        .AllowMultiSelect = False
        If .Show = -1 Then TMN_PickFolder_likeKGN = .SelectedItems(1)
    End With
    Exit Function
FAIL:
    TMN_PickFolder_likeKGN = ""
End Function

Private Function TMN_AppendPath_likeKGN(base As String, tail As String) As String
    If Right$(base, 1) = "\" Or Right$(base, 1) = "/" Then
        TMN_AppendPath_likeKGN = base & tail
    Else
        TMN_AppendPath_likeKGN = base & "\" & tail
    End If
End Function

Private Sub TMN_EnsureFolder_likeKGN(ByVal path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(path) Then fso.CreateFolder path
End Sub

Private Function TMN_SanitizeFileName_likeKGN(ByVal s As String) As String
    Dim badChars As Variant, ch As Variant
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In badChars
        s = Replace$(s, CStr(ch), "_")
    Next ch
    Do While Len(s) > 0 And (Right$(s, 1) = " " Or Right$(s, 1) = ".")
        s = Left$(s, Len(s) - 1)
    Loop
    If Len(s) = 0 Then s = "_"
    TMN_SanitizeFileName_likeKGN = s
End Function

Private Function TMN_UniqueSubfolderPath_likeKGN(ByVal baseFolder As String, ByVal folderName As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim basePath As String, n As Long, p As String
    If Right$(baseFolder, 1) = "\" Or Right$(baseFolder, 1) = "/" Then
        basePath = baseFolder
    Else
        basePath = baseFolder & "\"
    End If
    p = basePath & folderName
    If Not fso.FolderExists(p) Then TMN_UniqueSubfolderPath_likeKGN = p: Exit Function
    n = 1
    Do
        p = basePath & folderName & " - " & n
        If Not fso.FolderExists(p) Then TMN_UniqueSubfolderPath_likeKGN = p: Exit Function
        n = n + 1
    Loop
End Function

Private Function TMN_GetMergedTopValue_likeKGN(ByVal anyCell As Range) As String
    If anyCell.MergeCells Then
        TMN_GetMergedTopValue_likeKGN = CStr(anyCell.MergeArea.Cells(1, 1).value)
    Else
        TMN_GetMergedTopValue_likeKGN = CStr(anyCell.value)
    End If
End Function

Private Function TMN_SheetExists_likeKGN(nm As String) As Boolean
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ActiveWorkbook.Worksheets(nm)
    TMN_SheetExists_likeKGN = Not sh Is Nothing
    On Error GoTo 0
End Function

Private Function TMN_UniqueSheetName_likeKGN(baseName As String) As String
    Dim nm As String, i As Long
    nm = baseName: i = 2
    Do While TMN_SheetExists_likeKGN(nm)
        nm = baseName & "_" & i: i = i + 1
    Loop
    TMN_UniqueSheetName_likeKGN = nm
End Function

Private Function TMN_SafeFileExists_likeKGN(ByVal p As String) As Boolean
    On Error Resume Next
    TMN_SafeFileExists_likeKGN = CreateObject("Scripting.FileSystemObject").FileExists(p)
End Function


' --- Word Helpers ---

' Копирует шаблон в %AppData%, чтобы работать локально
Private Function TMN_ResolveTemplateForWord_likeKGN(ByVal netPath As String) As String
    On Error GoTo ErrHandler
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim appData As String
    appData = Environ$("AppData")
    If Right$(appData, 1) <> "\" Then appData = appData & "\"
    Dim templatesFolder As String
    templatesFolder = appData & "Microsoft\Templates\"
    If Not fso.FolderExists(templatesFolder) Then fso.CreateFolder templatesFolder

    If Len(Trim(netPath)) = 0 Or Not fso.FileExists(netPath) Then
        TMN_ResolveTemplateForWord_likeKGN = netPath
        Exit Function
    End If
    Dim localName As String, localPath As String
    localName = fso.GetFileName(netPath)
    localPath = templatesFolder & localName
    fso.CopyFile netPath, localPath, True
    TMN_ResolveTemplateForWord_likeKGN = localPath
    Exit Function
ErrHandler:
    TMN_ResolveTemplateForWord_likeKGN = netPath
End Function

Private Sub TMN_WordReplaceEverywhere_likeKGN(ByVal oDoc As Object, ByVal findText As String, ByVal replText As String)
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
                .Wrap = wdFindContinue_likeKGN
                .Execute Replace:=wdReplaceAll_likeKGN
            End With
            Set rng = rng.NextStoryRange
        Loop
    Next rngStory
End Sub

Private Function TMN_ReplacePlaceholderWithPaste_likeKGN(ByVal wdApp As Object, ByVal doc As Object, ByVal placeholder As String) As Boolean
    doc.Activate
    With wdApp.Selection
        .HomeKey wdStory_likeKGN
        With .Find
            .ClearFormatting: .Replacement.ClearFormatting
            .Text = placeholder
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue_likeKGN
            If .Execute Then
                Dim rng As Object
                Set rng = wdApp.Selection.Range
                rng.Text = ""
                wdApp.Selection.SetRange Start:=rng.End, End:=rng.End
                TMN_ReplacePlaceholderWithPaste_likeKGN = True
                Exit Function
            End If
        End With
    End With
    TMN_ReplacePlaceholderWithPaste_likeKGN = False
End Function

' Вставка таблицы 1x1 для Инфо
Private Sub TMN_InsertInfoSingleCell_likeKGN(ByVal wdApp As Object, ByVal infoText As String, ByVal doc As Object)
    On Error Resume Next
    Dim t As String
    t = Replace(infoText, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    t = Replace(t, vbLf, vbCr)

    Dim tbl As Object
    Set tbl = doc.Tables.Add(wdApp.Selection.Range, 1, 1)
    Dim targetW As Single: targetW = wdApp.CentimetersToPoints(17.5!)
    tbl.AllowAutoFit = False
    tbl.AutoFitBehavior 0 ' wdAutoFitFixed
    tbl.PreferredWidthType = wdPreferredWidthPoints_likeKGN
    tbl.PreferredWidth = targetW
    tbl.Columns(1).width = targetW
    tbl.rows.LeftIndent = 0
    tbl.LeftPadding = 0: tbl.RightPadding = 0
    With tbl.Borders
        .OutsideLineStyle = 1 ' wdLineStyleSingle
        .InsideLineStyle = 0
    End With
    With tbl.cell(1, 1).Range
        .Text = t
        .Font.name = "Times New Roman"
        .Font.Size = 12
        With .ParagraphFormat
            .Alignment = 0
            .LeftIndent = 0
            .RightIndent = 0
            .SpaceBefore = 0
            .SpaceAfter = 0
            .LineSpacingRule = 0
        End With
    End With
End Sub

' Создание общего файла
Private Sub TMN_CreateCombinedDocFromList_likeKGN(ByVal wdApp As Object, ByVal filesColl As Collection, ByVal destPath As String)
    On Error GoTo ErrHandler
    If filesColl.count <= 1 Then Exit Sub

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim i As Long
    Dim baseDoc As Object, srcDoc As Object, lastSection As Object

    Set baseDoc = wdApp.Documents.Open(fileName:=filesColl(1), ReadOnly:=False, AddToRecentFiles:=False, Visible:=False)

    For i = 2 To filesColl.count
        Set srcDoc = wdApp.Documents.Open(fileName:=filesColl(i), ReadOnly:=True, AddToRecentFiles:=False, Visible:=False)
        baseDoc.Activate
        With wdApp.Selection
            .EndKey Unit:=wdStory_likeKGN
        End With
        wdApp.Selection.InsertBreak Type:=wdSectionBreakNextPage_likeKGN
        wdApp.Selection.Range.FormattedText = srcDoc.Content.FormattedText

        Set lastSection = baseDoc.Sections(baseDoc.Sections.count)
        On Error Resume Next
        Dim hfType As Variant
        For Each hfType In Array(wdHeaderFooterPrimary_likeKGN, wdHeaderFooterFirstPage_likeKGN, wdHeaderFooterEvenPages_likeKGN)
            lastSection.headers(hfType).LinkToPrevious = False
            lastSection.Footers(hfType).LinkToPrevious = False
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

    baseDoc.SaveAs2 fileName:=destPath, FileFormat:=wdFormatXMLDocument_likeKGN
    baseDoc.Close SaveChanges:=False
    Set baseDoc = Nothing
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not baseDoc Is Nothing Then baseDoc.Close SaveChanges:=False
    If Not srcDoc Is Nothing Then srcDoc.Close SaveChanges:=False
End Sub

