Attribute VB_Name = "DocPacks"
Option Explicit

Sub Doc_Packets()
    Dim wsSrc As Worksheet, wsDest As Worksheet
    Dim concatCols As Range, opfrCol As Range
    Dim rowsPerBlk As Long, shortenFIO As Boolean
    Dim lastRow As Long, rDest As Long, pktNum As Long
    Dim sep As String: sep = " | "
    Dim c As Range
    
    Set wsSrc = ActiveSheet
    
    '=== 1. Размер пакета ==================================================
    rowsPerBlk = Application.InputBox( _
        prompt:="Сколько строк объединять в один пакет?", _
        title:="Размер пакета", Type:=1)
    If rowsPerBlk < 1 Then
        MsgBox "Отменено или введено недопустимое число.", vbExclamation
        Exit Sub
    End If
    
    '=== 2. Выбор столбцов -------------------------------------------------
    On Error Resume Next
    Set concatCols = Application.InputBox( _
        prompt:="Выделите ВСЕ столбцы, текст которых нужно сцепить (можно удерживать Ctrl).", _
        title:="Столбцы для объединения", Type:=8)
    If concatCols Is Nothing Then Exit Sub
    
    Set opfrCol = Application.InputBox( _
        prompt:="Выделите столбец с «ОПФР/ОСП».", _
        title:="Столбец ОПФР/ОСП", Type:=8)
    If opfrCol Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Dim opfrColNum As Long
    opfrColNum = opfrCol.Columns(1).Column
    
    '=== 3. Определяем, есть ли «ФИО» среди выбранных ---------------------
    Dim fioColNum As Long: fioColNum = 0
    For Each c In concatCols.Columns
        If Trim(UCase(wsSrc.Cells(1, c.Column).value)) = "ФИО" Then
            fioColNum = c.Column
            Exit For
        End If
    Next c
    
    If fioColNum <> 0 Then
        If MsgBox("Сократить ФИО до инициалов (Фамилия И.О.)? Если в ячейка ФИО была записана со старой фамилией в (скобках), то она будет удалена", _
                  vbYesNo + vbQuestion, "ФИО") = vbYes Then
            shortenFIO = True
        End If
    End If
    
    '=== 4. Создаём/пересоздаём лист «Реестр пакетов» ---------------------
    Application.ScreenUpdating = False
    On Error Resume Next: Worksheets("Реестр пакетов").Delete: On Error GoTo 0
    Set wsDest = Worksheets.Add(After:=Worksheets(Worksheets.count))
    wsDest.name = "Реестр пакетов"
    
    With wsDest.Range("A1:D1")
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    wsDest.Range("A2:D2").value = Array("Номер", "Адрес", "Получатель", "Инфо")
    wsDest.Range("A2:D2").HorizontalAlignment = xlCenter
    wsDest.Range("A2:D2").VerticalAlignment = xlCenter
    
    '=== 5. Готовим данные: группируем все строки по ОПФР/ОСП -------------
    lastRow = wsSrc.Cells(wsSrc.rows.count, opfrColNum).End(xlUp).Row
    If lastRow < 2 Then
        Application.ScreenUpdating = True
        MsgBox "Нет данных.", vbExclamation
        Exit Sub
    End If
    
    Dim dict As Object, dictDisp As Object
    Dim keysOrder As Collection
    Set dict = CreateObject("Scripting.Dictionary")    ' key -> Collection of row numbers
    Set dictDisp = CreateObject("Scripting.Dictionary") ' key -> отображаемый текст
    Set keysOrder = New Collection
    
    Dim r As Long, key As String, disp As String
    For r = 2 To lastRow
        disp = CStr(wsSrc.Cells(r, opfrColNum).value)
        key = NormalizeKey(disp) ' нормализуем
        If Not dict.exists(key) Then
            dict.Add key, New Collection
            dictDisp.Add key, disp
            keysOrder.Add key
        End If
        dict(key).Add r
    Next r

    '=== 5a. Загружаем справочник адресов из ОСП.xlsx ======================
    Dim addrMap As Object
    Set addrMap = LoadOSPAddressMap() ' может вернуть Nothing, тогда просто не заполняем адреса
    
    '=== 6. Выгрузка: сперва полные пакеты, затем неполные ----------------
    rDest = 3
    pktNum = 1
    
    Dim k As Variant
    Dim rowsColl As Collection
    Dim total As Long, fullCnt As Long, startIdx As Long, i As Long, rowIdx As Long
    Dim blk As Long, txt As String, lineTxt As String, val As String
    Dim addr As String
    
    '--- 6.1 Полные пакеты ---
    For Each k In keysOrder
        Set rowsColl = dict(k)
        total = rowsColl.count
        If total >= rowsPerBlk Then
            fullCnt = total \ rowsPerBlk
            For i = 0 To fullCnt - 1
                blk = rowsPerBlk
                txt = ""
                Dim j As Long, t As Long
                For j = 1 To blk
                    rowIdx = rowsColl(i * rowsPerBlk + j)
                    lineTxt = ""
                    For Each c In concatCols.Columns
                        val = CStr(wsSrc.Cells(rowIdx, c.Column).value)
                        If shortenFIO And c.Column = fioColNum Then
                            val = ToInitials(val)
                        End If
                        lineTxt = lineTxt & val & sep
                    Next c
                    lineTxt = Left(lineTxt, Len(lineTxt) - Len(sep))
                    txt = txt & lineTxt & vbLf
                Next j
                txt = Left(txt, Len(txt) - 1)
                
                ' Адрес по ОПФР/ОСП (если найден)
                addr = ""
                If Not addrMap Is Nothing Then
                    If addrMap.exists(k) Then addr = addrMap(k)
                End If
                
                ' Запись пакета
                With wsDest.Range(wsDest.Cells(rDest, "A"), wsDest.Cells(rDest + blk - 1, "A"))
                    .Merge: .value = pktNum
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With
                With wsDest.Range(wsDest.Cells(rDest, "B"), wsDest.Cells(rDest + blk - 1, "B"))
                    .Merge: .value = addr
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .WrapText = True
                End With
                With wsDest.Range(wsDest.Cells(rDest, "C"), wsDest.Cells(rDest + blk - 1, "C"))
                    .Merge: .value = dictDisp(k)
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .WrapText = True
                End With
                With wsDest.Range(wsDest.Cells(rDest, "D"), wsDest.Cells(rDest + blk - 1, "D"))
                    .Merge: .value = txt
                    .HorizontalAlignment = xlLeft
                    .VerticalAlignment = xlCenter
                    .WrapText = True
                End With
                
                rDest = rDest + blk
                pktNum = pktNum + 1
            Next i
        End If
    Next k
    
    '--- 6.2 Неполные пакеты (остатки) ---
    For Each k In keysOrder
        Set rowsColl = dict(k)
        total = rowsColl.count
        blk = total Mod rowsPerBlk
        If blk > 0 Then
            startIdx = total - blk + 1
            txt = ""
            For i = startIdx To total
                rowIdx = rowsColl(i)
                lineTxt = ""
                For Each c In concatCols.Columns
                    val = CStr(wsSrc.Cells(rowIdx, c.Column).value)
                    If shortenFIO And c.Column = fioColNum Then
                        val = ToInitials(val)
                    End If
                    lineTxt = lineTxt & val & sep
                Next c
                lineTxt = Left(lineTxt, Len(lineTxt) - Len(sep))
                txt = txt & lineTxt & vbLf
            Next i
            txt = Left(txt, Len(txt) - 1)
            
            ' Адрес по ОПФР/ОСП (если найден)
            addr = ""
            If Not addrMap Is Nothing Then
                If addrMap.exists(k) Then addr = addrMap(k)
            End If
            
            With wsDest.Range(wsDest.Cells(rDest, "A"), wsDest.Cells(rDest + blk - 1, "A"))
                .Merge: .value = pktNum
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            With wsDest.Range(wsDest.Cells(rDest, "B"), wsDest.Cells(rDest + blk - 1, "B"))
                .Merge: .value = addr
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = True
            End With
            With wsDest.Range(wsDest.Cells(rDest, "C"), wsDest.Cells(rDest + blk - 1, "C"))
                .Merge: .value = dictDisp(k)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = True
            End With
            With wsDest.Range(wsDest.Cells(rDest, "D"), wsDest.Cells(rDest + blk - 1, "D"))
                .Merge: .value = txt
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
                .WrapText = True
            End With
            
            rDest = rDest + blk
            pktNum = pktNum + 1
        End If
    Next k
    
    '=== 7. Форматирование -------------------------------------------------
    wsDest.Cells.Font.Size = 8
    wsDest.Columns("C").ColumnWidth = 20
    wsDest.Columns("B").ColumnWidth = wsDest.Columns("C").ColumnWidth
    wsDest.Columns("D").ColumnWidth = 150
    wsDest.Columns("A").AutoFit
    
    With wsDest.Range("A1:D" & wsDest.Cells(wsDest.rows.count, "A").End(xlUp).Row).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    Application.ScreenUpdating = True
    MsgBox "Реестр пакетов готов!"
End Sub

'=== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ =============================================

' Убирает любые круглые скобки и их содержимое, схлопывает пробелы
Private Function StripParens(ByVal s As String) As String
    Dim i As Long, depth As Long, ch As String, res As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case ch
            Case "("
                depth = depth + 1
            Case ")"
                If depth > 0 Then depth = depth - 1
            Case Else
                If depth = 0 Then res = res & ch
        End Select
    Next i
    res = Replace(res, vbTab, " ")
    Do While InStr(res, "  ") > 0
        res = Replace(res, "  ", " ")
    Loop
    StripParens = Trim$(res)
End Function

' Берёт первую «осмысленную» букву токена (пропуская точки/кавычки/дефисы и т.п.)
Private Function InitialFromToken(ByVal s As String) As String
    Dim i As Long, ch As String
    s = Trim$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch <> "." And ch <> "-" And ch <> "(" And ch <> ")" And ch <> "«" And ch <> "»" And ch <> """" Then
            InitialFromToken = ch
            Exit Function
        End If
    Next i
End Function

' === Обновлено: учитывает скобки и «мусор» между словами ===
Private Function ToInitials(fullName As String) As String
    Dim CLEAN As String
    Dim parts() As String
    Dim surname As String, firstInit As String, secondInit As String
    Dim i As Long, init As String
    
    CLEAN = StripParens(Trim$(fullName))
    If Len(CLEAN) = 0 Then ToInitials = "": Exit Function
    
    parts = Split(CLEAN, " ")
    surname = parts(0)
    
    ' Ищем первые две осмысленные буквы после фамилии
    For i = 1 To UBound(parts)
        init = InitialFromToken(parts(i))
        If Len(init) > 0 Then
            If firstInit = "" Then
                firstInit = init
            ElseIf secondInit = "" Then
                secondInit = init
                Exit For
            End If
        End If
    Next i
    
    ToInitials = surname
    If Len(firstInit) > 0 Then ToInitials = ToInitials & " " & firstInit & "."
    If Len(secondInit) > 0 Then ToInitials = ToInitials & secondInit & "."
End Function


Private Function NormalizeKey(ByVal s As String) As String
    NormalizeKey = UCase$(Trim$(CStr(s)))
End Function

Private Function FileExists(ByVal fullPath As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir(fullPath, vbNormal)) > 0)
    On Error GoTo 0
End Function

Private Function FindHeaderCell(ws As Worksheet, ByVal headerText As String) As Range
    Dim rng As Range
    On Error Resume Next
    Set rng = ws.UsedRange.Find(What:=headerText, LookAt:=xlWhole, LookIn:=xlValues, MatchCase:=False)
    On Error GoTo 0
    Set FindHeaderCell = rng
End Function

Private Function LoadOSPAddressMap() As Object
    ' Возвращает Dictionary: ключ = NormalizeKey(ОПФР/ОСП), значение = Адрес
    Dim primaryPath As String, fallbackPath As String, fallback2Path As String, chosenPath As String
    primaryPath = "\\corp.vostok-electra.ru\Kgn\Отделы\Отдел взыскания по исполнительным документам\Зуйкевич Данил Иванович\Excel\ВЗИД\ОСП.xlsx"
    fallbackPath = "\\corp.vostok-electra.ru\Tmn\Общая\ОВЗИД\Зуйкевич\ВЗИД\ОСП.xlsx"
    fallback2Path = "\\Ekb-vpfs01\екатеринбург\Отделы\Отдел взыскания задолженности по исполнительным документам\26. Макрос\ВЗИД\ОСП.xlsx"
    
    If FileExists(primaryPath) Then
        chosenPath = primaryPath
    ElseIf FileExists(fallbackPath) Then
        chosenPath = fallbackPath
    ElseIf FileExists(fallback2Path) Then
        chosenPath = fallback2Path
    Else
        ' Не нашли файл ни по одному пути — возврат Nothing
        Exit Function
    End If
    
    Dim wb As Workbook, ws As Worksheet, wsData As Worksheet
    Dim weOpened As Boolean
    Dim cellOP As Range, cellAddr As Range
    Dim opCol As Long, addrCol As Long, hdrRow As Long
    
    ' Пытаемся использовать уже открытую книгу
    On Error Resume Next
    Set wb = Workbooks("ОСП.xlsx")
    On Error GoTo 0
    
    If wb Is Nothing Then
        On Error Resume Next
        Set wb = Workbooks.Open(fileName:=chosenPath, ReadOnly:=True, UpdateLinks:=False)
        If Err.Number <> 0 Then
            Set wb = Nothing
            Exit Function
        End If
        On Error GoTo 0
        weOpened = True
    End If
    
    ' Ищем лист, где есть оба заголовка
    Set wsData = Nothing
    For Each ws In wb.Worksheets
        Set cellOP = FindHeaderCell(ws, "ОПФР/ОСП")
        Set cellAddr = FindHeaderCell(ws, "Адрес")
        If Not cellOP Is Nothing And Not cellAddr Is Nothing Then
            Set wsData = ws
            Exit For
        End If
    Next ws
    
    If wsData Is Nothing Then GoTo CLEANUP
    
    opCol = cellOP.Column
    addrCol = cellAddr.Column
    hdrRow = cellOP.Row  ' считаем строку заголовков по позиции "ОПФР/ОСП"
    
    Dim lastRow As Long, r As Long, k As String, a As String
    lastRow = wsData.Cells(wsData.rows.count, opCol).End(xlUp).Row
    
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    
    For r = hdrRow + 1 To lastRow
        k = NormalizeKey(wsData.Cells(r, opCol).value)
        If Len(k) > 0 Then
            a = CStr(wsData.Cells(r, addrCol).value)
            If Not d.exists(k) Then d.Add k, a
        End If
    Next r
    
    Set LoadOSPAddressMap = d

CLEANUP:
    If weOpened And Not wb Is Nothing Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If
End Function


