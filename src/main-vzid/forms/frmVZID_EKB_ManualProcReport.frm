VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVZID_EKB_ManualProcReport 
   Caption         =   "РИЦ (Отчет ручной обработки)"
   ClientHeight    =   10410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13575
   OleObjectBlob   =   "frmVZID_EKB_ManualProcReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmVZID_EKB_ManualProcReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' === Параметры назначения ===
Private Const TARGET_FILE As String = "\\Ekb-vpfs01\екатеринбург\Отделы\Отдел взыскания задолженности по исполнительным документам\26. Макрос\ВЗИД\Отчет ручной обработки\Отчет ручной обработки РИЦ.xlsx"
Private Const TARGET_SHEET As String = "Sheet1"     ' куда пишем
Private Const HEADER_ROW As Long = 1                ' строка заголовков
Private Const DATA_START_ROW As Long = 2            ' первая строка данных

' --- Фиксированные колонки таблицы ---
Private Const COL_USER As Long = 1                  ' Пользователь
Private Const COL_YEAR As Long = 2                  ' Год
Private Const COL_MONTH As Long = 3                 ' Месяц (словом, например, "Сентябрь")
Private Const COL_DATE As Long = 4                  ' Дата (dd.mm.yyyy)
Private Const COL_TOTAL As Long = 5                 ' ИТОГО

' --- Динамические поля начинаются с 6-й колонки ---
Private Const START_COL As Long = 6                 ' собираем столбцы, начиная с 6

' === Динамически сформированные поля ===
Private colIdx() As Long      ' номер столбца в книге-назначении
Private txtNames() As String  ' имена TextBox на форме
Private headers() As String   ' тексты заголовков
Private fieldCount As Long    ' количество динамических полей (только колонки >= START_COL)

' === Элементы для отображения даты (из InputBox) ===
Private Const DATE_TXT_NAME As String = "txtDateDisplay"
Private Const DATE_LBL_NAME As String = "lblDateDisplay"
Private selectedDate As Date   ' выбранная дата (фиксируется до создания полей)

' === Для перезаписи при сохранении ===
Private existingRow As Long    ' если >0 — есть запись текущего пользователя на выбранную дату

' === Утилиты для блокировки (многопользовательский режим) ===
Private Function LockFilePath() As String
    Dim f As String
    f = TARGET_FILE
    LockFilePath = Left$(f, InStrRev(f, "\")) & "Отчет ручной обработки РИЦ.xlsx.lck"
End Function

Private Function AcquireLock(Optional ByVal timeoutSec As Long = 60) As Integer
    Dim started As Date: started = Now
    Dim fn As Integer
RetryLock:
    fn = FreeFile
    On Error Resume Next
    Open LockFilePath For Binary Access Read Write Shared As #fn
    If Err.Number <> 0 Then
        Err.Clear
        If DateDiff("s", started, Now) < timeoutSec Then
            DoEvents
            Application.Wait Now + TimeSerial(0, 0, 1)
            GoTo RetryLock
        Else
            AcquireLock = 0
            Exit Function
        End If
    End If
    Lock #fn
    Put #fn, , "LOCKED " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " by " & GetUserNameEx & vbCrLf
    AcquireLock = fn
End Function

Private Sub ReleaseLock(ByVal fn As Integer)
    On Error Resume Next
    If fn > 0 Then
        Unlock #fn
        Close #fn
        Kill LockFilePath
    End If
End Sub

Private Function GetUserNameEx() As String
    On Error Resume Next
    Dim u As String
    u = Environ$("USERNAME")
    If Len(Trim$(Application.userName)) > 0 Then
        GetUserNameEx = Application.userName
    ElseIf Len(u) > 0 Then
        GetUserNameEx = u
    Else
        GetUserNameEx = "UnknownUser"
    End If
End Function

' === Чтение заголовков из внешнего файла (только для инициализации формы) ===
Private Function ReadHeadersFromTarget(ByRef outHeaders() As String, ByRef outCols() As Long, ByRef lastCol As Long) As Boolean
    On Error GoTo FAIL
    Dim wb As Workbook, ws As Worksheet
    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(fileName:=TARGET_FILE, UpdateLinks:=False, ReadOnly:=True)
    Set ws = wb.Worksheets(TARGET_SHEET)
    lastCol = ws.Cells(HEADER_ROW, ws.Columns.count).End(xlToLeft).Column
    
    Dim c As Long, cnt As Long, h As String
    cnt = 0
    For c = START_COL To lastCol
        h = CStr(ws.Cells(HEADER_ROW, c).value)
        If KeepHeader(h) Then cnt = cnt + 1
    Next c
    If cnt = 0 Then GoTo FAIL
    
    ReDim outHeaders(1 To cnt)
    ReDim outCols(1 To cnt)
    
    Dim i As Long: i = 0
    For c = START_COL To lastCol
        h = CStr(ws.Cells(HEADER_ROW, c).value)
        If KeepHeader(h) Then
            i = i + 1
            outHeaders(i) = h
            outCols(i) = c
        End If
    Next c
    
    ReadHeadersFromTarget = True
CleanExit:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Exit Function
FAIL:
    ReadHeadersFromTarget = False
    Resume CleanExit
End Function

Private Sub cmdExport_Click()
    On Error GoTo FAIL

    Dim userName As String: userName = GetUserNameEx()

    ' --- куда сохранять (по умолчанию Рабочий стол, иначе — папка книги) ---
    Dim desktop As String, defaultName As String, saveTo As Variant
    desktop = Environ$("USERPROFILE") & "\Desktop"
    If Len(Dir$(desktop, vbDirectory)) = 0 Then desktop = ThisWorkbook.path
    defaultName = "Отчет ручной отработки_" & SanitizeFileName(userName) & "_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
    saveTo = Application.GetSaveAsFilename(InitialFileName:=desktop & "\" & defaultName, _
                                           FileFilter:="Excel Workbook (*.xlsx), *.xlsx")
    If VarType(saveTo) = vbBoolean And saveTo = False Then Exit Sub

    Application.ScreenUpdating = False

    ' --- откроем источник (read-only) ---
    Dim wbSrc As Workbook, wsSrc As Worksheet
    Set wbSrc = Workbooks.Open(fileName:=TARGET_FILE, UpdateLinks:=False, ReadOnly:=True)
    Set wsSrc = wbSrc.Worksheets(TARGET_SHEET)

    Dim lastRow As Long, lastCol As Long
    lastRow = wsSrc.Cells(wsSrc.rows.count, 1).End(xlUp).Row
    lastCol = wsSrc.Cells(HEADER_ROW, wsSrc.Columns.count).End(xlToLeft).Column

    ' --- создаём книгу-результат ---
    Dim wbOut As Workbook, wsOut As Worksheet
    Set wbOut = Workbooks.Add(xlWBATWorksheet)
    Set wsOut = wbOut.Worksheets(1)
    On Error Resume Next: wsOut.name = "Выгрузка": On Error GoTo 0

    ' заголовки
    wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, lastCol)).value = _
        wsSrc.Range(wsSrc.Cells(HEADER_ROW, 1), wsSrc.Cells(HEADER_ROW, lastCol)).value

    ' --- копируем строки текущего пользователя ---
    Dim r As Long, wr As Long: wr = 1
    Dim uSrc As String: uSrc = LCase$(Trim$(userName))

    For r = DATA_START_ROW To lastRow
        If LCase$(Trim$(CStr(wsSrc.Cells(r, COL_USER).value))) = uSrc Then
            wr = wr + 1
            wsOut.Range(wsOut.Cells(wr, 1), wsOut.Cells(wr, lastCol)).value = _
                wsSrc.Range(wsSrc.Cells(r, 1), wsSrc.Cells(r, lastCol)).value
        End If
    Next r

    ' нет строк?
    If wr = 1 Then
        wbSrc.Close SaveChanges:=False
        wbOut.Close SaveChanges:=False
        Application.ScreenUpdating = True
        MsgBox "Нет данных для выгрузки для пользователя: " & userName, vbInformation
        Exit Sub
    End If

    ' --- оформление ---
    wsOut.Columns(COL_DATE).NumberFormat = "dd.mm.yyyy"

    Dim used As Range
    Set used = wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(wr, lastCol))

    With used.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    used.HorizontalAlignment = xlCenter
    wsOut.Columns(COL_USER).HorizontalAlignment = xlLeft
    wsOut.Columns(COL_DATE).HorizontalAlignment = xlRight
    wsOut.rows(1).Font.Bold = True
    used.EntireColumn.AutoFit

    ' --- сохраняем ---
    wbOut.SaveAs fileName:=CStr(saveTo), FileFormat:=xlOpenXMLWorkbook
    wbSrc.Close SaveChanges:=False
    Application.ScreenUpdating = True

    MsgBox "Готово! Выгружено строк: " & (wr - 1) & vbCrLf & "Файл: " & CStr(saveTo), vbInformation
    Exit Sub

FAIL:
    Application.ScreenUpdating = True
    MsgBox "Не удалось выполнить выгрузку. " & Err.Description, vbExclamation
    Unload Me
End Sub

Private Function SanitizeFileName(ByVal s As String) As String
    Dim bad As Variant, b As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each b In bad
        s = Replace(s, CStr(b), "_")
    Next b
    SanitizeFileName = s
End Function


' === ИНИЦИАЛИЗАЦИЯ ФОРМЫ ===
Private Sub UserForm_Initialize()
    Dim s As String, dt As Date
    Do
        s = InputBox("Введите дату в формате ДД.ММ.ГГГГ", "Дата", Format$(Date, "dd.mm.yyyy"))
        If Len(s) = 0 Then
            Unload Me
            Exit Sub
        End If
        If TryParseDate(s, dt) Then Exit Do
        MsgBox "Введите корректную дату в формате ДД.ММ.ГГГГ.", vbExclamation
    Loop
    selectedDate = dt
    
    Dim lastCol As Long
    If Not ReadHeadersFromTarget(headers, colIdx, lastCol) Then
        MsgBox "Не удалось открыть целевой файл или прочитать заголовки:" & vbCrLf & TARGET_FILE, vbExclamation
        Unload Me
        Exit Sub
    End If
    
    fieldCount = UBound(headers)
    ReDim txtNames(1 To fieldCount)
    
    Dim y As Single, i As Long
    y = 12
    
    ' Дата (read-only)
    Dim lblD As MSForms.Label, tbD As MSForms.TextBox
    Set lblD = Me.Controls.Add("Forms.Label.1", DATE_LBL_NAME, True)
    With lblD: .Caption = "Дата": .Left = 12: .Top = y: .width = 500: End With
    Set tbD = Me.Controls.Add("Forms.TextBox.1", DATE_TXT_NAME, True)
    With tbD: .Left = 520: .Top = y - 2: .width = 120: .Text = Format$(selectedDate, "dd.mm.yyyy"): .Locked = True: .Enabled = False: .TabStop = False: End With
    y = y + 24
    
    ' Поля по заголовкам начиная с START_COL
    For i = 1 To fieldCount
        Dim lbl As MSForms.Label, tb As MSForms.TextBox
        Set lbl = Me.Controls.Add("Forms.Label.1", "lbl" & i, True)
        With lbl: .Caption = headers(i): .Left = 12: .Top = y: .width = 500: End With
        Set tb = Me.Controls.Add("Forms.TextBox.1", "txt" & i, True)
        With tb: .Left = 520: .Top = y - 2: .width = 120: .Tag = CStr(colIdx(i)): .TabIndex = i - 1: End With
        txtNames(i) = "txt" & i
        y = y + 24
    Next i
    
    Me.ScrollBars = fmScrollBarsVertical
    Me.ScrollHeight = y + 40
    On Error Resume Next
    Me.cmdSave.Top = y + 8
    Me.cmdClose.Top = y + 8
    Me.cmdClose.Left = Me.cmdSave.Left + Me.cmdSave.width + 8
    
    With Me.cmdExport
        .Caption = "Выгрузить в Excel"
        .Top = Me.cmdSave.Top
        .Left = Me.cmdClose.Left + Me.cmdClose.width + 8
        .Height = Me.cmdSave.Height
        .width = Me.cmdSave.width
        .ZOrder 0
    End With
    
    On Error GoTo 0
    
    PrefillForCurrentUserAndDate
End Sub

Private Sub PrefillForCurrentUserAndDate()
    existingRow = 0
    Dim wb As Workbook, ws As Worksheet, r As Long
    On Error GoTo CLEAN
    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(fileName:=TARGET_FILE, UpdateLinks:=False, ReadOnly:=True)
    Set ws = wb.Worksheets(TARGET_SHEET)
    
    r = FindExistingRowExact(ws, GetUserNameEx, selectedDate)
    existingRow = r
    
    If r > 0 Then
        Dim i As Long, c As Long, v As Variant
        For i = 1 To fieldCount
            c = colIdx(i)
            v = ws.Cells(r, c).value
            Me.Controls(txtNames(i)).Text = ValueToTextboxStr(v)
        Next i
    Else
        ClearDynamicInputs
    End If
    
CLEAN:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
End Sub

Private Sub ClearDynamicInputs()
    Dim i As Long
    For i = 1 To fieldCount
        Me.Controls(txtNames(i)).Text = ""
    Next i
End Sub

' === Сохранение ===
Private Sub cmdSave_Click()
    Dim dt As Date: dt = selectedDate
    
    Dim hLock As Integer: hLock = AcquireLock(60)
    If hLock = 0 Then
        MsgBox "Файл занят слишком долго. Повторите позже." & vbCrLf & TARGET_FILE, vbExclamation
        Exit Sub
    End If
    
    On Error GoTo SaveFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Dim wb As Workbook, ws As Worksheet, attempt As Long
    For attempt = 1 To 60
        On Error Resume Next
        Set wb = Workbooks.Open(fileName:=TARGET_FILE, UpdateLinks:=False, ReadOnly:=False, Notify:=False)
        If Err.Number <> 0 Then
            Err.Clear: DoEvents: Application.Wait Now + TimeSerial(0, 0, 1)
        ElseIf Not wb Is Nothing And wb.ReadOnly = False Then
            Exit For
        Else
            On Error Resume Next
            If Not wb Is Nothing Then wb.Close SaveChanges:=False
            Set wb = Nothing
            DoEvents: Application.Wait Now + TimeSerial(0, 0, 1)
        End If
    Next attempt
    If wb Is Nothing Or wb.ReadOnly Then GoTo SaveFail
    
    Set ws = wb.Worksheets(TARGET_SHEET)
    
    Dim r As Long: r = FindExistingRowExact(ws, GetUserNameEx, dt)
    Dim isUpdate As Boolean, targetRow As Long
    If r > 0 Then
        isUpdate = True: targetRow = r
    Else
        isUpdate = False
        Dim lastRow As Long: lastRow = LastUsedRow(ws)
        If lastRow < DATA_START_ROW - 1 Then lastRow = DATA_START_ROW - 1
        targetRow = lastRow + 1
    End If
    
    ' --- фиксированные колонки ---
    ws.Cells(targetRow, COL_USER).value = GetUserNameEx()
    ws.Cells(targetRow, COL_YEAR).value = Year(dt)
    ws.Cells(targetRow, COL_MONTH).value = MonthNameRu(Month(dt))
    ws.Cells(targetRow, COL_DATE).value = dt
    ws.Cells(targetRow, COL_DATE).NumberFormat = "dd.mm.yyyy"
    
    ' --- динамические поля + подсчёт ИТОГО ---
    Dim i As Long, tgtCol As Long, s As String, maxCol As Long
    Dim total As Double: total = 0
    maxCol = COL_TOTAL
    
    For i = 1 To fieldCount
        tgtCol = colIdx(i)
        If tgtCol > maxCol Then maxCol = tgtCol
        
        s = Me.Controls(txtNames(i)).Text
        If Len(Trim$(s)) = 0 Then
            ws.Cells(targetRow, tgtCol).value = 0
        Else
            ws.Cells(targetRow, tgtCol).value = s
        End If
        
        ' в сумму учитываем только числовые значения
        If IsNumeric(ws.Cells(targetRow, tgtCol).value) Then
            total = total + CDbl(ws.Cells(targetRow, tgtCol).value)
        End If
    Next i
    
    ' --- записать ИТОГО (колонка 5) ---
    ws.Cells(targetRow, COL_TOTAL).value = total
    
    ' --- оформление строки ---
    If maxCol < COL_TOTAL Then maxCol = COL_TOTAL
    ApplyRowFormatting ws, targetRow, maxCol
    
    ' --- сохранить и закрыть книгу ---
    wb.Save: wb.Close SaveChanges:=True
    
    ' --- Сообщение + вопрос про другую дату ---
    Dim prompt As String, resp As VbMsgBoxResult
    If isUpdate Then
        prompt = "Данные за " & Format$(dt, "dd.mm.yyyy") & " успешно перезаписаны!"
    Else
        prompt = "Данные за " & Format$(dt, "dd.mm.yyyy") & " успешно внесены!"
    End If
    prompt = prompt & vbCrLf & vbCrLf & "Хотите внести данные за другой день?"
    resp = MsgBox(prompt, vbQuestion + vbYesNo + vbDefaultButton2)
    
    Dim askAnother As Boolean: askAnother = False
    Dim shouldUnload As Boolean: shouldUnload = False
    Dim newDt As Date
    
    If resp = vbYes Then
        Dim sDate As String, tmp As Date
        Do
            sDate = InputBox("Введите дату в формате ДД.ММ.ГГГГ", "Дата", Format$(Date, "dd.mm.yyyy"))
            If Len(sDate) = 0 Then
                shouldUnload = True
                Exit Do
            End If
            If TryParseDate(sDate, tmp) Then
                newDt = tmp
                askAnother = True
                Exit Do
            Else
                MsgBox "Введите корректную дату в формате ДД.ММ.ГГГГ.", vbExclamation
            End If
        Loop
    Else
        shouldUnload = True
    End If
    
CleanExit:
    On Error Resume Next
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ReleaseLock hLock
    
    If askAnother Then
        selectedDate = newDt
        Me.Controls(DATE_TXT_NAME).Text = Format$(selectedDate, "dd.mm.yyyy")
        PrefillForCurrentUserAndDate
        Exit Sub
    End If
    
    If shouldUnload Then Unload Me
    Exit Sub
    
SaveFail:
    MsgBox "Не удалось записать данные. Возможно, файл занят или недоступен:" & vbCrLf & TARGET_FILE, vbExclamation
    Resume CleanExit
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub

' === ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ===
Private Function KeepHeader(ByVal h As String) As Boolean
    h = Trim$(CStr(h))
    KeepHeader = (Len(h) > 0)
End Function

Private Function TryParseDate(ByVal s As String, ByRef outDt As Date) As Boolean
    On Error GoTo bad
    s = Trim$(s)
    If Len(s) = 0 Then GoTo bad
    Dim dd As Integer, mm As Integer, yy As Integer
    Dim parts() As String: parts = Split(s, ".")
    If UBound(parts) <> 2 Then GoTo bad
    dd = CInt(parts(0)): mm = CInt(parts(1)): yy = CInt(parts(2))
    outDt = DateSerial(yy, mm, dd)
    TryParseDate = True
    Exit Function
bad:
    TryParseDate = False
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    On Error Resume Next
    Dim r As Long
    r = ws.Cells.Find(What:="*", After:=ws.Range("A1"), LookIn:=xlFormulas, LookAt:=xlPart, _
                      SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row
    If r <= 0 Then r = HEADER_ROW
    LastUsedRow = r
End Function

Private Function NzToString(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or VarType(v) = vbEmpty Then NzToString = "" Else NzToString = CStr(v)
End Function

Private Function ValueToTextboxStr(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or VarType(v) = vbEmpty Then Exit Function
    If IsNumeric(v) Then If CDbl(v) = 0 Then Exit Function
    Dim s As String: s = CStr(v)
    If Trim$(s) = "0" Then Exit Function
    ValueToTextboxStr = s
End Function

Private Function MonthNameRu(ByVal m As Long) As String
    Dim names As Variant
    names = Array("", "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", _
                       "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь")
    If m >= 1 And m <= 12 Then MonthNameRu = names(m)
End Function

Private Function CellsDateEqual(ByVal cell As Range, ByVal dt As Date) As Boolean
    On Error GoTo Fallback
    If IsDate(cell.value) Then
        CellsDateEqual = (CLng(DateValue(cell.value)) = CLng(DateValue(dt)))
        Exit Function
    End If
Fallback:
    On Error Resume Next
    CellsDateEqual = (Format$(cell.value, "dd.mm.yyyy") = Format$(dt, "dd.mm.yyyy"))
End Function

Private Function FindExistingRowExact(ByVal ws As Worksheet, ByVal userName As String, ByVal dt As Date) As Long
    Dim lastRow As Long: lastRow = LastUsedRow(ws)
    If lastRow < DATA_START_ROW Then Exit Function
    
    Dim u As String: u = LCase$(Trim$(userName))
    Dim r As Long
    For r = DATA_START_ROW To lastRow
        If LCase$(Trim$(CStr(ws.Cells(r, COL_USER).value))) = u Then
            If CellsDateEqual(ws.Cells(r, COL_DATE), dt) Then
                FindExistingRowExact = r
                Exit Function
            End If
        End If
    Next r
End Function

Private Sub ApplyRowFormatting(ByVal ws As Worksheet, ByVal r As Long, ByVal maxCol As Long)
    On Error Resume Next
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(r, 1), ws.Cells(r, maxCol))
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    rng.HorizontalAlignment = xlCenter
    ws.Cells(r, COL_USER).HorizontalAlignment = xlLeft
    ws.Cells(r, COL_DATE).HorizontalAlignment = xlRight
End Sub






