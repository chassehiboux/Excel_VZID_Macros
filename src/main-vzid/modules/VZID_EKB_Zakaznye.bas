Attribute VB_Name = "VZID_EKB_Zakaznye"
Option Explicit
Private Const TEMPLATE_PATH As String = "\\Ekb-vpfs01\екатеринбург\Отделы\Отдел взыскания задолженности по исполнительным документам\26. Макрос\ВЗИД\ЗАКАЗНЫЕ.xls"

Public Sub RIC_Zakaznye_Create()
    ' --- безопасно получаем лист; если нет — показываем понятное сообщение и выходим
    Dim ws As Worksheet: Set ws = RIC_GetSheetOrError("Реестр пакетов")
    If ws Is Nothing Then Exit Sub

    Dim creditor As String
    If Not RIC_SelectCreditor(creditor) Then Exit Sub

    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    Dim r As Long, blockRows As Long, hasErr As Boolean
    Dim wsRes As Worksheet

    ' Проверка
    r = 3
    Do While r <= lastRow
        blockRows = 1
        If ws.Cells(r, "A").MergeCells Then
            If ws.Cells(r, "A").Address <> ws.Cells(r, "A").MergeArea.Cells(1, 1).Address Then r = r + 1: GoTo nxtCheck
            blockRows = ws.Cells(r, "A").MergeArea.rows.count
        End If
        Dim infoText As String: infoText = CStr(ws.Cells(r, "D").MergeArea.Cells(1, 1).value)
        If LenB(infoText) > 0 Then
            Dim ordLen As Long: ordLen = Len(creditor & vbLf & infoText)   ' перенос строки между взыскателем и Инфо
            If ordLen > 200 Then
                hasErr = True
                If wsRes Is Nothing Then
                    Set wsRes = RIC_CopyAsResultSheet(ws)
                    RIC_PrepareResultHeaders wsRes
                End If
                RIC_WriteResult wsRes, r, blockRows, "ОШИБКА", "COMMENT > 200: " & ordLen & " символов"
            End If
        End If
        r = r + blockRows
nxtCheck:
    Loop
    If hasErr Then
        MsgBox "Превышено количество символов, смотри лист ""Результат формирования"".", vbCritical
        Exit Sub
    End If

    ' Формирование
    Dim folder As String: folder = RIC_PickFolder("Куда сохранить реестр ""ЗАКАЗНЫЕ""?")
    If Len(folder) = 0 Then Exit Sub
    Dim outPath As String: outPath = RIC_UniquePath(folder, creditor & " " & Format(Date, "dd.mm.yyyy") & ".xls")

    Dim wbT As Workbook, wb As Workbook
    Application.ScreenUpdating = False
    Set wbT = Workbooks.Open(fileName:=TEMPLATE_PATH, ReadOnly:=True)
    wbT.SaveCopyAs outPath: wbT.Close False
    Set wb = Workbooks.Open(outPath)

    Dim wsData As Worksheet, cOrder As Long, cAddr As Long, cAdresat As Long, hdrRow As Long, tplRow As Long
    Set wsData = RIC_FindTemplateSheet(wb, cOrder, cAddr, cAdresat, hdrRow, tplRow)

    Dim outRow As Long: outRow = tplRow
    Dim first As Boolean: first = True
    r = 3
    Do While r <= lastRow
        blockRows = 1
        If ws.Cells(r, "A").MergeCells Then
            If ws.Cells(r, "A").Address <> ws.Cells(r, "A").MergeArea.Cells(1, 1).Address Then r = r + 1: GoTo nxtFill
            blockRows = ws.Cells(r, "A").MergeArea.rows.count
        End If
        Dim info As String: info = CStr(ws.Cells(r, "D").MergeArea.Cells(1, 1).value)
        If LenB(info) > 0 Then
            Dim recipient As String: recipient = Trim(RIC_GetMergedTopValue(ws.Cells(r, "C")))
            Dim addr As String: addr = Trim(RIC_GetMergedTopValue(ws.Cells(r, "B")))
            If first Then
                first = False                      ' первую пишем прямо в tplRow
            Else
                outRow = outRow + 1
                wsData.rows(outRow).Insert xlShiftDown
                wsData.rows(tplRow).Copy wsData.rows(outRow)
            End If

            wsData.Cells(outRow, cOrder).value = creditor & vbLf & info
            wsData.Cells(outRow, cOrder).WrapText = True
            ' при желании: wsData.Rows(outRow).EntireRow.AutoFit

            wsData.Cells(outRow, cAddr).value = addr
            wsData.Cells(outRow, cAdresat).value = recipient
        End If
        r = r + blockRows
nxtFill:
    Loop

    Application.CutCopyMode = False
    wb.Save: wb.Close False
    Application.ScreenUpdating = True
    MsgBox "Реестр сформирован:" & vbCrLf & outPath, vbInformation
    On Error Resume Next
CreateObject("WScript.Shell").Run "explorer.exe /select,""" & outPath & """", 1, False
On Error GoTo 0

End Sub

' === Новое: мягкая проверка наличия листа ===
Private Function RIC_GetSheetOrError(ByVal sheetName As String) As Worksheet
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ActiveWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If sh Is Nothing Then
        MsgBox "В активной книге отсутствует лист """ & sheetName & """." & vbCrLf & _
               "Создайте его или откройте корректный файл и повторите попытку.", _
               vbCritical, "РИЦ — ЗАКАЗНЫЕ"
    Else
        Set RIC_GetSheetOrError = sh
    End If
End Function

Private Function RIC_SelectCreditor(ByRef creditor As String) As Boolean
    Dim q As Variant
    Do
        q = Application.InputBox( _
            "Выберите взыскателя:" & vbCrLf & _
            "1 – РИФЕЙ" & vbCrLf & _
            "2 – ОКЭ/ОТСК", "РИЦ", Type:=1)
        If VarType(q) = vbBoolean And q = False Then Exit Function  ' Отмена
        If CLng(q) = 1 Then creditor = "РИФЕЙ": RIC_SelectCreditor = True: Exit Function
        If CLng(q) = 2 Then creditor = "ОКЭ/ОТСК": RIC_SelectCreditor = True: Exit Function
        MsgBox "Нужно ввести 1 или 2.", vbExclamation
    Loop
End Function

' --- Утилиты (префикс RIC)
Private Function RIC_FindTemplateSheet(ByVal wb As Workbook, _
    ByRef cOrder As Long, ByRef cAddr As Long, ByRef cAdresat As Long, _
    ByRef hdrRow As Long, ByRef tplRow As Long) As Worksheet
    Dim sh As Worksheet, r As Long, c As Long
    For Each sh In wb.Worksheets
        For r = 1 To 50
            Dim fO As Long, fA As Long, fAD As Long
            For c = 1 To 30
                Select Case Trim$(UCase$(CStr(sh.Cells(r, c).value)))
                    Case "COMMENT": fO = c
                    Case "ADDRESSLINE": fA = c
                    Case "ADRESAT": fAD = c
                End Select
            Next c
            If fO > 0 And fA > 0 And fAD > 0 Then
                Set RIC_FindTemplateSheet = sh
                cOrder = fO: cAddr = fA: cAdresat = fAD
                hdrRow = r: tplRow = r + 1
                Exit Function
            End If
        Next r
    Next sh
    Err.Raise vbObjectError + 7303, , "В шаблоне не найдены заголовки COMMENT/ADDRESSLINE/ADRESAT."
End Function

Private Function RIC_GetMergedTopValue(ByVal anyCell As Range) As String
    If anyCell.MergeCells Then RIC_GetMergedTopValue = CStr(anyCell.MergeArea.Cells(1, 1).value) _
    Else RIC_GetMergedTopValue = CStr(anyCell.value)
End Function

Private Function RIC_PickFolder(Optional ByVal title As String = "Выбор папки") As String
    Dim fd As FileDialog: Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd: .title = title: .AllowMultiSelect = False: If .Show = -1 Then RIC_PickFolder = .SelectedItems(1)
    End With
End Function

Private Function RIC_UniquePath(ByVal folder As String, ByVal fileName As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim base As String, ext As String, n As Long, p As String, sep As String

    ' 1) Санитизируем имя файла от недопустимых символов
    fileName = RIC_SafeFileName(fileName)

    sep = IIf(Right$(folder, 1) = "\" Or Right$(folder, 1) = "/", "", "\")
    base = Left$(fileName, InStrRev(fileName, ".") - 1)
    ext = Mid$(fileName, InStrRev(fileName, "."))

    ' 2) Гарантируем ограничение длины полного пути (218 символов для Excel)
    p = folder & sep & base & ext
    If Len(p) > 218 Then
        Dim allowLen As Long
        allowLen = 218 - Len(folder & sep) - Len(ext)
        If allowLen < 1 Then allowLen = 1
        base = Left$(base, allowLen)
        p = folder & sep & base & ext
    End If

    If Not fso.FileExists(p) Then RIC_UniquePath = p: Exit Function

    ' 3) Находим уникальное имя с суффиксом " - n" с учетом лимита пути
    n = 1
    Do
        Dim suffix As String: suffix = " - " & n
        Dim allow2 As Long
        allow2 = 218 - Len(folder & sep) - Len(ext) - Len(suffix)
        If allow2 < 1 Then allow2 = 1
        p = folder & sep & Left$(base, allow2) & suffix & ext
        If Not fso.FileExists(p) Then RIC_UniquePath = p: Exit Function
        n = n + 1
    Loop
End Function

Private Function RIC_CopyAsResultSheet(ByVal ws As Worksheet) As Worksheet
    Dim nm As String: nm = RIC_UniqueSheetName("Результат формирования")
    ws.Copy After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.count)
    Set RIC_CopyAsResultSheet = ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.count)
    On Error Resume Next: RIC_CopyAsResultSheet.name = nm: On Error GoTo 0
End Function

Private Sub RIC_PrepareResultHeaders(ByVal wsRes As Worksheet)
    wsRes.Range("E2").value = "Результат"
    wsRes.Range("F2").value = "Описание"
    wsRes.Columns("E").ColumnWidth = 16
    wsRes.Columns("F").ColumnWidth = 60
End Sub

' === Новое: санитизация имени файла ===
Private Function RIC_SafeFileName(ByVal rawName As String) As String
    Dim s As String: s = rawName
    ' Заменяем недопустимые в имени файла символы на дефис
    Dim bad As Variant: bad = Array("\", "/", ":", "*", "?", "\""", "<", ">", "|")
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        s = Replace$(s, CStr(bad(i)), "-")
    Next i
    ' Удаляем завершающие точки и пробелы
    Do While Len(s) > 0 And (Right$(s, 1) = "." Or Right$(s, 1) = " ")
        s = Left$(s, Len(s) - 1)
    Loop
    ' Пустое имя недопустимо — подставим _
    If LenB(s) = 0 Then s = "_"
    RIC_SafeFileName = s
End Function

Private Sub RIC_WriteResult(ByVal wsRes As Worksheet, ByVal topRow As Long, ByVal blockRows As Long, _
                            ByVal resultText As String, ByVal descr As String)
    Dim rngE As Range, rngF As Range
    Set rngE = wsRes.Range(wsRes.Cells(topRow, "E"), wsRes.Cells(topRow + blockRows - 1, "E"))
    Set rngF = wsRes.Range(wsRes.Cells(topRow, "F"), wsRes.Cells(topRow + blockRows - 1, "F"))
    On Error Resume Next: If rngE.MergeCells Then rngE.UnMerge: If rngF.MergeCells Then rngF.UnMerge: On Error GoTo 0
    rngE.Merge: rngF.Merge
    With rngE: .value = resultText: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True: End With
    With rngF: .value = descr: .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .WrapText = True: End With
End Sub

Private Function RIC_UniqueSheetName(baseName As String) As String
    Dim nm As String, i As Long: nm = baseName: i = 2
    Do While RIC_SheetExists(nm): nm = baseName & "_" & i: i = i + 1: Loop
    RIC_UniqueSheetName = nm
End Function

Private Function RIC_SheetExists(nm As String) As Boolean
    Dim sh As Worksheet: On Error Resume Next: Set sh = ActiveWorkbook.Worksheets(nm)
    RIC_SheetExists = Not sh Is Nothing: On Error GoTo 0
End Function


