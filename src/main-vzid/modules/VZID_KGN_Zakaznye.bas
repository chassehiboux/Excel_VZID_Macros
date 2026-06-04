Attribute VB_Name = "VZID_KGN_Zakaznye"
Option Explicit
Private Const TEMPLATE_PATH As String = "\\corp.vostok-electra.ru\Kgn\Отделы\Отдел взыскания по исполнительным документам\Зуйкевич Данил Иванович\Excel\ВЗИД\ЗАКАЗНЫЕ.xls"

Public Sub KGN_Zakaznye_Create()
    Dim ws As Worksheet: Set ws = ActiveWorkbook.Worksheets("Реестр пакетов")
    Dim creditor As String
    If Not KGN_SelectCreditor(creditor) Then Exit Sub

    ' --- ПРОХОД №1: проверка лимита 200 символов
    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    Dim r As Long, blockRows As Long, hasErr As Boolean
    Dim wsRes As Worksheet

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
                    Set wsRes = KGN_CopyAsResultSheet(ws)
                    KGN_PrepareResultHeaders wsRes
                End If
                KGN_WriteResult wsRes, r, blockRows, "ОШИБКА", "ORDERNUM > 200: " & ordLen & " символов"
            End If
        End If
        r = r + blockRows
nxtCheck:
    Loop

    If hasErr Then
        MsgBox "Превышено количество символов, смотри лист ""Результат формирования"".", vbCritical
        Exit Sub
    End If

    ' --- ПРОХОД №2: формирование файла
    Dim folder As String: folder = KGN_PickFolder("Куда сохранить реестр ""ЗАКАЗНЫЕ""?")
    If Len(folder) = 0 Then Exit Sub
    Dim outPath As String: outPath = KGN_UniquePath(folder, creditor & " " & Format(Date, "dd.mm.yyyy") & ".xls")

    Dim wbT As Workbook, wb As Workbook
    Application.ScreenUpdating = False
    Set wbT = Workbooks.Open(fileName:=TEMPLATE_PATH, ReadOnly:=True)
    wbT.SaveCopyAs outPath: wbT.Close False
    Set wb = Workbooks.Open(outPath)

    Dim wsData As Worksheet, cOrder As Long, cAddr As Long, cAdresat As Long, hdrRow As Long, tplRow As Long
    Set wsData = KGN_FindTemplateSheet(wb, cOrder, cAddr, cAdresat, hdrRow, tplRow)

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
            Dim recipient As String: recipient = Trim(KGN_GetMergedTopValue(ws.Cells(r, "C")))
            Dim addr As String: addr = Trim(KGN_GetMergedTopValue(ws.Cells(r, "B")))
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
End Sub

' --- Выбор взыскателя (Курган)
Private Function KGN_SelectCreditor(ByRef creditor As String) As Boolean
    Dim q As Variant
    q = Application.InputBox( _
        "Выберите взыскателя:" & vbCrLf & _
        "1 – АО ЭК Восток" & vbCrLf & _
        "2 – Водный союз" & vbCrLf & _
        "3 – КГК" & vbCrLf & _
        "4 – СКС" & vbCrLf & _
        "5 – Чистый город", "Курган", Type:=1)
    If VarType(q) = vbBoolean And q = False Then Exit Function
    Select Case CLng(q)
        Case 1: creditor = "АО ЭК Восток"
        Case 2: creditor = "Водный союз"
        Case 3: creditor = "КГК"
        Case 4: creditor = "СКС"
        Case 5: creditor = "Чистый город"
        Case Else: Exit Function
    End Select
    KGN_SelectCreditor = True
End Function

' --- Поиск листа и колонок шаблона
Private Function KGN_FindTemplateSheet(ByVal wb As Workbook, _
    ByRef cOrder As Long, ByRef cAddr As Long, ByRef cAdresat As Long, _
    ByRef hdrRow As Long, ByRef tplRow As Long) As Worksheet

    Dim sh As Worksheet, r As Long, c As Long
    For Each sh In wb.Worksheets
        For r = 1 To 50
            Dim fO As Long, fA As Long, fAD As Long
            For c = 1 To 30
                Select Case Trim$(UCase$(CStr(sh.Cells(r, c).value)))
                    Case "ORDERNUM": fO = c
                    Case "ADDRESSLINE": fA = c
                    Case "ADRESAT": fAD = c
                End Select
            Next c
            If fO > 0 And fA > 0 And fAD > 0 Then
                Set KGN_FindTemplateSheet = sh
                cOrder = fO: cAddr = fA: cAdresat = fAD
                hdrRow = r: tplRow = r + 1
                Exit Function
            End If
        Next r
    Next sh
    Err.Raise vbObjectError + 7301, , "В шаблоне не найдены заголовки ORDERNUM/ADDRESSLINE/ADRESAT."
End Function

' --- Утилиты
Private Function KGN_GetMergedTopValue(ByVal anyCell As Range) As String
    If anyCell.MergeCells Then KGN_GetMergedTopValue = CStr(anyCell.MergeArea.Cells(1, 1).value) _
    Else KGN_GetMergedTopValue = CStr(anyCell.value)
End Function

Private Function KGN_PickFolder(Optional ByVal title As String = "Выбор папки") As String
    Dim fd As FileDialog: Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd: .title = title: .AllowMultiSelect = False: If .Show = -1 Then KGN_PickFolder = .SelectedItems(1)
    End With
End Function

Private Function KGN_UniquePath(ByVal folder As String, ByVal fileName As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim base As String, ext As String, n As Long, p As String
    base = Left$(fileName, InStrRev(fileName, ".") - 1)
    ext = Mid$(fileName, InStrRev(fileName, "."))
    p = folder & IIf(Right$(folder, 1) = "\" Or Right$(folder, 1) = "/", "", "\") & fileName
    If Not fso.FileExists(p) Then KGN_UniquePath = p: Exit Function
    n = 1
    Do
        p = folder & IIf(Right$(folder, 1) = "\" Or Right$(folder, 1) = "/", "", "\") & base & " - " & n & ext
        If Not fso.FileExists(p) Then KGN_UniquePath = p: Exit Function
        n = n + 1
    Loop
End Function

Private Function KGN_CopyAsResultSheet(ByVal ws As Worksheet) As Worksheet
    Dim nm As String: nm = KGN_UniqueSheetName("Результат формирования")
    ws.Copy After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.count)
    Set KGN_CopyAsResultSheet = ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.count)
    On Error Resume Next: KGN_CopyAsResultSheet.name = nm: On Error GoTo 0
End Function

Private Sub KGN_PrepareResultHeaders(ByVal wsRes As Worksheet)
    wsRes.Range("E2").value = "Результат"
    wsRes.Range("F2").value = "Описание"
    wsRes.Columns("E").ColumnWidth = 16
    wsRes.Columns("F").ColumnWidth = 60
End Sub

Private Sub KGN_WriteResult(ByVal wsRes As Worksheet, ByVal topRow As Long, ByVal blockRows As Long, _
                            ByVal resultText As String, ByVal descr As String)
    Dim rngE As Range, rngF As Range
    Set rngE = wsRes.Range(wsRes.Cells(topRow, "E"), wsRes.Cells(topRow + blockRows - 1, "E"))
    Set rngF = wsRes.Range(wsRes.Cells(topRow, "F"), wsRes.Cells(topRow + blockRows - 1, "F"))
    On Error Resume Next: If rngE.MergeCells Then rngE.UnMerge: If rngF.MergeCells Then rngF.UnMerge: On Error GoTo 0
    rngE.Merge: rngF.Merge
    With rngE: .value = resultText: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True: End With
    With rngF: .value = descr: .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .WrapText = True: End With
End Sub

Private Function KGN_UniqueSheetName(baseName As String) As String
    Dim nm As String, i As Long: nm = baseName: i = 2
    Do While KGN_SheetExists(nm): nm = baseName & "_" & i: i = i + 1: Loop
    KGN_UniqueSheetName = nm
End Function

Private Function KGN_SheetExists(nm As String) As Boolean
    Dim sh As Worksheet: On Error Resume Next: Set sh = ActiveWorkbook.Worksheets(nm)
    KGN_SheetExists = Not sh Is Nothing: On Error GoTo 0
End Function


