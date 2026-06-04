Attribute VB_Name = "OspSelector"
Option Explicit

' ПУТИ К ФАЙЛАМ ОСП ПО РЕГИОНАМ
Public Const PATH_KGN As String = "\\corp.vostok-electra.ru\Kgn\Отделы\Отдел взыскания по исполнительным документам\Зуйкевич Данил Иванович\Excel\ВЗИД\ОСПСелектор.xlsx"
Public Const PATH_TMN As String = "\\corp.vostok-electra.ru\Tmn\Общая\ОВЗИД\Зуйкевич\ВЗИД\ОСПСелектор.xlsx"
Public Const PATH_RIC As String = "\\Ekb-vpfs01\екатеринбург\Отделы\Отдел взыскания задолженности по исполнительным документам\26. Макрос\ВЗИД\ОСПСелектор.xlsx"
Public Const PATH_CHLB As String = "\\corp.vostok-electra.ru\Kgn\Отделы\Отдел взыскания по исполнительным документам\Зуйкевич Данил Иванович\Excel\ВЗИД\ОСПСелектор.xlsx"

' НАСТРОЙКИ ФИЛЬТРОВ РЕГИОНОВ
' Если нужно несколько значений — пиши через ;, например:
' "Курган;Курганская область"
Public Const REGION_KGN As String = "Курган"
Public Const REGION_TMN As String = "Тюмень"
Public Const REGION_RIC As String = "РИЦ"
Public Const REGION_CHLB As String = "ЧЭС"

' НОМЕРА СТОЛБЦОВ В ОСП.xlsx
Public Const COL_ADDR As Long = 1      ' Адрес
Public Const COL_OSP As Long = 2       ' ОСП
Public Const COL_REGION As Long = 3    ' Регион

' Структура для одной строки справочника
Public Type TOspItem
    OspName As String
    Address As String
    Region As String
End Type

' Глобальные переменные, к которым обращается форма
Public OspList() As TOspItem   ' массив ОСП
Public OspCount As Long        ' сколько элементов в массиве
Public TargetCell As Range     ' выбранная пользователем ячейка

' ========= ПУБЛИЧНЫЕ ВХОДЫ ДЛЯ РАЗНЫХ ТЕРРИТОРИЙ =========

Public Sub OspSelect_Run_KGN()
    OspSelect_RunCore PATH_KGN, REGION_KGN
End Sub

Public Sub OspSelect_Run_TMN()
    OspSelect_RunCore PATH_TMN, REGION_TMN
End Sub

Public Sub OspSelect_Run_RIC()
    OspSelect_RunCore PATH_RIC, REGION_RIC
End Sub

Public Sub OspSelect_Run_CHLB()
    OspSelect_RunCore PATH_CHLB, REGION_CHLB
End Sub

' Если где-то уже используется старый OspSelect_Run — можно сделать алиас.
Public Sub OspSelect_Run()
    ' По желанию: какой регион считать "по умолчанию"
    OspSelect_Run_TMN
End Sub

' ========= ОБЩАЯ ЛОГИКА ВЫБОРА =========

Private Sub OspSelect_RunCore(ByVal ospPath As String, ByVal regionFilter As String)
    Dim wbOsp As Workbook
    Dim ok As Boolean

    ' проверяем, что есть выбранная ячейка
    If TypeName(Selection) <> "Range" Then
        MsgBox "Сначала выбери ячейку, куда вставлять ОСП.", vbExclamation
        Exit Sub
    End If

    Set TargetCell = Selection(1)

    ' открываем конкретный файл ОСП
    Set wbOsp = OpenOspWorkbook(ospPath)
    If wbOsp Is Nothing Then
        MsgBox "Не удалось открыть файл ОСП:" & vbCrLf & ospPath, vbCritical
        Exit Sub
    End If

    ' читаем данные из файла с учётом фильтра региона
    ok = LoadOspData(wbOsp, regionFilter)
    wbOsp.Close SaveChanges:=False
    Set wbOsp = Nothing

    If Not ok Or OspCount = 0 Then
        MsgBox "Подходящих ОСП (по фильтру региона """ & regionFilter & """) не найдено.", vbExclamation
        Exit Sub
    End If

    ' показываем форму выбора
    frmVZID_OspSelector.Show
End Sub

' Открываем конкретный файл
Private Function OpenOspWorkbook(ByVal ospPath As String) As Workbook
    Dim wb As Workbook

    On Error Resume Next
    Set wb = Workbooks.Open(fileName:=ospPath, ReadOnly:=True)
    On Error GoTo 0

    If Not wb Is Nothing Then
        Set OpenOspWorkbook = wb
    End If
End Function

' Читаем ОСП из книги, фильтруем по региону
Public Function LoadOspData(wb As Workbook, ByVal regionFilter As String) As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim regionVal As String

    ' Если лист в OСП.xlsx называется иначе — поменяй имя/индекс.
    Set ws = wb.Worksheets(1)

    lastRow = ws.Cells(ws.rows.count, COL_REGION).End(xlUp).Row
    If lastRow < 2 Then Exit Function   ' только заголовки или пусто

    ReDim OspList(1 To lastRow - 1)
    OspCount = 0

    For i = 2 To lastRow   ' предполагаем, что заголовки в первой строке
        regionVal = CStr(ws.Cells(i, COL_REGION).value)

        If RegionPasses(regionVal, regionFilter) Then
            OspCount = OspCount + 1
            With OspList(OspCount)
                .OspName = CStr(ws.Cells(i, COL_OSP).value)
                .Address = CStr(ws.Cells(i, COL_ADDR).value)
                .Region = regionVal
            End With
        End If
    Next i

    If OspCount > 0 Then
        ReDim Preserve OspList(1 To OspCount)
        LoadOspData = True
    End If
End Function

' Проверка: подходит ли строка по региону (можно несколько через ;)
Public Function RegionPasses(ByVal regionVal As String, ByVal regionFilter As String) As Boolean
    Dim filters As Variant
    Dim i As Long
    Dim f As String

    ' Если фильтр пустой — берём все регионы
    If Trim$(regionFilter) = "" Then
        RegionPasses = True
        Exit Function
    End If

    ' Несколько регионов разделяем по ";"
    filters = Split(regionFilter, ";")

    For i = LBound(filters) To UBound(filters)
        f = Trim$(filters(i))
        If f <> "" Then
            If StrComp(Trim$(regionVal), f, vbTextCompare) = 0 Then
                RegionPasses = True
                Exit Function
            End If
        End If
    Next i
End Function

