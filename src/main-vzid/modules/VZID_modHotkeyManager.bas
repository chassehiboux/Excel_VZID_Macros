Attribute VB_Name = "modHotkeyManager"
Option Explicit

' Один общий файл для всех макросов
Public Const HOTKEY_CONFIG_FILE As String = "MacroHotkeysConfig.txt"

' === Для автозапуска через ThisWorkbook ===
' Вставить в ThisWorkbook следующее – modHotkeyManager.Hotkey_LoadAll

Public Sub ReloadHotkeysDelayed()
    On Error Resume Next
    modHotkeyManager.Hotkey_LoadAll
End Sub


' === ВЫЗОВ ОКНА НАСТРОЙКИ ДЛЯ КОНКРЕТНОГО МАКРОСА ===
' Вызываешь так: Hotkey_ShowSetup "OspSelect_Run"
' или явно: Hotkey_ShowSetup "PERSONAL.XLSB!OspSelect_Run"
Public Sub Hotkey_ShowSetup(ByVal macroName As String)
    Dim fullName As String

    ' если нет "!" — считаем, что макрос в этом .xlam
    If InStr(1, macroName, "!", vbTextCompare) = 0 Then
        fullName = "'" & ThisWorkbook.name & "'!" & macroName
    Else
        fullName = macroName
    End If

    frmHotkey.InitForMacro fullName
End Sub


' === ЗАГРУЗКА ВСЕХ ГОРЯЧИХ КЛАВИШ ИЗ КОНФИГА ===
Public Sub Hotkey_LoadAll()
    Dim fPath As String
    Dim fnum As Integer
    Dim lineTxt As String
    Dim parts As Variant
    Dim macroName As String
    Dim keyString As String
    Dim cnt As Long

    fPath = GetConfigFolder() & "\" & HOTKEY_CONFIG_FILE
    If Dir(fPath, vbNormal) = "" Then
        MsgBox "Hotkey_LoadAll: файл конфига не найден: " & fPath, vbInformation
        Exit Sub
    End If

    fnum = FreeFile
    On Error GoTo ErrHandler
    Open fPath For Input As #fnum

    Do While Not EOF(fnum)
        Line Input #fnum, lineTxt
        lineTxt = Trim$(lineTxt)
        If lineTxt <> "" Then
            parts = Split(lineTxt, "|")
            If UBound(parts) >= 1 Then
                macroName = Trim$(parts(0))
                keyString = Trim$(parts(1))
                If macroName <> "" And keyString <> "" Then
                    Application.onKey keyString, macroName
                    cnt = cnt + 1
                End If
            End If
        End If
    Loop

    Close #fnum

    MsgBox "Hotkey_LoadAll: назначено горячих клавиш: " & cnt, vbInformation
    Exit Sub

ErrHandler:
    On Error Resume Next
    If fnum <> 0 Then Close #fnum
    MsgBox "Ошибка в Hotkey_LoadAll: " & Err.Number & " - " & Err.Description, vbCritical
End Sub


' === ПАПКА ДЛЯ КОНФИГА: %APPDATA%\Microsoft\Excel\LocalCache ===
Public Function GetConfigFolder() As String
    Dim basePath As String
    Dim cachePath As String

    basePath = Environ$("APPDATA") & "\Microsoft\Excel"
    cachePath = basePath & "\LocalCache"

    If Dir(basePath, vbDirectory) = "" Then
        On Error Resume Next
        MkDir basePath
        On Error GoTo 0
    End If

    If Dir(cachePath, vbDirectory) = "" Then
        On Error Resume Next
        MkDir cachePath
        On Error GoTo 0
    End If

    GetConfigFolder = cachePath
End Function

' === СОХРАНЕНИЕ ГОРЯЧЕЙ КЛАВИШИ ДЛЯ КОНКРЕТНОГО МАКРОСА ===
' В файле каждая строка: ИмяМакроса|OnKeyСтрока
Public Sub SaveHotkeyString(ByVal macroName As String, ByVal keyString As String)
    Dim fPath As String
    Dim tempPath As String
    Dim fnumIn As Integer
    Dim fnumOut As Integer
    Dim lineTxt As String
    Dim parts As Variant
    Dim nameInFile As String
    Dim found As Boolean

    fPath = GetConfigFolder() & "\" & HOTKEY_CONFIG_FILE
    tempPath = fPath & ".tmp"

    fnumOut = FreeFile
    On Error GoTo ErrHandler
    Open tempPath For Output As #fnumOut

    ' если файл есть — читаем и перезаписываем с заменой строки для macroName
    If Dir(fPath, vbNormal) <> "" Then
        fnumIn = FreeFile
        Open fPath For Input As #fnumIn

        Do While Not EOF(fnumIn)
            Line Input #fnumIn, lineTxt
            If Trim$(lineTxt) <> "" Then
                parts = Split(lineTxt, "|")
                If UBound(parts) >= 1 Then
                    nameInFile = Trim$(parts(0))
                    If StrComp(nameInFile, macroName, vbTextCompare) = 0 Then
                        ' перезаписываем строку для этого макроса
                        lineTxt = macroName & "|" & keyString
                        found = True
                    End If
                End If
                Print #fnumOut, lineTxt
            End If
        Loop

        Close #fnumIn
    End If

    ' если строки для этого макроса не было — добавим
    If Not found Then
        Print #fnumOut, macroName & "|" & keyString
    End If

    Close #fnumOut

    ' заменяем исходный файл
    On Error Resume Next
    Kill fPath
    Name tempPath As fPath
    On Error GoTo 0
    Exit Sub

ErrHandler:
    On Error Resume Next
    If fnumIn <> 0 Then Close #fnumIn
    If fnumOut <> 0 Then Close #fnumOut
    On Error Resume Next
    If Dir(tempPath, vbNormal) <> "" Then Kill tempPath
End Sub

' === ЗАГРУЗКА ГОРЯЧЕЙ КЛАВИШИ ДЛЯ КОНКРЕТНОГО МАКРОСА ===
Public Function LoadHotkeyString(ByVal macroName As String) As String
    Dim fPath As String
    Dim fnum As Integer
    Dim lineTxt As String
    Dim parts As Variant
    Dim nameInFile As String

    fPath = GetConfigFolder() & "\" & HOTKEY_CONFIG_FILE
    If Dir(fPath, vbNormal) = "" Then Exit Function

    fnum = FreeFile
    On Error GoTo ErrHandler
    Open fPath For Input As #fnum

    Do While Not EOF(fnum)
        Line Input #fnum, lineTxt
        If Trim$(lineTxt) <> "" Then
            parts = Split(lineTxt, "|")
            If UBound(parts) >= 1 Then
                nameInFile = Trim$(parts(0))
                If StrComp(nameInFile, macroName, vbTextCompare) = 0 Then
                    LoadHotkeyString = Trim$(parts(1))
                    Exit Do
                End If
            End If
        End If
    Loop

    Close #fnum
    Exit Function

ErrHandler:
    On Error Resume Next
    If fnum <> 0 Then Close #fnum
End Function

' === ПОИСК МАКРОСА ПО ГОРЯЧЕЙ КЛАВИШЕ ===
' Если такое сочетание уже есть в конфиге — вернёт имя макроса.
Public Function FindMacroByHotkey(ByVal keyString As String) As String
    Dim fPath As String
    Dim fnum As Integer
    Dim lineTxt As String
    Dim parts As Variant
    Dim macroName As String
    Dim keyInFile As String

    fPath = GetConfigFolder() & "\" & HOTKEY_CONFIG_FILE
    If Dir(fPath, vbNormal) = "" Then Exit Function

    fnum = FreeFile
    On Error GoTo ErrHandler
    Open fPath For Input As #fnum

    Do While Not EOF(fnum)
        Line Input #fnum, lineTxt
        If Trim$(lineTxt) <> "" Then
            parts = Split(lineTxt, "|")
            If UBound(parts) >= 1 Then
                macroName = Trim$(parts(0))
                keyInFile = Trim$(parts(1))
                If StrComp(keyInFile, keyString, vbTextCompare) = 0 Then
                    FindMacroByHotkey = macroName
                    Exit Do
                End If
            End If
        End If
    Loop

    Close #fnum
    Exit Function

ErrHandler:
    On Error Resume Next
    If fnum <> 0 Then Close #fnum
End Function

' === СПИСОК "ЗАРЕЗЕРВИРОВАННЫХ" СОЧЕТАНИЙ ===
' Здесь отмечаем то, что не хотим давать пользователю.
' Например, Ctrl+P (печать), Ctrl+S, Ctrl+O и т.п.
Public Function IsReservedHotkey(ByVal keyString As String) As Boolean
    Dim reserved As Variant
    Dim i As Long

    ' Список можно расширять по вкусу.
    reserved = Array( _
        "^P", _
        "^N", _
        "^O", _
        "^S", _
        "^W", _
        "^Q", _
        "^F", _
        "^H", _
        "^Z", _
        "^Y", _
        "^C", _
        "^V", _
        "^X" _
    )

    For i = LBound(reserved) To UBound(reserved)
        If StrComp(keyString, reserved(i), vbTextCompare) = 0 Then
            IsReservedHotkey = True
            Exit Function
        End If
    Next i
End Function

' === ПОСТРОЕНИЕ СТРОКИ ДЛЯ Application.OnKey ===
' Примеры:
'   Ctrl+Shift+O  -> "^+O"
'   Ctrl+Alt+F2   -> "^%{F2}"
Public Function BuildOnKeyString(ByVal KeyCode As Integer, ByVal Shift As Integer) As String
    Dim prefix As String
    Dim keyPart As String
    Dim fIndex As Integer

    prefix = ""
    If (Shift And 2) <> 0 Then prefix = prefix & "^"   ' Ctrl
    If (Shift And 1) <> 0 Then prefix = prefix & "+"   ' Shift
    If (Shift And 4) <> 0 Then prefix = prefix & "%"   ' Alt

    Select Case KeyCode
        Case vbKeyA To vbKeyZ
            keyPart = Chr$(KeyCode)                    ' буквы
        Case vbKey0 To vbKey9
            keyPart = Chr$(KeyCode)                    ' цифры
        Case vbKeyF1 To vbKeyF16
            fIndex = KeyCode - vbKeyF1 + 1
            keyPart = "{F" & CStr(fIndex) & "}"        ' F1..F16
        Case Else
            ' Неподдерживаемая клавиша
            BuildOnKeyString = ""
            Exit Function
    End Select

    BuildOnKeyString = prefix & keyPart
End Function

' === ЧЕЛОВЕЧЕСКОЕ ОТОБРАЖЕНИЕ СОЧЕТАНИЯ (для формы) ===
' Примеры:
'   Ctrl+Shift+O
'   Ctrl+Alt+F2
Public Function BuildDisplayString(ByVal KeyCode As Integer, ByVal Shift As Integer) As String
    Dim parts As String
    Dim fIndex As Integer

    parts = ""
    If (Shift And 2) <> 0 Then parts = parts & "Ctrl+"
    If (Shift And 1) <> 0 Then parts = parts & "Shift+"
    If (Shift And 4) <> 0 Then parts = parts & "Alt+"

    Select Case KeyCode
        Case vbKeyA To vbKeyZ
            parts = parts & UCase$(Chr$(KeyCode))
        Case vbKey0 To vbKey9
            parts = parts & Chr$(KeyCode)
        Case vbKeyF1 To vbKeyF16
            fIndex = KeyCode - vbKeyF1 + 1
            parts = parts & "F" & CStr(fIndex)
        Case Else
            parts = parts & "(неподдерживаемая клавиша)"
    End Select

    BuildDisplayString = parts
End Function


