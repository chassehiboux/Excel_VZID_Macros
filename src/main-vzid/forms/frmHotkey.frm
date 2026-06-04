VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHotkey 
   Caption         =   "Горячая клавиша"
   ClientHeight    =   2655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmHotkey.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHotkey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mOnKeyString As String
Private mMacroName As String

' Инициализация формы для конкретного макроса
Public Sub InitForMacro(ByVal macroName As String)
    Dim existing As String
    Dim shortName As String
    Dim pos As Long

    mMacroName = macroName
    mOnKeyString = ""

    ' вытащим часть после "!"
    shortName = macroName
    pos = InStr(1, shortName, "!", vbTextCompare)
    If pos > 0 Then shortName = Mid$(shortName, pos + 1)

    Me.Caption = "Горячая клавиша: " & shortName

    Me.lblInfo.Caption = _
        "Нажми сочетание клавиш в поле ниже (например, Ctrl+Shift+O)," & vbCrLf & _
        "затем нажми ОК. Это сочетание будет вызывать макрос:" & vbCrLf & _
        macroName

    existing = modHotkeyManager.LoadHotkeyString(mMacroName)

    If existing <> "" Then
        Me.lblCurrent.Caption = "Текущая горячая клавиша: " & existing
    Else
        Me.lblCurrent.Caption = "Текущая горячая клавиша: не задана"
    End If

    Me.txtHotkey.Text = ""
    Me.txtHotkey.SetFocus

    Me.Show
End Sub

Private Sub txtHotkey_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim onKey As String
    Dim disp As String

    onKey = modHotkeyManager.BuildOnKeyString(KeyCode, Shift)
    If onKey = "" Then
        mOnKeyString = ""
        Me.txtHotkey.Text = ""
    Else
        disp = modHotkeyManager.BuildDisplayString(KeyCode, Shift)
        mOnKeyString = onKey
        Me.txtHotkey.Text = disp
    End If

    ' чтобы в текстбоксе не рисовались реальные символы
    KeyCode = 0
End Sub

Private Sub cmdOK_Click()
    Dim usedBy As String

    If mMacroName = "" Then
        MsgBox "Не задано имя макроса.", vbCritical
        Exit Sub
    End If

    If mOnKeyString = "" Then
        MsgBox "Сначала нажми сочетание клавиш в поле.", vbExclamation
        Exit Sub
    End If

    ' 1) Проверяем, что сочетание не зарезервировано Excel (Ctrl+P и т.п.)
    If modHotkeyManager.IsReservedHotkey(mOnKeyString) Then
        MsgBox "Это сочетание уже занято встроенной командой Excel. Выбери другую горячую клавишу.", _
               vbExclamation
        Exit Sub
    End If

    ' 2) Проверяем, не назначено ли уже на другой макрос (по нашему конфигу)
    usedBy = modHotkeyManager.FindMacroByHotkey(mOnKeyString)
    If usedBy <> "" And StrComp(usedBy, mMacroName, vbTextCompare) <> 0 Then
        MsgBox "Это сочетание уже назначено на макрос '" & usedBy & "'. Выбери другую горячую клавишу.", _
               vbExclamation
        Exit Sub
    End If

    ' Всё ок — сохраняем и вешаем OnKey
    modHotkeyManager.SaveHotkeyString mMacroName, mOnKeyString
    Application.onKey mOnKeyString, mMacroName

    MsgBox "Для макроса '" & mMacroName & "' сохранена горячая клавиша: " & mOnKeyString, vbInformation
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


