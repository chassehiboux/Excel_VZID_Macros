VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVZID_TMN 
   Caption         =   "ТЮМЕНЬ"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5145
   OleObjectBlob   =   "frmVZID_TMN.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "frmVZID_TMN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
' ====== КОД ДЛЯ ДОСТУПА К КНОПКЕ ======
Private Function Allowed(users As String) As Boolean
    Allowed = InStr(";" & LCase$(users) & ";", ";" & LCase$(Environ$("Username")) & ";") > 0
End Function

' ====== ШАБЛОН ДЛЯ ДОСТУПА ======
' If Not Allowed("petrova") Then MsgBox "Нет доступа.", vbExclamation: Exit Sub

' ====== СЛУЖЕБНОЕ: обновление надписей версии/файла ======
Private Sub RefreshAboutVZID()
    On Error Resume Next
    lblVersion.Caption = "Версия (локальная): " & ReadLocalVersionVZID
End Sub

Private Sub CommandButton10_Click()
Unload Me
TMN_Make_Cover_Letters_likeKGN
End Sub

Private Sub CommandButton11_Click()
Unload Me
frmInstallBrowser.Show
End Sub

Private Sub CommandButton12_Click()
Unload Me
RunPDFScanner
End Sub

Private Sub CommandButton4_Click()
Dim p As String
    p = "\\corp.vostok-electra.ru\Tmn\Общая\ОВЗИД\Зуйкевич\ВЗИД\СОП\Порядок работы с макросом для формирования сопроводительных документов_Тюмень.docx"   ' <— ваш путь

    If Len(Dir(p)) = 0 Then
        MsgBox "Файл не найден: " & p, vbExclamation
        Exit Sub
    End If

    ThisWorkbook.FollowHyperlink p   ' откроется в Word
End Sub

Private Sub CommandButton5_Click()
Unload Me
TMN_Zakaznye_Create
End Sub

Private Sub CommandButton6_Click()
Unload Me
frmVZID_TMN_ManualProcReport.Show
End Sub

Private Sub CommandButton7_Click()
Unload Me
OspSelect_Run_TMN
End Sub

Private Sub CommandButton8_Click()
modHotkeyManager.Hotkey_ShowSetup "OspSelect_Run"
End Sub

Private Sub CommandButton9_Click()
Unload Me
TMN_Make_Cover_Letters_ByRecipient_likeKGN
End Sub

Private Sub UserForm_Initialize()
    RefreshAboutVZID
End Sub

Private Sub UserForm_Activate()
    RefreshAboutVZID
End Sub


Private Sub CommandButton1_Click()
    Unload Me
    Doc_Packets
End Sub

Private Sub CommandButton2_Click()
    Unload Me
    TMN_Make_Cover_Letters
End Sub

