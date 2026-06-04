VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVZID_KGN 
   Caption         =   "КУРГАН"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5145
   OleObjectBlob   =   "frmVZID_KGN.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmVZID_KGN"
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
frmInstallBrowser.Show
End Sub

Private Sub CommandButton11_Click()
Unload Me
RunPDFScanner
End Sub

Private Sub PDFSigner_Click()
Unload Me
LaunchPDFSigner
End Sub

Private Sub CommandButton7_Click()
If Not Allowed("skrasilova;dzuikevich") Then MsgBox "Нет доступа.", vbExclamation: Exit Sub
Unload Me
Run_PDF_AutoSign_FromExcel
End Sub


Private Sub CommandButton8_Click()
If Not Allowed("skrasilova;dzuikevich") Then MsgBox "Нет доступа.", vbExclamation: Exit Sub
Dim p As String
    p = "\\corp.vostok-electra.ru\Kgn\Отделы\Отдел взыскания по исполнительным документам\Зуйкевич Данил Иванович\Excel\ВЗИД\СОП\Порядок работы с макросом для автоподстановки подписи в уведомления и заявления.docx"   ' <— ваш путь

    If Len(Dir(p)) = 0 Then
        MsgBox "Файл не найден: " & p, vbExclamation
        Exit Sub
    End If

    ThisWorkbook.FollowHyperlink p   ' откроется в Word
End Sub

Private Sub CommandButton9_Click()
    If Not Allowed("skrasilova;dzuikevich") Then MsgBox "Нет доступа.", vbExclamation: Exit Sub
    Dim folderPath As String
    folderPath = "\\corp.vostok-electra.ru\Kgn\Отделы\Отдел взыскания по исполнительным документам\Зуйкевич Данил Иванович\Excel\ВЗИД\Подпись уведомлений и заявлений\" & Application.userName

    ' если папки нет — показать сообщение
    If Len(Dir(folderPath, vbDirectory)) = 0 Then
        MsgBox "Папка не найдена:" & vbCrLf & folderPath, vbExclamation
        Exit Sub
    End If

    ' открыть в Проводнике
    shell "explorer.exe """ & folderPath & """", vbNormalFocus
End Sub

Private Sub UserForm_Resize()
    Me.Image1.width = Me.width
    Me.Image1.Height = Me.Height
End Sub


Private Sub CommandButton1_Click()
    Unload Me
    Doc_Packets
End Sub

Private Sub CommandButton2_Click()
Unload Me
Make_Cover_Letters
End Sub

Private Sub CommandButton3_Click()
Unload Me
Make_Cover_Letters_ByRecipient
End Sub

Private Sub CommandButton4_Click()
Dim p As String
    p = "\\corp.vostok-electra.ru\Kgn\Отделы\Отдел взыскания по исполнительным документам\Зуйкевич Данил Иванович\Excel\ВЗИД\СОП\Порядок работы с макросом для формирования сопроводительных документов.docx"   ' <— ваш путь

    If Len(Dir(p)) = 0 Then
        MsgBox "Файл не найден: " & p, vbExclamation
        Exit Sub
    End If

    ThisWorkbook.FollowHyperlink p   ' откроется в Word
End Sub

Private Sub CommandButton5_Click()
Unload Me
KGN_Zakaznye_Create
End Sub

Private Sub CommandButton6_Click()
Unload Me
frmVZID_KGN_ManualProcReport.Show
End Sub

Private Sub Image1_Click()

End Sub

