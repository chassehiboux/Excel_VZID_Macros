VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVZID_EKB 
   Caption         =   "РИЦ"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5145
   OleObjectBlob   =   "frmVZID_EKB.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmVZID_EKB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ====== СЛУЖЕБНОЕ: обновление надписей версии/файла ======
Private Sub RefreshAboutVZID()
    On Error Resume Next
    lblVersion.Caption = "Версия (локальная): " & ReadLocalVersionVZID
End Sub

Private Sub CommandButton1_Click()
    Unload Me
    Doc_Packets
End Sub

Private Sub CommandButton10_Click()
Unload Me
frmInstallBrowser.Show
End Sub

Private Sub CommandButton4_Click()
Dim p As String
    p = "\\Ekb-vpfs01\екатеринбург\Отделы\Отдел взыскания задолженности по исполнительным документам\26. Макрос\ВЗИД\СОП\Порядок работы с макросом для формирования заказных реестров_Екат.docx"   ' <— ваш путь

    If Len(Dir(p)) = 0 Then
        MsgBox "Файл не найден: " & p, vbExclamation
        Exit Sub
    End If

    ThisWorkbook.FollowHyperlink p   ' откроется в Word
End Sub

Private Sub CommandButton5_Click()
Unload Me
RIC_Zakaznye_Create
End Sub

Private Sub CommandButton6_Click()
Unload Me
frmVZID_EKB_ManualProcReport.Show
End Sub

Private Sub UserForm_Initialize()
    RefreshAboutVZID
End Sub

Private Sub UserForm_Activate()
    RefreshAboutVZID
End Sub

