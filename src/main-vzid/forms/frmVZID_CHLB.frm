VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVZID_CHLB 
   Caption         =   "ЧЭС"
   ClientHeight    =   2970
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5145
   OleObjectBlob   =   "frmVZID_CHLB.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmVZID_CHLB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Unload Me
    Doc_Packets
End Sub

Private Sub CommandButton10_Click()
Unload Me
frmInstallBrowser.Show
End Sub

Private Sub CommandButton2_Click()
Unload Me
CHLB_Make_Cover_Letters
End Sub

Private Sub CommandButton3_Click()
Unload Me
CHLB_Make_Cover_Letters_ByRecipient
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

