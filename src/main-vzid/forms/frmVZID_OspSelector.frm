VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVZID_OspSelector 
   Caption         =   "Выбор ОСП"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8565.001
   OleObjectBlob   =   "frmVZID_OspSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmVZID_OspSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim i As Long

    With Me.lstOsp
        .Clear
        .ColumnCount = 3
        ' Видно только название ОСП, адрес и регион спрятаны (можно ширину увеличить)
        .ColumnWidths = "220;0;0"
    End With

    For i = 1 To OspSelector.OspCount
        With Me.lstOsp
            .AddItem OspSelector.OspList(i).OspName                    ' колонка 0
            .List(.ListCount - 1, 1) = OspSelector.OspList(i).Address  ' колонка 1
            .List(.ListCount - 1, 2) = OspSelector.OspList(i).Region   ' колонка 2
        End With
    Next i
End Sub

Private Sub cmdOK_Click()
    Dim idx As Long

    idx = Me.lstOsp.ListIndex
    If idx < 0 Then
        MsgBox "Выбери ОСП в списке.", vbExclamation
        Exit Sub
    End If

    If Not OspSelector.TargetCell Is Nothing Then
        ' В выбранную ячейку — ОСП
        OspSelector.TargetCell.value = Me.lstOsp.List(idx, 0)

        ' В ячейку слева — Адрес (если столбец есть)
        If OspSelector.TargetCell.Column > 1 Then
            OspSelector.TargetCell.Offset(0, -1).value = Me.lstOsp.List(idx, 1)
        End If
    End If

    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub lstOsp_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' Двойной щелчок по списку = как ОК
    cmdOK_Click
End Sub

