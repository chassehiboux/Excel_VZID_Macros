Attribute VB_Name = "CommandDispatcher"
Option Explicit

Public Sub CommandDispatcher_Run(ByVal commandId As String)
    On Error GoTo failed

    Dim regionId As String
    Dim macroName As String

    regionId = RegionState_SelectedId()
    If Not AccessPolicy_IsCommandEnabled(commandId) Then
        MsgBox "У вас нет доступа к этой команде.", vbExclamation, "VZID"
        Exit Sub
    End If

    macroName = CommandRegistry_GetMacroName(commandId, regionId)
    If LenB(macroName) = 0 Then
        MsgBox "Для выбранного региона эта команда пока не настроена.", vbInformation, "VZID"
        Exit Sub
    End If

    Application.Run "'" & ThisWorkbook.Name & "'!" & macroName
    Exit Sub

failed:
    MsgBox "Не удалось запустить макрос " & macroName & "." & vbCrLf & Err.Description, vbExclamation, "VZID"
End Sub
