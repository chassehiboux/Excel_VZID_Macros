Attribute VB_Name = "MainEntryPoint"
Option Explicit

Private Const MAIN_LOADER_WORKBOOK As String = "LoaderVZID.xlam"
Private Const MAIN_LOADER_CHECK_MACRO As String = "LoaderUpdates_RunManualCheck"

Public Sub MainEntryPoint_Initialize()
    MainConfig_EnsureConfig
    RegionState_Initialize
End Sub

Public Sub MainEntryPoint_RunCheckUpdates()
    On Error GoTo failed

    Application.Run "'" & MAIN_LOADER_WORKBOOK & "'!" & MAIN_LOADER_CHECK_MACRO
    RibbonVZID_Invalidate
    Exit Sub

failed:
    MainConfig_WriteValue "lastUpdateStatus", "check_failed"
    MainConfig_WriteValue "lastUpdateMessage", "Не удалось вызвать загрузчик для проверки обновлений."
    RibbonVZID_Invalidate

    MsgBox "Не удалось запустить проверку обновлений. Убедитесь, что LoaderVZID подключен.", vbExclamation, "VZID"
End Sub

Public Sub MainEntryPoint_InstallUpdateFromFile()
    On Error GoTo failed

    Dim selectedFile As Variant
    selectedFile = Application.GetOpenFilename("Надстройки Excel (*.xlam), *.xlam", , "Выберите файл обновления MainVZID")
    If VarType(selectedFile) = vbBoolean Then Exit Sub

    MainConfig_EnsureBaseFolders

    If Not MainConfig_CopyFile(CStr(selectedFile), MainConfig_PendingMainPath()) Then
        MsgBox "Не удалось подготовить файл обновления.", vbExclamation, "VZID"
        Exit Sub
    End If

    MainConfig_WriteValue "pendingMainPath", MainConfig_PendingMainPath()
    MainConfig_WriteValue "pendingMainVersion", "manual-file"
    MainConfig_WriteValue "lastUpdateStatus", "downloaded"
    MainConfig_WriteValue "lastUpdateMessage", "Файл обновления подготовлен. Полностью закройте все окна Excel и откройте Excel заново."

    RibbonVZID_Invalidate
    MainEntryPoint_ShowRestartPrompt
    Exit Sub

failed:
    MsgBox "Не удалось подготовить обновление из файла." & vbCrLf & Err.Description, vbExclamation, "VZID"
End Sub

Private Sub MainEntryPoint_ShowRestartPrompt()
    Dim reply As VbMsgBoxResult

    reply = MsgBox( _
        "Обновление подготовлено." & vbCrLf & _
        "Оно установится только после полного закрытия всех окон Excel и нового запуска." & vbCrLf & vbCrLf & _
        "Закрыть Excel сейчас?", _
        vbYesNo + vbQuestion, _
        "VZID")

    If reply = vbYes Then
        Application.Quit
    End If
End Sub
