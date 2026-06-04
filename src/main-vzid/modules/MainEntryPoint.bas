Attribute VB_Name = "MainEntryPoint"
Option Explicit

Private mInitialized As Boolean

Public Sub MainEntryPoint_Initialize()
    If mInitialized Then Exit Sub

    MainConfig_EnsureConfig
    RegionState_Initialize
    MainUpdates_CheckStartup

    mInitialized = True
End Sub

Public Sub MainEntryPoint_RunCheckUpdates()
    MainUpdates_RunManualCheck
    RibbonVZID_Invalidate
End Sub

Public Sub MainEntryPoint_RunDownloadUpdate()
    MainUpdates_RunDownloadAvailableRelease
    RibbonVZID_Invalidate
End Sub

Public Sub MainEntryPoint_InstallUpdateFromFile()
    On Error GoTo failed

    Dim selectedFile As Variant
    selectedFile = Application.GetOpenFilename("Надстройки Excel (*.xlam), *.xlam", , "Выберите файл обновления MainVZID")
    If VarType(selectedFile) = vbBoolean Then Exit Sub

    MainUpdates_RunInstallFromFile CStr(selectedFile)
    RibbonVZID_Invalidate
    Exit Sub

failed:
    MainLogging_Write "MainEntryPoint_InstallUpdateFromFile failed: " & Err.Number & " - " & Err.Description
    MsgBox "Не удалось подготовить обновление из файла." & vbCrLf & Err.Description, vbExclamation, "VZID"
End Sub
