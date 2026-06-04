Attribute VB_Name = "UpdaterBridge"
Option Explicit

Public Function UpdaterBridge_StartPreparedUpdate(ByVal sourcePath As String, ByVal releaseVersion As String, ByVal expectedSha256 As String, ByVal restartExcel As Boolean, Optional ByVal installerMode As Boolean = False) As Boolean
    On Error GoTo failed

    If LenB(Dir$(MainConfig_UpdaterExePath())) = 0 Then Exit Function
    If LenB(Dir$(sourcePath)) = 0 Then Exit Function

    Dim commandText As String
    commandText = UpdaterBridge_Quote(MainConfig_UpdaterExePath()) & _
        " --source " & UpdaterBridge_Quote(sourcePath) & _
        " --target " & UpdaterBridge_Quote(ThisWorkbook.FullName) & _
        " --config " & UpdaterBridge_Quote(MainConfig_ConfigPath()) & _
        " --backup-dir " & UpdaterBridge_Quote(MainConfig_BackupDir()) & _
        " --log " & UpdaterBridge_Quote(MainConfig_LogsDir() & "updater.log") & _
        " --version " & UpdaterBridge_Quote(releaseVersion) & _
        " --mode " & UpdaterBridge_Quote(IIf(installerMode, "setup", "main")) & _
        " --restart-excel " & CStr(Abs(restartExcel))

    If LenB(Trim$(expectedSha256)) > 0 Then
        commandText = commandText & " --expected-sha256 " & UpdaterBridge_Quote(expectedSha256)
    End If

    Dim shellObject As Object
    Set shellObject = CreateObject("WScript.Shell")
    shellObject.Run commandText, 0, False

    UpdaterBridge_StartPreparedUpdate = True
    Exit Function

failed:
    MainLogging_Write "UpdaterBridge_StartPreparedUpdate failed: " & Err.Number & " - " & Err.Description
End Function

Private Function UpdaterBridge_Quote(ByVal value As String) As String
    UpdaterBridge_Quote = """" & Replace(value, """", """""") & """"
End Function
