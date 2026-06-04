Attribute VB_Name = "UpdateStatus"
Option Explicit

Public Function UpdateStatus_RibbonText() As String
    Dim statusValue As String
    Dim availableVersion As String
    Dim activeVersion As String
    Dim preparedVersion As String

    statusValue = MainConfig_ReadValue("lastUpdateStatus", "never")
    availableVersion = MainConfig_ReadValue("availableMainVersion", "")
    activeVersion = MainConfig_ReadValue("activeMainVersion", VZID_MAIN_VERSION)
    preparedVersion = MainConfig_ReadValue("preparedMainVersion", "")

    Select Case LCase$(statusValue)
        Case "update_available"
            UpdateStatus_RibbonText = "Обновления: доступна " & UpdateStatus_FormatVersion(availableVersion)
        Case "downloaded"
            If LenB(preparedVersion) > 0 Then
                UpdateStatus_RibbonText = "Обновления: скачана " & UpdateStatus_FormatVersion(preparedVersion)
            Else
                UpdateStatus_RibbonText = "Обновления: ждёт закрытия Excel"
            End If
        Case "scheduled"
            If LenB(preparedVersion) > 0 Then
                UpdateStatus_RibbonText = "Обновления: подготовлена " & UpdateStatus_FormatVersion(preparedVersion)
            Else
                UpdateStatus_RibbonText = "Обновления: ждёт закрытия Excel"
            End If
        Case "updater_upgrade_required"
            If LenB(availableVersion) > 0 Then
                UpdateStatus_RibbonText = "Обновления: нужен setup для " & UpdateStatus_FormatVersion(availableVersion)
            Else
                UpdateStatus_RibbonText = "Обновления: нужен новый setup.exe"
            End If
        Case "check_failed"
            UpdateStatus_RibbonText = "Обновления: проверить не удалось"
        Case "download_failed"
            UpdateStatus_RibbonText = "Обновления: скачивание не удалось"
        Case "activation_failed"
            UpdateStatus_RibbonText = "Обновления: применить не удалось"
        Case "up_to_date"
            UpdateStatus_RibbonText = "Обновления: " & UpdateStatus_FormatVersion(activeVersion)
        Case Else
            UpdateStatus_RibbonText = "Обновления: " & UpdateStatus_FormatVersion(activeVersion)
    End Select
End Function

Public Function UpdateStatus_CanDownloadUpdate() As Boolean
    UpdateStatus_CanDownloadUpdate = ( _
        (LCase$(MainConfig_ReadValue("lastUpdateStatus", "")) = "update_available" Or _
         LCase$(MainConfig_ReadValue("lastUpdateStatus", "")) = "updater_upgrade_required") And _
        LenB(MainConfig_ReadValue("availableMainVersion", "")) > 0)
End Function

Private Function UpdateStatus_FormatVersion(ByVal versionText As String) As String
    versionText = Trim$(versionText)
    If LenB(versionText) = 0 Then
        UpdateStatus_FormatVersion = "v?"
    ElseIf Left$(versionText, 1) = "v" Or Left$(versionText, 1) = "V" Then
        UpdateStatus_FormatVersion = versionText
    Else
        UpdateStatus_FormatVersion = "v" & versionText
    End If
End Function
