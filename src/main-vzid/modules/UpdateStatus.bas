Attribute VB_Name = "UpdateStatus"
Option Explicit

Public Function UpdateStatus_RibbonText() As String
    Dim statusValue As String

    statusValue = MainConfig_ReadValue("lastUpdateStatus", "never")

    Select Case LCase$(statusValue)
        Case "update_available"
            UpdateStatus_RibbonText = "Обновления: доступна " & MainConfig_ReadValue("availableMainVersion", "")
        Case "downloaded"
            UpdateStatus_RibbonText = "Обновления: скачано, перезапустите Excel"
        Case "check_failed"
            UpdateStatus_RibbonText = "Обновления: проверить не удалось"
        Case "up_to_date"
            UpdateStatus_RibbonText = "Обновления: версия актуальна"
        Case Else
            UpdateStatus_RibbonText = "Обновления: еще не проверялись"
    End Select
End Function
