Attribute VB_Name = "MainUpdates"
Option Explicit

Public Sub MainUpdates_CheckStartup()
    MainUpdates_CheckManifest False
End Sub

Public Sub MainUpdates_RunManualCheck()
    MainUpdates_CheckManifest True
End Sub

Public Sub MainUpdates_RunDownloadAvailableRelease()
    MainUpdates_DownloadAvailableRelease True
End Sub

Public Sub MainUpdates_RunInstallFromFile(ByVal sourceFile As String)
    On Error GoTo failed

    Dim targetPath As String
    Dim expectedSha256 As String

    MainConfig_EnsureBaseFolders
    targetPath = MainConfig_PreparedMainPathForVersion("manual-file")

    If Not MainConfig_CopyFile(sourceFile, targetPath) Then
        MsgBox "Не удалось подготовить файл обновления.", vbExclamation, "VZID"
        Exit Sub
    End If

    expectedSha256 = MainHash_FileSha256Hex(targetPath)
    MainUpdates_LaunchPreparedUpdate targetPath, "manual-file", expectedSha256, True, False
    Exit Sub

failed:
    MainLogging_Write "MainUpdates_RunInstallFromFile failed: " & Err.Number & " - " & Err.Description
    MainConfig_WriteUpdateState "activation_failed", "Не удалось подготовить обновление из локального файла.", "", ""
    MsgBox "Не удалось подготовить обновление из файла." & vbCrLf & Err.Description, vbExclamation, "VZID"
End Sub

Private Sub MainUpdates_CheckManifest(ByVal interactive As Boolean)
    On Error GoTo failed

    Dim manifestText As String
    Dim httpStatus As Long
    Dim activeVersion As String
    Dim availableVersion As String
    Dim currentUpdaterVersion As String
    Dim minimumUpdaterVersion As String
    Dim statusMessage As String
    Dim downloadUrl As String

    manifestText = MainUpdates_DownloadCurrentManifest(httpStatus)
    If httpStatus <> 200 Or LenB(manifestText) = 0 Then
        statusMessage = "Не удалось проверить обновления по GitHub Releases."
        MainConfig_WriteUpdateState "check_failed", statusMessage, "", ""
        If interactive Then MsgBox statusMessage, vbExclamation, "VZID"
        Exit Sub
    End If

    activeVersion = MainConfig_ReadValue("activeMainVersion", VZID_MAIN_VERSION)
    currentUpdaterVersion = MainConfig_ReadValue("activeUpdaterVersion", VZID_UPDATER_VERSION)
    availableVersion = MainConfig_ReadJsonString(manifestText, "releaseVersion", "")

    If LenB(availableVersion) = 0 Then
        statusMessage = "Файл manifest.json не содержит releaseVersion."
        MainConfig_WriteUpdateState "check_failed", statusMessage, "", ""
        If interactive Then MsgBox statusMessage, vbExclamation, "VZID"
        Exit Sub
    End If

    If MainUpdates_CompareVersions(availableVersion, activeVersion) <= 0 Then
        statusMessage = "Установлена версия " & MainUpdates_FormatVersion(Trim$(activeVersion)) & "."
        MainConfig_WriteUpdateState "up_to_date", statusMessage, "", ""
        If interactive Then MsgBox statusMessage, vbInformation, "VZID"
        Exit Sub
    End If

    minimumUpdaterVersion = MainUpdates_ReadMinUpdaterVersion(manifestText)
    If MainUpdates_CompareVersions(minimumUpdaterVersion, currentUpdaterVersion) > 0 Then
        downloadUrl = MainConfig_ReadJsonString(manifestText, "setupDownloadUrl", "")
        statusMessage = "Для версии " & MainUpdates_FormatVersion(availableVersion) & " нужен новый setup.exe."
        MainConfig_WriteUpdateState "updater_upgrade_required", statusMessage, availableVersion, downloadUrl
        If interactive Then
            MsgBox statusMessage & vbCrLf & "Нажмите кнопку 'Обновить' на вкладке VZID.", vbInformation, "VZID"
        End If
        Exit Sub
    End If

    downloadUrl = MainConfig_ReadJsonString(manifestText, "mainDownloadUrl", "")
    If LenB(downloadUrl) = 0 Then
        statusMessage = "Файл manifest.json не содержит mainDownloadUrl."
        MainConfig_WriteUpdateState "check_failed", statusMessage, "", ""
        If interactive Then MsgBox statusMessage, vbExclamation, "VZID"
        Exit Sub
    End If

    statusMessage = "Доступна новая версия " & MainUpdates_FormatVersion(availableVersion) & "."
    MainConfig_WriteUpdateState "update_available", statusMessage, availableVersion, downloadUrl
    If interactive Then
        MsgBox statusMessage & vbCrLf & "Нажмите кнопку 'Обновить' на вкладке VZID.", vbInformation, "VZID"
    End If
    Exit Sub

failed:
    MainLogging_Write "MainUpdates_CheckManifest failed: " & Err.Number & " - " & Err.Description
    MainConfig_WriteUpdateState "check_failed", "Проверка обновлений завершилась ошибкой.", "", ""
    If interactive Then
        MsgBox "Проверка обновлений завершилась ошибкой: " & Err.Description, vbExclamation, "VZID"
    End If
End Sub

Private Sub MainUpdates_DownloadAvailableRelease(ByVal interactive As Boolean)
    On Error GoTo failed

    Dim manifestText As String
    Dim manifestStatus As Long
    Dim activeVersion As String
    Dim availableVersion As String
    Dim currentUpdaterVersion As String
    Dim minimumUpdaterVersion As String

    manifestText = MainUpdates_DownloadCurrentManifest(manifestStatus)
    If manifestStatus <> 200 Or LenB(manifestText) = 0 Then
        MainConfig_WriteUpdateState "check_failed", "Не удалось повторно получить manifest.json перед скачиванием.", "", ""
        If interactive Then MsgBox "Не удалось повторно получить manifest.json перед скачиванием.", vbExclamation, "VZID"
        Exit Sub
    End If

    activeVersion = MainConfig_ReadValue("activeMainVersion", VZID_MAIN_VERSION)
    currentUpdaterVersion = MainConfig_ReadValue("activeUpdaterVersion", VZID_UPDATER_VERSION)
    availableVersion = MainConfig_ReadJsonString(manifestText, "releaseVersion", "")

    If LenB(availableVersion) = 0 Then
        MainConfig_WriteUpdateState "check_failed", "manifest.json не содержит releaseVersion.", "", ""
        If interactive Then MsgBox "manifest.json не содержит releaseVersion.", vbExclamation, "VZID"
        Exit Sub
    End If

    If MainUpdates_CompareVersions(availableVersion, activeVersion) <= 0 Then
        MainConfig_WriteUpdateState "up_to_date", "Установлена версия " & MainUpdates_FormatVersion(activeVersion) & ".", "", ""
        If interactive Then MsgBox "У вас уже установлена версия " & MainUpdates_FormatVersion(activeVersion) & ".", vbInformation, "VZID"
        Exit Sub
    End If

    minimumUpdaterVersion = MainUpdates_ReadMinUpdaterVersion(manifestText)
    If MainUpdates_CompareVersions(minimumUpdaterVersion, currentUpdaterVersion) > 0 Then
        MainUpdates_DownloadSetupRelease manifestText, availableVersion, interactive
    Else
        MainUpdates_DownloadMainRelease manifestText, availableVersion, interactive
    End If
    Exit Sub

failed:
    MainLogging_Write "MainUpdates_DownloadAvailableRelease failed: " & Err.Number & " - " & Err.Description
    MainConfig_WriteUpdateState "download_failed", "Ошибка при скачивании обновления.", "", ""
    If interactive Then
        MsgBox "Ошибка при скачивании обновления: " & Err.Description, vbExclamation, "VZID"
    End If
End Sub

Private Sub MainUpdates_DownloadMainRelease(ByVal manifestText As String, ByVal availableVersion As String, ByVal interactive As Boolean)
    On Error GoTo failed

    Dim downloadUrl As String
    Dim expectedSha256 As String
    Dim actualSha256 As String
    Dim targetPath As String
    Dim statusCode As Long

    downloadUrl = MainConfig_ReadJsonString(manifestText, "mainDownloadUrl", "")
    expectedSha256 = LCase$(Trim$(MainConfig_ReadJsonString(manifestText, "mainSha256", "")))

    If LenB(downloadUrl) = 0 Then
        MainConfig_WriteUpdateState "download_failed", "manifest.json не содержит mainDownloadUrl.", availableVersion, ""
        If interactive Then MsgBox "manifest.json не содержит mainDownloadUrl.", vbExclamation, "VZID"
        Exit Sub
    End If

    If LenB(expectedSha256) = 0 Then
        MainConfig_WriteUpdateState "download_failed", "manifest.json не содержит mainSha256.", availableVersion, downloadUrl
        If interactive Then MsgBox "manifest.json не содержит mainSha256.", vbExclamation, "VZID"
        Exit Sub
    End If

    MainConfig_EnsureBaseFolders
    targetPath = MainConfig_PreparedMainPathForVersion(availableVersion)

    If Not MainUpdates_DownloadBinary(downloadUrl, targetPath, statusCode) Then
        MainConfig_WriteUpdateState "download_failed", "Не удалось скачать файл обновления из GitHub Releases.", availableVersion, downloadUrl
        If interactive Then MsgBox "Не удалось скачать файл обновления из GitHub Releases.", vbExclamation, "VZID"
        Exit Sub
    End If

    actualSha256 = MainHash_FileSha256Hex(targetPath)
    If LenB(actualSha256) = 0 Or StrComp(actualSha256, expectedSha256, vbTextCompare) <> 0 Then
        On Error Resume Next
        Kill targetPath
        On Error GoTo failed

        MainConfig_WriteUpdateState "download_failed", "Контрольная сумма обновления не совпала. Файл отклонен.", availableVersion, downloadUrl
        If interactive Then MsgBox "Контрольная сумма обновления не совпала. Файл отклонен.", vbExclamation, "VZID"
        Exit Sub
    End If

    MainConfig_WriteUpdateState "update_available", "Доступна новая версия " & MainUpdates_FormatVersion(availableVersion) & ".", availableVersion, downloadUrl
    MainUpdates_LaunchPreparedUpdate targetPath, availableVersion, expectedSha256, interactive, False
    Exit Sub

failed:
    MainLogging_Write "MainUpdates_DownloadMainRelease failed: " & Err.Number & " - " & Err.Description
    MainConfig_WriteUpdateState "download_failed", "Ошибка при скачивании файла MainVZID.xlam.", availableVersion, ""
    If interactive Then
        MsgBox "Ошибка при скачивании файла MainVZID.xlam: " & Err.Description, vbExclamation, "VZID"
    End If
End Sub

Private Sub MainUpdates_DownloadSetupRelease(ByVal manifestText As String, ByVal availableVersion As String, ByVal interactive As Boolean)
    On Error GoTo failed

    Dim downloadUrl As String
    Dim expectedSha256 As String
    Dim actualSha256 As String
    Dim targetPath As String
    Dim statusCode As Long

    downloadUrl = MainConfig_ReadJsonString(manifestText, "setupDownloadUrl", "")
    expectedSha256 = LCase$(Trim$(MainConfig_ReadJsonString(manifestText, "setupSha256", "")))

    If LenB(downloadUrl) = 0 Then
        MainConfig_WriteUpdateState "download_failed", "manifest.json не содержит setupDownloadUrl.", availableVersion, ""
        If interactive Then MsgBox "manifest.json не содержит setupDownloadUrl.", vbExclamation, "VZID"
        Exit Sub
    End If

    If LenB(expectedSha256) = 0 Then
        MainConfig_WriteUpdateState "download_failed", "manifest.json не содержит setupSha256.", availableVersion, downloadUrl
        If interactive Then MsgBox "manifest.json не содержит setupSha256.", vbExclamation, "VZID"
        Exit Sub
    End If

    MainConfig_EnsureBaseFolders
    targetPath = MainConfig_PreparedSetupPathForVersion(availableVersion)

    If Not MainUpdates_DownloadBinary(downloadUrl, targetPath, statusCode) Then
        MainConfig_WriteUpdateState "download_failed", "Не удалось скачать новый setup.exe из GitHub Releases.", availableVersion, downloadUrl
        If interactive Then MsgBox "Не удалось скачать новый setup.exe из GitHub Releases.", vbExclamation, "VZID"
        Exit Sub
    End If

    actualSha256 = MainHash_FileSha256Hex(targetPath)
    If LenB(actualSha256) = 0 Or StrComp(actualSha256, expectedSha256, vbTextCompare) <> 0 Then
        On Error Resume Next
        Kill targetPath
        On Error GoTo failed

        MainConfig_WriteUpdateState "download_failed", "Контрольная сумма setup.exe не совпала. Файл отклонен.", availableVersion, downloadUrl
        If interactive Then MsgBox "Контрольная сумма setup.exe не совпала. Файл отклонен.", vbExclamation, "VZID"
        Exit Sub
    End If

    MainConfig_WriteUpdateState "updater_upgrade_required", "Для версии " & MainUpdates_FormatVersion(availableVersion) & " нужен новый setup.exe.", availableVersion, downloadUrl
    MainUpdates_LaunchPreparedUpdate targetPath, availableVersion, expectedSha256, interactive, True
    Exit Sub

failed:
    MainLogging_Write "MainUpdates_DownloadSetupRelease failed: " & Err.Number & " - " & Err.Description
    MainConfig_WriteUpdateState "download_failed", "Ошибка при скачивании файла setup.exe.", availableVersion, ""
    If interactive Then
        MsgBox "Ошибка при скачивании файла setup.exe: " & Err.Description, vbExclamation, "VZID"
    End If
End Sub

Private Sub MainUpdates_LaunchPreparedUpdate(ByVal targetPath As String, ByVal releaseVersion As String, ByVal expectedSha256 As String, ByVal interactive As Boolean, ByVal installerMode As Boolean)
    On Error GoTo failed

    Dim restartNow As Boolean
    Dim statusMessage As String
    Dim successMessage As String

    restartNow = False
    MainConfig_WriteValue "preparedMainVersion", releaseVersion
    MainConfig_WriteValue "preparedMainPath", targetPath

    If installerMode Then
        statusMessage = "Новый setup.exe подготовлен. Он будет запущен после полного закрытия Excel."
    Else
        statusMessage = "Обновление подготовлено. Оно установится после полного закрытия Excel."
    End If

    MainConfig_WriteUpdateState "downloaded", statusMessage, releaseVersion, MainConfig_ReadValue("availableMainDownloadUrl", ""))

    If interactive Then
        restartNow = MainUpdates_AskRestartNow()
    End If

    If Not UpdaterBridge_StartPreparedUpdate(targetPath, releaseVersion, expectedSha256, restartNow, installerMode) Then
        MainConfig_WriteUpdateState "activation_failed", "Не удалось запустить updater.exe.", releaseVersion, ""
        If interactive Then
            MsgBox "Не удалось запустить updater.exe.", vbExclamation, "VZID"
        End If
        Exit Sub
    End If

    If restartNow Then
        Application.Quit
    ElseIf interactive Then
        successMessage = "Обновление подготовлено." & vbCrLf & "Оно установится после полного закрытия всех окон Excel."
        If installerMode Then
            successMessage = "Новый setup.exe подготовлен." & vbCrLf & "Он будет запущен после полного закрытия всех окон Excel."
        End If
        MsgBox successMessage, vbInformation, "VZID"
    End If
    Exit Sub

failed:
    MainLogging_Write "MainUpdates_LaunchPreparedUpdate failed: " & Err.Number & " - " & Err.Description
    MainConfig_WriteUpdateState "activation_failed", "Не удалось подготовить обновление к установке.", releaseVersion, ""
    If interactive Then
        MsgBox "Не удалось подготовить обновление к установке: " & Err.Description, vbExclamation, "VZID"
    End If
End Sub

Private Function MainUpdates_AskRestartNow() As Boolean
    Dim reply As VbMsgBoxResult

    reply = MsgBox( _
        "Обновление подготовлено." & vbCrLf & _
        "Оно установится только после полного закрытия всех окон Excel." & vbCrLf & vbCrLf & _
        "Закрыть Excel сейчас?", _
        vbYesNo + vbQuestion, _
        "VZID")

    MainUpdates_AskRestartNow = (reply = vbYes)
End Function

Private Function MainUpdates_DownloadCurrentManifest(ByRef statusCode As Long) As String
    MainUpdates_DownloadCurrentManifest = MainUpdates_DownloadText( _
        MainConfig_ReadValue("manifestUrl", VZID_MANIFEST_URL), _
        statusCode)
End Function

Private Function MainUpdates_ReadMinUpdaterVersion(ByVal manifestText As String) As String
    MainUpdates_ReadMinUpdaterVersion = MainConfig_ReadJsonString(manifestText, "minUpdaterVersion", "")
    If LenB(MainUpdates_ReadMinUpdaterVersion) = 0 Then
        MainUpdates_ReadMinUpdaterVersion = MainConfig_ReadJsonString(manifestText, "minLoaderVersion", VZID_UPDATER_VERSION)
    End If
End Function

Private Function MainUpdates_FormatVersion(ByVal versionText As String) As String
    versionText = Trim$(versionText)
    If LenB(versionText) = 0 Then
        MainUpdates_FormatVersion = "v?"
    ElseIf Left$(versionText, 1) = "v" Or Left$(versionText, 1) = "V" Then
        MainUpdates_FormatVersion = versionText
    Else
        MainUpdates_FormatVersion = "v" & versionText
    End If
End Function

Private Function MainUpdates_DownloadText(ByVal sourceUrl As String, ByRef statusCode As Long) As String
    On Error GoTo failed

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    http.setTimeouts 3000, 3000, 5000, 5000
    http.Open "GET", sourceUrl, False
    http.setRequestHeader "User-Agent", "VZID-Main/" & MainConfig_ReadValue("activeMainVersion", VZID_MAIN_VERSION)
    http.send

    statusCode = CLng(http.Status)
    If statusCode = 200 Then
        MainUpdates_DownloadText = CStr(http.responseText)
    End If
    Exit Function

failed:
    statusCode = 0
    MainLogging_Write "MainUpdates_DownloadText failed: " & Err.Number & " - " & Err.Description
End Function

Private Function MainUpdates_DownloadBinary(ByVal sourceUrl As String, ByVal targetPath As String, ByRef statusCode As Long) As Boolean
    On Error GoTo failed

    Dim http As Object
    Dim stream As Object

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 5000, 5000, 30000, 30000
    http.Open "GET", sourceUrl, False
    http.setRequestHeader "User-Agent", "VZID-Main/" & MainConfig_ReadValue("activeMainVersion", VZID_MAIN_VERSION)
    http.send

    statusCode = CLng(http.Status)
    If statusCode <> 200 Then Exit Function

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile targetPath, 2
    stream.Close

    MainUpdates_DownloadBinary = (LenB(Dir$(targetPath)) > 0)
    Exit Function

failed:
    statusCode = 0
    MainLogging_Write "MainUpdates_DownloadBinary failed: " & Err.Number & " - " & Err.Description
End Function

Private Function MainUpdates_CompareVersions(ByVal leftVersion As String, ByVal rightVersion As String) As Long
    Dim leftParts() As String
    Dim rightParts() As String
    Dim maxIndex As Long
    Dim index As Long
    Dim leftPart As Long
    Dim rightPart As Long

    leftParts = Split(leftVersion, ".")
    rightParts = Split(rightVersion, ".")

    maxIndex = UBound(leftParts)
    If UBound(rightParts) > maxIndex Then maxIndex = UBound(rightParts)

    For index = 0 To maxIndex
        leftPart = 0
        rightPart = 0

        If index <= UBound(leftParts) Then leftPart = CLng(Val(leftParts(index)))
        If index <= UBound(rightParts) Then rightPart = CLng(Val(rightParts(index)))

        If leftPart > rightPart Then
            MainUpdates_CompareVersions = 1
            Exit Function
        End If

        If leftPart < rightPart Then
            MainUpdates_CompareVersions = -1
            Exit Function
        End If
    Next index
End Function
