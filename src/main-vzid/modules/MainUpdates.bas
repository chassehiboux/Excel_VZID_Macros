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
    MainUpdates_LaunchPreparedUpdate targetPath, "manual-file", expectedSha256, True
    Exit Sub

failed:
    MainLogging_Write "MainUpdates_RunInstallFromFile failed: " & Err.Number & " - " & Err.Description
    MainConfig_WriteUpdateState "activation_failed", "Не удалось подготовить обновление из локального файла.", "", ""
    MsgBox "Не удалось подготовить обновление из файла." & vbCrLf & Err.Description, vbExclamation, "VZID"
End Sub

Private Sub MainUpdates_CheckManifest(ByVal interactive As Boolean)
    On Error GoTo failed

    Dim manifestUrl As String
    Dim manifestText As String
    Dim httpStatus As Long
    Dim activeVersion As String
    Dim availableVersion As String
    Dim downloadUrl As String
    Dim minimumUpdaterVersion As String
    Dim currentUpdaterVersion As String
    Dim statusMessage As String

    manifestUrl = MainConfig_ReadValue("manifestUrl", VZID_MANIFEST_URL)
    manifestText = MainUpdates_DownloadText(manifestUrl, httpStatus)

    If httpStatus <> 200 Or LenB(manifestText) = 0 Then
        statusMessage = "Не удалось проверить обновления по GitHub Releases."
        MainConfig_WriteUpdateState "check_failed", statusMessage, "", ""
        If interactive Then MsgBox statusMessage, vbExclamation, "VZID"
        Exit Sub
    End If

    activeVersion = MainConfig_ReadValue("activeMainVersion", VZID_MAIN_VERSION)
    currentUpdaterVersion = MainConfig_ReadValue("activeUpdaterVersion", VZID_UPDATER_VERSION)
    availableVersion = MainConfig_ReadJsonString(manifestText, "releaseVersion", "")
    downloadUrl = MainConfig_ReadJsonString(manifestText, "mainDownloadUrl", "")
    minimumUpdaterVersion = MainConfig_ReadJsonString(manifestText, "minUpdaterVersion", "")
    If LenB(minimumUpdaterVersion) = 0 Then
        minimumUpdaterVersion = MainConfig_ReadJsonString(manifestText, "minLoaderVersion", VZID_UPDATER_VERSION)
    End If

    If LenB(availableVersion) = 0 Then
        statusMessage = "Файл manifest.json не содержит releaseVersion."
        MainConfig_WriteUpdateState "check_failed", statusMessage, "", ""
        If interactive Then MsgBox statusMessage, vbExclamation, "VZID"
        Exit Sub
    End If

    If MainUpdates_CompareVersions(minimumUpdaterVersion, currentUpdaterVersion) > 0 Then
        statusMessage = "Для новой версии нужен новый setup.exe."
        MainConfig_WriteUpdateState "updater_upgrade_required", statusMessage, availableVersion, downloadUrl
        If interactive Then MsgBox statusMessage, vbExclamation, "VZID"
        Exit Sub
    End If

    If MainUpdates_CompareVersions(availableVersion, activeVersion) > 0 Then
        statusMessage = "Доступна новая версия " & availableVersion & "."
        MainConfig_WriteUpdateState "update_available", statusMessage, availableVersion, downloadUrl
        If interactive Then MsgBox statusMessage & vbCrLf & "Нажмите кнопку 'Обновить' на вкладке VZID.", vbInformation, "VZID"
    Else
        statusMessage = "Установлена актуальная версия " & activeVersion & "."
        MainConfig_WriteUpdateState "up_to_date", statusMessage, "", ""
        If interactive Then MsgBox statusMessage, vbInformation, "VZID"
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

    Dim downloadUrl As String
    Dim availableVersion As String
    Dim targetPath As String
    Dim statusCode As Long
    Dim expectedSha256 As String
    Dim actualSha256 As String
    Dim manifestUrl As String
    Dim manifestText As String
    Dim manifestStatus As Long

    availableVersion = MainConfig_ReadValue("availableMainVersion", "")
    downloadUrl = MainConfig_ReadValue("availableMainDownloadUrl", "")

    If LenB(availableVersion) = 0 Or LenB(downloadUrl) = 0 Then
        manifestUrl = MainConfig_ReadValue("manifestUrl", VZID_MANIFEST_URL)
        manifestText = MainUpdates_DownloadText(manifestUrl, manifestStatus)
        If manifestStatus = 200 And LenB(manifestText) > 0 Then
            availableVersion = MainConfig_ReadJsonString(manifestText, "releaseVersion", "")
            downloadUrl = MainConfig_ReadJsonString(manifestText, "mainDownloadUrl", "")
            expectedSha256 = LCase$(Trim$(MainConfig_ReadJsonString(manifestText, "mainSha256", "")))
        End If
    End If

    If LenB(expectedSha256) = 0 Then
        manifestUrl = MainConfig_ReadValue("manifestUrl", VZID_MANIFEST_URL)
        manifestText = MainUpdates_DownloadText(manifestUrl, manifestStatus)
        If manifestStatus <> 200 Or LenB(manifestText) = 0 Then
            MainConfig_WriteUpdateState "check_failed", "Не удалось повторно получить manifest.json перед скачиванием.", "", ""
            If interactive Then MsgBox "Не удалось повторно получить manifest.json перед скачиванием.", vbExclamation, "VZID"
            Exit Sub
        End If

        availableVersion = MainConfig_ReadJsonString(manifestText, "releaseVersion", availableVersion)
        downloadUrl = MainConfig_ReadJsonString(manifestText, "mainDownloadUrl", downloadUrl)
        expectedSha256 = LCase$(Trim$(MainConfig_ReadJsonString(manifestText, "mainSha256", "")))
    End If

    If LenB(availableVersion) = 0 Or LenB(downloadUrl) = 0 Then
        MainConfig_WriteUpdateState "check_failed", "В config нет данных о доступном обновлении.", "", ""
        If interactive Then MsgBox "Сначала выполните проверку обновлений.", vbExclamation, "VZID"
        Exit Sub
    End If

    targetPath = MainConfig_PreparedMainPathForVersion(availableVersion)
    MainConfig_EnsureBaseFolders

    If Not MainUpdates_DownloadBinary(downloadUrl, targetPath, statusCode) Then
        MainConfig_WriteUpdateState "download_failed", "Не удалось скачать файл обновления из GitHub Releases.", availableVersion, downloadUrl
        If interactive Then MsgBox "Не удалось скачать файл обновления из GitHub Releases.", vbExclamation, "VZID"
        Exit Sub
    End If

    If LenB(expectedSha256) = 0 Then
        MainConfig_WriteUpdateState "download_failed", "manifest.json не содержит mainSha256.", availableVersion, downloadUrl
        If interactive Then MsgBox "manifest.json не содержит mainSha256.", vbExclamation, "VZID"
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

    MainUpdates_LaunchPreparedUpdate targetPath, availableVersion, expectedSha256, interactive
    Exit Sub

failed:
    MainLogging_Write "MainUpdates_DownloadAvailableRelease failed: " & Err.Number & " - " & Err.Description
    MainConfig_WriteUpdateState "download_failed", "Ошибка при скачивании обновления.", "", ""
    If interactive Then
        MsgBox "Ошибка при скачивании обновления: " & Err.Description, vbExclamation, "VZID"
    End If
End Sub

Private Sub MainUpdates_LaunchPreparedUpdate(ByVal targetPath As String, ByVal releaseVersion As String, ByVal expectedSha256 As String, ByVal interactive As Boolean)
    On Error GoTo failed

    Dim restartNow As Boolean
    restartNow = False

    MainConfig_WriteValue "preparedMainVersion", releaseVersion
    MainConfig_WriteValue "preparedMainPath", targetPath
    MainConfig_WriteUpdateState "downloaded", "Обновление подготовлено. Оно установится после полного закрытия Excel.", releaseVersion, MainConfig_ReadValue("availableMainDownloadUrl", ""))

    If interactive Then
        restartNow = MainUpdates_AskRestartNow()
    End If

    If Not UpdaterBridge_StartPreparedUpdate(targetPath, releaseVersion, expectedSha256, restartNow) Then
        MainConfig_WriteUpdateState "updater_upgrade_required", "Не удалось запустить updater.exe. Запустите новый setup.exe.", "", ""
        If interactive Then
            MsgBox "Не удалось запустить updater.exe. Запустите новый setup.exe.", vbExclamation, "VZID"
        End If
        Exit Sub
    End If

    If restartNow Then
        Application.Quit
    ElseIf interactive Then
        MsgBox "Обновление подготовлено." & vbCrLf & "Оно установится после полного закрытия всех окон Excel.", vbInformation, "VZID"
    End If
    Exit Sub

failed:
    MainLogging_Write "MainUpdates_LaunchPreparedUpdate failed: " & Err.Number & " - " & Err.Description
    MainConfig_WriteUpdateState "activation_failed", "Не удалось подготовить обновление к установке.", "", ""
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
