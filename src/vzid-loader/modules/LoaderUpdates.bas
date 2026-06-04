Attribute VB_Name = "LoaderUpdates"
Option Explicit

Public Sub LoaderUpdates_CheckStartup()
    LoaderUpdates_CheckManifest False
End Sub

Public Sub LoaderUpdates_RunManualCheck()
    LoaderUpdates_CheckManifest True
End Sub

Public Sub LoaderUpdates_ActivatePendingIfPresent()
    On Error GoTo failed

    Dim pendingPath As String
    Dim pendingVersion As String

    pendingPath = LoaderConfig_ReadValue("pendingMainPath", "")
    pendingVersion = LoaderConfig_ReadValue("pendingMainVersion", "")

    If LenB(pendingPath) = 0 Then Exit Sub
    If LenB(Dir$(pendingPath)) = 0 Then Exit Sub

    LoaderAddinHost_UnloadByPath LoaderPaths_MainAddinPath()

    If LoaderUpdates_CopyFile(pendingPath, LoaderPaths_MainAddinPath()) Then
        LoaderConfig_WriteValue "activeMainVersion", pendingVersion
        LoaderConfig_WriteValue "pendingMainVersion", ""
        LoaderConfig_WriteValue "pendingMainPath", ""
        LoaderConfig_WriteUpdateState "up_to_date", "Обновление применено после перезапуска Excel.", "", ""

        On Error Resume Next
        Kill pendingPath

        LoaderLogging_Write "Pending update activated: " & pendingVersion
    Else
        LoaderConfig_WriteUpdateState "activation_failed", "Не удалось применить подготовленное обновление.", "", ""
    End If
    Exit Sub

failed:
    LoaderLogging_Write "LoaderUpdates_ActivatePendingIfPresent failed: " & Err.Number & " - " & Err.Description
End Sub

Private Sub LoaderUpdates_CheckManifest(ByVal interactive As Boolean)
    On Error GoTo failed

    Dim manifestUrl As String
    Dim manifestText As String
    Dim httpStatus As Long
    Dim activeVersion As String
    Dim availableVersion As String
    Dim downloadUrl As String
    Dim statusMessage As String

    manifestUrl = LoaderConfig_ReadValue("manifestUrl", VZID_MANIFEST_URL)
    manifestText = LoaderUpdates_DownloadText(manifestUrl, httpStatus)

    If httpStatus <> 200 Or LenB(manifestText) = 0 Then
        statusMessage = "Не удалось проверить обновления по GitHub Releases."
        LoaderConfig_WriteUpdateState "check_failed", statusMessage, "", ""
        If interactive Then MsgBox statusMessage, vbExclamation, "VZID"
        Exit Sub
    End If

    activeVersion = LoaderConfig_ReadValue("activeMainVersion", "0.0.0")
    availableVersion = LoaderConfig_ReadJsonString(manifestText, "releaseVersion", "")
    downloadUrl = LoaderConfig_ReadJsonString(manifestText, "mainDownloadUrl", "")

    If LenB(availableVersion) = 0 Then
        statusMessage = "Файл manifest.json не содержит releaseVersion."
        LoaderConfig_WriteUpdateState "check_failed", statusMessage, "", ""
        If interactive Then MsgBox statusMessage, vbExclamation, "VZID"
        Exit Sub
    End If

    If LoaderUpdates_CompareVersions(availableVersion, activeVersion) > 0 Then
        statusMessage = "Доступна новая версия " & availableVersion & "."
        LoaderConfig_WriteUpdateState "update_available", statusMessage, availableVersion, downloadUrl
        If interactive Then MsgBox statusMessage & vbCrLf & "Скачивание релиза подключим следующим шагом.", vbInformation, "VZID"
    Else
        statusMessage = "Установлена актуальная версия " & activeVersion & "."
        LoaderConfig_WriteUpdateState "up_to_date", statusMessage, availableVersion, downloadUrl
        If interactive Then MsgBox statusMessage, vbInformation, "VZID"
    End If
    Exit Sub

failed:
    LoaderLogging_Write "LoaderUpdates_CheckManifest failed: " & Err.Number & " - " & Err.Description
    LoaderConfig_WriteUpdateState "check_failed", "Проверка обновлений завершилась ошибкой.", "", ""
    If interactive Then
        MsgBox "Проверка обновлений завершилась ошибкой: " & Err.Description, vbExclamation, "VZID"
    End If
End Sub

Private Function LoaderUpdates_DownloadText(ByVal sourceUrl As String, ByRef statusCode As Long) As String
    On Error GoTo failed

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    http.setTimeouts 3000, 3000, 5000, 5000
    http.Open "GET", sourceUrl, False
    http.setRequestHeader "User-Agent", "VZID-Loader/" & VZID_LOADER_VERSION
    http.send

    statusCode = CLng(http.Status)
    If statusCode = 200 Then
        LoaderUpdates_DownloadText = CStr(http.responseText)
    End If
    Exit Function

failed:
    statusCode = 0
    LoaderLogging_Write "LoaderUpdates_DownloadText failed: " & Err.Number & " - " & Err.Description
End Function

Private Function LoaderUpdates_CopyFile(ByVal sourcePath As String, ByVal targetPath As String) As Boolean
    On Error GoTo failed

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(targetPath) Then
        fso.DeleteFile targetPath, True
    End If

    fso.CopyFile sourcePath, targetPath, True
    LoaderUpdates_CopyFile = fso.FileExists(targetPath)
    Exit Function

failed:
    LoaderLogging_Write "LoaderUpdates_CopyFile failed: " & Err.Number & " - " & Err.Description
End Function

Private Function LoaderUpdates_CompareVersions(ByVal leftVersion As String, ByVal rightVersion As String) As Long
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
            LoaderUpdates_CompareVersions = 1
            Exit Function
        End If

        If leftPart < rightPart Then
            LoaderUpdates_CompareVersions = -1
            Exit Function
        End If
    Next index
End Function
