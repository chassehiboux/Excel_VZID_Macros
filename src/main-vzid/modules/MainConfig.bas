Attribute VB_Name = "MainConfig"
Option Explicit

Public Sub MainConfig_EnsureConfig()
    If LenB(Dir$(MainConfig_ConfigPath())) > 0 Then Exit Sub

    MainConfig_EnsureBaseFolders
    MainConfig_SaveText MainConfig_DefaultJsonText()
End Sub

Public Sub MainConfig_EnsureBaseFolders()
    MainConfig_EnsureFolder MainConfig_BaseDir()
    MainConfig_EnsureFolder MainConfig_AddinDir()
    MainConfig_EnsureFolder MainConfig_UpdaterDir()
    MainConfig_EnsureFolder MainConfig_UpdatesDir()
    MainConfig_EnsureFolder MainConfig_BackupDir()
    MainConfig_EnsureFolder MainConfig_ConfigDir()
    MainConfig_EnsureFolder MainConfig_LogsDir()
End Sub

Public Function MainConfig_BaseDir() As String
    Dim rootPath As String
    rootPath = Environ$("APPDATA")
    If LenB(rootPath) = 0 Then rootPath = Environ$("LOCALAPPDATA")
    MainConfig_BaseDir = rootPath & "\Microsoft\Excel\LocalCache\" & VZID_APP_NAME & "\"
End Function

Public Function MainConfig_AddinDir() As String
    MainConfig_AddinDir = MainConfig_BaseDir() & "addin\"
End Function

Public Function MainConfig_UpdaterDir() As String
    MainConfig_UpdaterDir = MainConfig_BaseDir() & "updater\"
End Function

Public Function MainConfig_UpdatesDir() As String
    MainConfig_UpdatesDir = MainConfig_BaseDir() & "updates\"
End Function

Public Function MainConfig_BackupDir() As String
    MainConfig_BackupDir = MainConfig_BaseDir() & "backup\"
End Function

Public Function MainConfig_ConfigDir() As String
    MainConfig_ConfigDir = MainConfig_BaseDir() & "config\"
End Function

Public Function MainConfig_LogsDir() As String
    MainConfig_LogsDir = MainConfig_BaseDir() & "logs\"
End Function

Public Function MainConfig_ConfigPath() As String
    MainConfig_ConfigPath = MainConfig_ConfigDir() & VZID_CONFIG_FILE
End Function

Public Function MainConfig_LogPath() As String
    MainConfig_LogPath = MainConfig_LogsDir() & VZID_LOG_FILE
End Function

Public Function MainConfig_InstalledAddinPath() As String
    MainConfig_InstalledAddinPath = MainConfig_AddinDir() & VZID_MAIN_ADDIN_FILE
End Function

Public Function MainConfig_UpdaterExePath() As String
    MainConfig_UpdaterExePath = MainConfig_UpdaterDir() & VZID_UPDATER_FILE
End Function

Public Function MainConfig_PreparedMainPathForVersion(ByVal versionText As String) As String
    MainConfig_PreparedMainPathForVersion = MainConfig_UpdatesDir() & "MainVZID-" & MainConfig_SanitizeFileToken(versionText) & ".xlam"
End Function

Public Function MainConfig_PreparedSetupPathForVersion(ByVal versionText As String) As String
    MainConfig_PreparedSetupPathForVersion = MainConfig_UpdatesDir() & "setup-" & MainConfig_SanitizeFileToken(versionText) & ".exe"
End Function

Public Function MainConfig_DefaultJsonText() As String
    MainConfig_DefaultJsonText = MainConfig_JoinLines( _
        "{", _
        "  ""schemaVersion"": ""2"",", _
        "  ""repoUrl"": ""https://github.com/chassehiboux/Excel_VZID_Macros"",", _
        "  ""manifestUrl"": """ & VZID_MANIFEST_URL & """,", _
        "  ""selectedRegion"": """ & VZID_DEFAULT_REGION & """,", _
        "  ""activeMainVersion"": """ & VZID_MAIN_VERSION & """,", _
        "  ""activeUpdaterVersion"": """ & VZID_UPDATER_VERSION & """,", _
        "  ""availableMainVersion"": """",", _
        "  ""availableMainDownloadUrl"": """"," _
    )
    MainConfig_DefaultJsonText = MainConfig_DefaultJsonText & vbCrLf & MainConfig_JoinLines( _
        "  ""preparedMainVersion"": """",", _
        "  ""preparedMainPath"": """",", _
        "  ""lastUpdateCheckAt"": """",", _
        "  ""lastUpdateStatus"": ""never"",", _
        "  ""lastUpdateMessage"": ""Проверка обновлений еще не выполнялась."",", _
        "  ""fullAccessUsersCsv"": ""dzuikevich"",", _
        "  ""commandAccessDocPacketsCsv"": ""*"",", _
        "  ""commandAccessMakeCoverLettersCsv"": ""*"",", _
        "  ""commandAccessMakeCoverLettersByRecipientCsv"": ""*""," _
    )
    MainConfig_DefaultJsonText = MainConfig_DefaultJsonText & vbCrLf & MainConfig_JoinLines( _
        "  ""commandAccessZakaznyeCreateCsv"": ""*"",", _
        "  ""commandAccessManualProcReportCsv"": ""*"",", _
        "  ""commandAccessPdfSignerCsv"": ""*"",", _
        "  ""commandAccessPdfScannerCsv"": ""*"",", _
        "  ""commandAccessOspSelectCsv"": ""*"",", _
        "  ""commandAccessHotkeySetupCsv"": ""*""", _
        "}" _
    )
End Function

Public Function MainConfig_LoadText() As String
    On Error GoTo failed

    MainConfig_EnsureConfig

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile MainConfig_ConfigPath()
    MainConfig_LoadText = stream.ReadText(-1)
    stream.Close
    Exit Function

failed:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    MainConfig_LoadText = MainConfig_DefaultJsonText()
End Function

Public Sub MainConfig_SaveText(ByVal jsonText As String)
    On Error Resume Next

    MainConfig_EnsureBaseFolders

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText jsonText
    stream.SaveToFile MainConfig_ConfigPath(), 2
    stream.Close
End Sub

Public Function MainConfig_ReadValue(ByVal key As String, Optional ByVal defaultValue As String = "") As String
    MainConfig_ReadValue = MainConfig_ReadJsonString(MainConfig_LoadText(), key, defaultValue)
End Function

Public Sub MainConfig_WriteValue(ByVal key As String, ByVal newValue As String)
    Dim jsonText As String
    jsonText = MainConfig_LoadText()
    MainConfig_SaveText MainConfig_UpsertStringValue(jsonText, key, newValue)
End Sub

Public Sub MainConfig_WriteUpdateState(ByVal statusValue As String, ByVal messageValue As String, ByVal availableVersion As String, ByVal downloadUrl As String)
    MainConfig_WriteValue "lastUpdateCheckAt", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    MainConfig_WriteValue "lastUpdateStatus", statusValue
    MainConfig_WriteValue "lastUpdateMessage", messageValue
    MainConfig_WriteValue "availableMainVersion", availableVersion
    MainConfig_WriteValue "availableMainDownloadUrl", downloadUrl
End Sub

Public Function MainConfig_ReadJsonString(ByVal jsonText As String, ByVal key As String, Optional ByVal defaultValue As String = "") As String
    Dim token As String
    Dim keyPos As Long
    Dim colonPos As Long
    Dim quoteStart As Long
    Dim quoteEnd As Long

    token = """" & key & """"
    keyPos = InStr(1, jsonText, token, vbTextCompare)
    If keyPos = 0 Then
        MainConfig_ReadJsonString = defaultValue
        Exit Function
    End If

    colonPos = InStr(keyPos + Len(token), jsonText, ":")
    quoteStart = InStr(colonPos + 1, jsonText, """")
    quoteEnd = InStr(quoteStart + 1, jsonText, """")

    If colonPos = 0 Or quoteStart = 0 Or quoteEnd <= quoteStart Then
        MainConfig_ReadJsonString = defaultValue
        Exit Function
    End If

    MainConfig_ReadJsonString = Mid$(jsonText, quoteStart + 1, quoteEnd - quoteStart - 1)
    MainConfig_ReadJsonString = Replace(MainConfig_ReadJsonString, "\\", "\")
End Function

Public Function MainConfig_CopyFile(ByVal sourcePath As String, ByVal targetPath As String) As Boolean
    On Error GoTo failed

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    MainConfig_EnsureFolder fso.GetParentFolderName(targetPath)

    If fso.FileExists(targetPath) Then
        fso.DeleteFile targetPath, True
    End If

    fso.CopyFile sourcePath, targetPath, True
    MainConfig_CopyFile = fso.FileExists(targetPath)
    Exit Function

failed:
    MainConfig_CopyFile = False
End Function

Private Function MainConfig_UpsertStringValue(ByVal jsonText As String, ByVal key As String, ByVal newValue As String) As String
    Dim token As String
    Dim keyPos As Long
    Dim colonPos As Long
    Dim quoteStart As Long
    Dim quoteEnd As Long
    Dim insertPos As Long
    Dim prefix As String
    Dim escapedValue As String

    token = """" & key & """"
    keyPos = InStr(1, jsonText, token, vbTextCompare)
    escapedValue = MainConfig_EscapeJsonString(newValue)

    If keyPos = 0 Then
        insertPos = InStrRev(jsonText, "}")
        If insertPos = 0 Then
            MainConfig_UpsertStringValue = jsonText
            Exit Function
        End If

        prefix = Left$(jsonText, insertPos - 1)
        prefix = RTrim$(prefix)
        If Right$(prefix, 1) <> "{" And Right$(prefix, 1) <> "," Then
            prefix = prefix & ","
        End If

        MainConfig_UpsertStringValue = prefix & vbCrLf & "  """ & key & """: """ & escapedValue & """" & vbCrLf & "}"
        Exit Function
    End If

    colonPos = InStr(keyPos + Len(token), jsonText, ":")
    quoteStart = InStr(colonPos + 1, jsonText, """")
    quoteEnd = InStr(quoteStart + 1, jsonText, """")

    If colonPos = 0 Or quoteStart = 0 Or quoteEnd <= quoteStart Then
        MainConfig_UpsertStringValue = jsonText
        Exit Function
    End If

    MainConfig_UpsertStringValue = Left$(jsonText, quoteStart) & escapedValue & Mid$(jsonText, quoteEnd)
End Function

Private Function MainConfig_EscapeJsonString(ByVal value As String) As String
    Dim escaped As String
    escaped = Replace(value, "\", "\\")
    escaped = Replace(escaped, """", "'")
    MainConfig_EscapeJsonString = escaped
End Function

Private Function MainConfig_JoinLines(ParamArray values() As Variant) As String
    Dim index As Long

    For index = LBound(values) To UBound(values)
        If index > LBound(values) Then
            MainConfig_JoinLines = MainConfig_JoinLines & vbCrLf
        End If
        MainConfig_JoinLines = MainConfig_JoinLines & CStr(values(index))
    Next index
End Function

Private Sub MainConfig_EnsureFolder(ByVal fullPath As String)
    On Error Resume Next

    If LenB(fullPath) = 0 Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(fullPath) Then
        fso.CreateFolder fullPath
    End If
End Sub

Private Function MainConfig_SanitizeFileToken(ByVal value As String) As String
    Dim resultText As String
    Dim index As Long
    Dim ch As String

    resultText = Trim$(value)
    If LenB(resultText) = 0 Then resultText = "prepared"

    For index = 1 To Len(resultText)
        ch = Mid$(resultText, index, 1)
        Select Case ch
            Case "A" To "Z", "a" To "z", "0" To "9", "-", "_", "."
                MainConfig_SanitizeFileToken = MainConfig_SanitizeFileToken & ch
            Case Else
                MainConfig_SanitizeFileToken = MainConfig_SanitizeFileToken & "_"
        End Select
    Next index
End Function
