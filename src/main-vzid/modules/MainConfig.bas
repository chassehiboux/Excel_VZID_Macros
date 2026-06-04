Attribute VB_Name = "MainConfig"
Option Explicit

Private Const MAIN_APP_NAME As String = "VZID"
Private Const MAIN_CONFIG_FILE As String = "config.json"
Private Const MAIN_CONFIG_MANIFEST_URL As String = "https://github.com/chassehiboux/Excel_VZID_Macros/releases/latest/download/manifest.json"

Public Sub MainConfig_EnsureConfig()
    If LenB(Dir$(MainConfig_ConfigPath())) > 0 Then Exit Sub

    MainConfig_EnsureBaseFolders
    MainConfig_SaveText MainConfig_DefaultJsonText()
End Sub

Public Sub MainConfig_EnsureBaseFolders()
    MainConfig_EnsureFolder MainConfig_BaseDir()
    MainConfig_EnsureFolder MainConfig_VersionsDir()
    MainConfig_EnsureFolder MainConfig_CurrentDir()
    MainConfig_EnsureFolder MainConfig_PendingDir()
    MainConfig_EnsureFolder MainConfig_ConfigDir()
    MainConfig_EnsureFolder MainConfig_LogsDir()
End Sub

Public Function MainConfig_BaseDir() As String
    Dim rootPath As String
    rootPath = Environ$("LOCALAPPDATA")
    If LenB(rootPath) = 0 Then rootPath = Environ$("APPDATA")
    MainConfig_BaseDir = rootPath & "\" & MAIN_APP_NAME & "\"
End Function

Public Function MainConfig_VersionsDir() As String
    MainConfig_VersionsDir = MainConfig_BaseDir() & "versions\"
End Function

Public Function MainConfig_CurrentDir() As String
    MainConfig_CurrentDir = MainConfig_VersionsDir() & "current\"
End Function

Public Function MainConfig_PendingDir() As String
    MainConfig_PendingDir = MainConfig_VersionsDir() & "pending\"
End Function

Public Function MainConfig_ConfigDir() As String
    MainConfig_ConfigDir = MainConfig_BaseDir() & "config\"
End Function

Public Function MainConfig_LogsDir() As String
    MainConfig_LogsDir = MainConfig_BaseDir() & "logs\"
End Function

Public Function MainConfig_ConfigPath() As String
    MainConfig_ConfigPath = MainConfig_ConfigDir() & MAIN_CONFIG_FILE
End Function

Public Function MainConfig_PendingMainPath() As String
    MainConfig_PendingMainPath = MainConfig_PendingDir() & "MainVZID.xlam"
End Function

Public Function MainConfig_DefaultJsonText() As String
    MainConfig_DefaultJsonText = Join(Array( _
        "{", _
        "  ""schemaVersion"": ""1"",", _
        "  ""repoUrl"": ""https://github.com/chassehiboux/Excel_VZID_Macros"",", _
        "  ""manifestUrl"": """ & MAIN_CONFIG_MANIFEST_URL & """,", _
        "  ""selectedRegion"": ""KGN"",", _
        "  ""activeMainVersion"": ""0.1.1"",", _
        "  ""activeLoaderVersion"": ""0.1.1"",", _
        "  ""availableMainVersion"": """",", _
        "  ""availableMainDownloadUrl"": """",", _
        "  ""pendingMainVersion"": """",", _
        "  ""pendingMainPath"": """",", _
        "  ""lastUpdateCheckAt"": """",", _
        "  ""lastUpdateStatus"": ""never"",", _
        "  ""lastUpdateMessage"": ""Проверка обновлений еще не выполнялась."",", _
        "  ""fullAccessUsersCsv"": ""dzuikevich"",", _
        "  ""commandAccessDocPacketsCsv"": ""*"",", _
        "  ""commandAccessMakeCoverLettersCsv"": ""dzuikevich"",", _
        "  ""commandAccessMakeCoverLettersByRecipientCsv"": ""dzuikevich"",", _
        "  ""commandAccessRicZakaznyeCreateCsv"": ""dzuikevich""", _
        "}"), vbCrLf)
End Function

Public Function MainConfig_LoadText() As String
    On Error GoTo failed

    MainConfig_EnsureConfig

    Dim fileNumber As Integer
    fileNumber = FreeFile
    Open MainConfig_ConfigPath() For Input As #fileNumber
    MainConfig_LoadText = Input$(LOF(fileNumber), fileNumber)
    Close #fileNumber
    Exit Function

failed:
    On Error Resume Next
    Close #fileNumber
    MainConfig_LoadText = MainConfig_DefaultJsonText()
End Function

Public Sub MainConfig_SaveText(ByVal jsonText As String)
    On Error Resume Next

    MainConfig_EnsureBaseFolders

    Dim fileNumber As Integer
    fileNumber = FreeFile
    Open MainConfig_ConfigPath() For Output As #fileNumber
    Print #fileNumber, jsonText
    Close #fileNumber
End Sub

Public Function MainConfig_ReadValue(ByVal key As String, Optional ByVal defaultValue As String = "") As String
    MainConfig_ReadValue = MainConfig_ReadJsonString(MainConfig_LoadText(), key, defaultValue)
End Function

Public Sub MainConfig_WriteValue(ByVal key As String, ByVal newValue As String)
    Dim jsonText As String
    jsonText = MainConfig_LoadText()
    MainConfig_SaveText MainConfig_UpsertStringValue(jsonText, key, newValue)
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

    token = """" & key & """"
    keyPos = InStr(1, jsonText, token, vbTextCompare)
    If keyPos = 0 Then
        MainConfig_UpsertStringValue = jsonText
        Exit Function
    End If

    colonPos = InStr(keyPos + Len(token), jsonText, ":")
    quoteStart = InStr(colonPos + 1, jsonText, """")
    quoteEnd = InStr(quoteStart + 1, jsonText, """")

    If colonPos = 0 Or quoteStart = 0 Or quoteEnd <= quoteStart Then
        MainConfig_UpsertStringValue = jsonText
        Exit Function
    End If

    MainConfig_UpsertStringValue = Left$(jsonText, quoteStart) & MainConfig_EscapeJsonString(newValue) & Mid$(jsonText, quoteEnd)
End Function

Private Function MainConfig_EscapeJsonString(ByVal value As String) As String
    Dim escaped As String
    escaped = Replace(value, "\", "\\")
    escaped = Replace(escaped, """", "'")
    MainConfig_EscapeJsonString = escaped
End Function

Private Sub MainConfig_EnsureFolder(ByVal fullPath As String)
    On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(fullPath) Then
        fso.CreateFolder fullPath
    End If
End Sub
