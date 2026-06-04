Attribute VB_Name = "LoaderConfig"
Option Explicit

Public Sub LoaderConfig_EnsureConfig()
    If LenB(Dir$(LoaderPaths_ConfigPath())) > 0 Then Exit Sub

    LoaderPaths_EnsureBaseFolders
    LoaderConfig_SaveText LoaderConfig_DefaultJsonText()
End Sub

Public Function LoaderConfig_DefaultJsonText() As String
    LoaderConfig_DefaultJsonText = Join(Array( _
        "{", _
        "  ""schemaVersion"": ""1"",", _
        "  ""repoUrl"": ""https://github.com/chassehiboux/Excel_VZID_Macros"",", _
        "  ""manifestUrl"": """ & VZID_MANIFEST_URL & """,", _
        "  ""selectedRegion"": """ & VZID_DEFAULT_REGION & """,", _
        "  ""activeMainVersion"": ""0.1.1"",", _
        "  ""activeLoaderVersion"": """ & VZID_LOADER_VERSION & """,", _
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

Public Function LoaderConfig_LoadText() As String
    On Error GoTo failed

    LoaderConfig_EnsureConfig

    Dim fileNumber As Integer
    fileNumber = FreeFile
    Open LoaderPaths_ConfigPath() For Input As #fileNumber
    LoaderConfig_LoadText = Input$(LOF(fileNumber), fileNumber)
    Close #fileNumber
    Exit Function

failed:
    On Error Resume Next
    Close #fileNumber
    LoaderConfig_LoadText = LoaderConfig_DefaultJsonText()
End Function

Public Sub LoaderConfig_SaveText(ByVal jsonText As String)
    On Error Resume Next

    LoaderPaths_EnsureBaseFolders

    Dim fileNumber As Integer
    fileNumber = FreeFile
    Open LoaderPaths_ConfigPath() For Output As #fileNumber
    Print #fileNumber, jsonText
    Close #fileNumber
End Sub

Public Function LoaderConfig_ReadValue(ByVal key As String, Optional ByVal defaultValue As String = "") As String
    LoaderConfig_ReadValue = LoaderConfig_ReadJsonString(LoaderConfig_LoadText(), key, defaultValue)
End Function

Public Sub LoaderConfig_WriteValue(ByVal key As String, ByVal newValue As String)
    Dim jsonText As String
    jsonText = LoaderConfig_LoadText()
    LoaderConfig_SaveText LoaderConfig_UpsertStringValue(jsonText, key, newValue)
End Sub

Public Sub LoaderConfig_WriteUpdateState(ByVal statusValue As String, ByVal messageValue As String, ByVal availableVersion As String, ByVal downloadUrl As String)
    LoaderConfig_WriteValue "lastUpdateCheckAt", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    LoaderConfig_WriteValue "lastUpdateStatus", statusValue
    LoaderConfig_WriteValue "lastUpdateMessage", messageValue
    LoaderConfig_WriteValue "availableMainVersion", availableVersion
    LoaderConfig_WriteValue "availableMainDownloadUrl", downloadUrl
End Sub

Public Function LoaderConfig_ReadJsonString(ByVal jsonText As String, ByVal key As String, Optional ByVal defaultValue As String = "") As String
    Dim token As String
    Dim keyPos As Long
    Dim colonPos As Long
    Dim quoteStart As Long
    Dim quoteEnd As Long

    token = """" & key & """"
    keyPos = InStr(1, jsonText, token, vbTextCompare)
    If keyPos = 0 Then
        LoaderConfig_ReadJsonString = defaultValue
        Exit Function
    End If

    colonPos = InStr(keyPos + Len(token), jsonText, ":")
    quoteStart = InStr(colonPos + 1, jsonText, """")
    quoteEnd = InStr(quoteStart + 1, jsonText, """")

    If colonPos = 0 Or quoteStart = 0 Or quoteEnd <= quoteStart Then
        LoaderConfig_ReadJsonString = defaultValue
        Exit Function
    End If

    LoaderConfig_ReadJsonString = Mid$(jsonText, quoteStart + 1, quoteEnd - quoteStart - 1)
    LoaderConfig_ReadJsonString = Replace(LoaderConfig_ReadJsonString, "\\", "\")
End Function

Private Function LoaderConfig_UpsertStringValue(ByVal jsonText As String, ByVal key As String, ByVal newValue As String) As String
    Dim token As String
    Dim keyPos As Long
    Dim colonPos As Long
    Dim quoteStart As Long
    Dim quoteEnd As Long

    token = """" & key & """"
    keyPos = InStr(1, jsonText, token, vbTextCompare)
    If keyPos = 0 Then
        LoaderConfig_UpsertStringValue = jsonText
        Exit Function
    End If

    colonPos = InStr(keyPos + Len(token), jsonText, ":")
    quoteStart = InStr(colonPos + 1, jsonText, """")
    quoteEnd = InStr(quoteStart + 1, jsonText, """")

    If colonPos = 0 Or quoteStart = 0 Or quoteEnd <= quoteStart Then
        LoaderConfig_UpsertStringValue = jsonText
        Exit Function
    End If

    LoaderConfig_UpsertStringValue = Left$(jsonText, quoteStart) & LoaderConfig_EscapeJsonString(newValue) & Mid$(jsonText, quoteEnd)
End Function

Private Function LoaderConfig_EscapeJsonString(ByVal value As String) As String
    Dim escaped As String
    escaped = Replace(value, "\", "\\")
    escaped = Replace(escaped, """", "'")
    LoaderConfig_EscapeJsonString = escaped
End Function
