Attribute VB_Name = "LoaderPaths"
Option Explicit

Public Function LoaderPaths_BaseDir() As String
    Dim rootPath As String
    rootPath = Environ$("LOCALAPPDATA")
    If LenB(rootPath) = 0 Then rootPath = Environ$("APPDATA")
    LoaderPaths_BaseDir = rootPath & "\" & VZID_APP_NAME & "\"
End Function

Public Function LoaderPaths_LoaderDir() As String
    LoaderPaths_LoaderDir = LoaderPaths_BaseDir() & "loader\"
End Function

Public Function LoaderPaths_VersionsDir() As String
    LoaderPaths_VersionsDir = LoaderPaths_BaseDir() & "versions\"
End Function

Public Function LoaderPaths_CurrentDir() As String
    LoaderPaths_CurrentDir = LoaderPaths_VersionsDir() & "current\"
End Function

Public Function LoaderPaths_PendingDir() As String
    LoaderPaths_PendingDir = LoaderPaths_VersionsDir() & "pending\"
End Function

Public Function LoaderPaths_ConfigDir() As String
    LoaderPaths_ConfigDir = LoaderPaths_BaseDir() & "config\"
End Function

Public Function LoaderPaths_LogsDir() As String
    LoaderPaths_LogsDir = LoaderPaths_BaseDir() & "logs\"
End Function

Public Function LoaderPaths_MainAddinPath() As String
    LoaderPaths_MainAddinPath = LoaderPaths_CurrentDir() & VZID_MAIN_ADDIN_FILE
End Function

Public Function LoaderPaths_ConfigPath() As String
    LoaderPaths_ConfigPath = LoaderPaths_ConfigDir() & VZID_CONFIG_FILE
End Function

Public Function LoaderPaths_LogPath() As String
    LoaderPaths_LogPath = LoaderPaths_LogsDir() & VZID_LOG_FILE
End Function

Public Function LoaderPaths_PendingMainPathForVersion(ByVal versionText As String) As String
    LoaderPaths_PendingMainPathForVersion = LoaderPaths_PendingDir() & "MainVZID-" & LoaderPaths_SanitizeFileToken(versionText) & ".xlam"
End Function

Public Sub LoaderPaths_EnsureBaseFolders()
    LoaderPaths_EnsureFolder LoaderPaths_BaseDir()
    LoaderPaths_EnsureFolder LoaderPaths_LoaderDir()
    LoaderPaths_EnsureFolder LoaderPaths_VersionsDir()
    LoaderPaths_EnsureFolder LoaderPaths_CurrentDir()
    LoaderPaths_EnsureFolder LoaderPaths_PendingDir()
    LoaderPaths_EnsureFolder LoaderPaths_ConfigDir()
    LoaderPaths_EnsureFolder LoaderPaths_LogsDir()
End Sub

Private Sub LoaderPaths_EnsureFolder(ByVal fullPath As String)
    On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(fullPath) Then
        fso.CreateFolder fullPath
    End If
End Sub

Private Function LoaderPaths_SanitizeFileToken(ByVal value As String) As String
    Dim resultText As String
    Dim index As Long
    Dim ch As String

    resultText = Trim$(value)
    If LenB(resultText) = 0 Then resultText = "pending"

    For index = 1 To Len(resultText)
        ch = Mid$(resultText, index, 1)
        Select Case ch
            Case "A" To "Z", "a" To "z", "0" To "9", "-", "_", "."
                LoaderPaths_SanitizeFileToken = LoaderPaths_SanitizeFileToken & ch
            Case Else
                LoaderPaths_SanitizeFileToken = LoaderPaths_SanitizeFileToken & "_"
        End Select
    Next index
End Function
