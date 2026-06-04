Attribute VB_Name = "LoaderLogging"
Option Explicit

Public Sub LoaderLogging_Write(ByVal message As String)
    On Error Resume Next

    Dim fileNumber As Integer
    LoaderPaths_EnsureBaseFolders

    fileNumber = FreeFile
    Open LoaderPaths_LogPath() For Append As #fileNumber
    Print #fileNumber, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & message
    Close #fileNumber
End Sub
