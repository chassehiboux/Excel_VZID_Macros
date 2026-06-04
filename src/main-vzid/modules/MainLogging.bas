Attribute VB_Name = "MainLogging"
Option Explicit

Public Sub MainLogging_Write(ByVal message As String)
    On Error Resume Next

    Dim fileNumber As Integer
    MainConfig_EnsureBaseFolders

    fileNumber = FreeFile
    Open MainConfig_LogPath() For Append As #fileNumber
    Print #fileNumber, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & message
    Close #fileNumber
End Sub
