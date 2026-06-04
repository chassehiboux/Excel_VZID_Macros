Attribute VB_Name = "LoaderAddinHost"
Option Explicit

Public Sub LoaderAddinHost_LoadMainAddin(ByVal fullPath As String)
    On Error GoTo failed

    If LenB(Dir$(fullPath)) = 0 Then
        LoaderLogging_Write "Main add-in file not found: " & fullPath
        Exit Sub
    End If

    Dim addinRef As AddIn
    Set addinRef = LoaderAddinHost_FindByPath(fullPath)
    If addinRef Is Nothing Then
        Set addinRef = Application.AddIns.Add(Filename:=fullPath, CopyFile:=False)
    End If

    If Not addinRef Is Nothing Then
        If Not addinRef.Installed Then addinRef.Installed = True
        LoaderLogging_Write "Main add-in connected: " & fullPath
    End If
    Exit Sub

failed:
    LoaderLogging_Write "LoaderAddinHost_LoadMainAddin failed: " & Err.Number & " - " & Err.Description
End Sub

Public Sub LoaderAddinHost_UnloadByPath(ByVal fullPath As String)
    On Error Resume Next

    Dim addinRef As AddIn
    Set addinRef = LoaderAddinHost_FindByPath(fullPath)
    If Not addinRef Is Nothing Then
        If addinRef.Installed Then addinRef.Installed = False
    End If
End Sub

Public Function LoaderAddinHost_FindByPath(ByVal fullPath As String) As AddIn
    Dim index As Long

    For index = 1 To Application.AddIns.Count
        If StrComp(Application.AddIns(index).FullName, fullPath, vbTextCompare) = 0 Then
            Set LoaderAddinHost_FindByPath = Application.AddIns(index)
            Exit Function
        End If
    Next index
End Function
