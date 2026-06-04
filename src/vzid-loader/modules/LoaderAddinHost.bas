Attribute VB_Name = "LoaderAddinHost"
Option Explicit

Public Sub LoaderAddinHost_LoadMainAddin(ByVal fullPath As String)
    On Error GoTo failed

    If LenB(Dir$(fullPath)) = 0 Then
        LoaderLogging_Write "Main add-in file not found: " & fullPath
        Exit Sub
    End If

    Dim workbookRef As Workbook
    Set workbookRef = LoaderAddinHost_FindWorkbookByPath(fullPath)
    If Not workbookRef Is Nothing Then
        LoaderLogging_Write "Main add-in already open: " & fullPath
        Exit Sub
    End If

    Set workbookRef = LoaderAddinHost_FindWorkbookByName(Dir$(fullPath))
    If Not workbookRef Is Nothing Then
        If workbookRef.IsAddin Then
            LoaderLogging_Write "Closing stale add-in copy: " & workbookRef.FullName
            workbookRef.Close SaveChanges:=False
            DoEvents
        Else
            LoaderLogging_Write "Workbook with same name is already open and visible: " & workbookRef.FullName
            Exit Sub
        End If
    End If

    Set workbookRef = Application.Workbooks.Open(Filename:=fullPath, ReadOnly:=True, AddToMru:=False)
    If workbookRef Is Nothing Then Exit Sub

    On Error Resume Next
    workbookRef.IsAddin = True
    On Error GoTo failed
    LoaderLogging_Write "Main add-in connected: " & fullPath
    Exit Sub

failed:
    LoaderLogging_Write "LoaderAddinHost_LoadMainAddin failed: " & Err.Number & " - " & Err.Description
End Sub

Public Sub LoaderAddinHost_UnloadByPath(ByVal fullPath As String)
    On Error Resume Next

    Dim workbookRef As Workbook
    Set workbookRef = LoaderAddinHost_FindWorkbookByPath(fullPath)
    If Not workbookRef Is Nothing Then
        workbookRef.Close SaveChanges:=False
    End If

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

Public Function LoaderAddinHost_FindWorkbookByPath(ByVal fullPath As String) As Workbook
    Dim index As Long

    For index = 1 To Application.Workbooks.Count
        If StrComp(Application.Workbooks(index).FullName, fullPath, vbTextCompare) = 0 Then
            Set LoaderAddinHost_FindWorkbookByPath = Application.Workbooks(index)
            Exit Function
        End If
    Next index
End Function

Public Function LoaderAddinHost_FindWorkbookByName(ByVal workbookName As String) As Workbook
    Dim index As Long

    For index = 1 To Application.Workbooks.Count
        If StrComp(Application.Workbooks(index).Name, workbookName, vbTextCompare) = 0 Then
            Set LoaderAddinHost_FindWorkbookByName = Application.Workbooks(index)
            Exit Function
        End If
    Next index
End Function
