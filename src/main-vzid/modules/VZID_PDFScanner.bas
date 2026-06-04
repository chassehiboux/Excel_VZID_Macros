Attribute VB_Name = "PDFScanner"
Option Explicit

Private Const APP_FOLDER_NAME As String = "PDFScanner"
Private Const APP_EXE_NAME As String = "PDFScanner.exe"
Private Const LOADER_FOLDER_NAME As String = "PDFScannerLoader"
Private Const LOADER_EXE_NAME As String = "PDFScannerLoader.exe"

Public Sub RunPDFScanner()
    Dim localRoot As String
    Dim localAppDir As String
    Dim localExePath As String
    Dim loaderExePath As String
    Dim sourceDir As String

    On Error GoTo ErrHandler

    localRoot = Environ$("AppData") & "\Microsoft\Excel\LocalCache"
    localAppDir = localRoot & "\" & APP_FOLDER_NAME
    localExePath = localAppDir & "\" & APP_EXE_NAME

    EnsureFolderExists localRoot

    If FileExists(localExePath) Then
        LaunchExecutable localExePath
        Exit Sub
    End If

    sourceDir = GetFirstAvailableSource()
    If Len(sourceDir) = 0 Then
        MsgBox "Не удалось найти доступную сетевую папку PDFScanner.", vbExclamation, "PDFScanner"
        Exit Sub
    End If

    loaderExePath = sourceDir & "\" & LOADER_FOLDER_NAME & "\" & LOADER_EXE_NAME
    If FileExists(loaderExePath) Then
        LaunchExecutable loaderExePath
        Exit Sub
    End If

    DownloadPDFScanner sourceDir, localAppDir

    If Not FileExists(localExePath) Then
        MsgBox "Файл программы не найден: " & localExePath, vbCritical, "PDFScanner"
        Exit Sub
    End If

    LaunchExecutable localExePath
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox Err.Description, vbCritical, "PDFScanner"
End Sub

Private Function GetFirstAvailableSource() As String
    Dim sources As Variant
    Dim item As Variant

    sources = Array( _
        "\\corp.vostok-electra.ru\Kgn\Отделы\Отдел взыскания по исполнительным документам\Зуйкевич Данил Иванович\Excel\ВЗИД\Extensions\PDFScanner", _
        "\\corp.vostok-electra.ru\Tmn\Общая\ОВЗИД\Зуйкевич\ВЗИД\Extensions\PDFScanner", _
        "\\Ekb-vpfs01\екатеринбург\Отделы\Отдел взыскания задолженности по исполнительным документам\26. Макрос\ВЗИД\Extensions\PDFScanner" _
    )

    For Each item In sources
        If FolderExists(CStr(item)) Then
            GetFirstAvailableSource = CStr(item)
            Exit Function
        End If
    Next item
End Function

Private Sub DownloadPDFScanner(ByVal sourceDir As String, ByVal targetDir As String)
    Dim shell As Object
    Dim command As String
    Dim exitCode As Long
    Dim expectedExePath As String

    Set shell = CreateObject("WScript.Shell")
    expectedExePath = targetDir & "\" & APP_EXE_NAME

    command = "cmd /c robocopy " & Quote(sourceDir) & " " & Quote(targetDir) & _
              " /E /MT:16 /R:1 /W:1 /NFL /NDL /NJH /NJS /NP /XD " & Quote(sourceDir & "\" & LOADER_FOLDER_NAME)

    Application.StatusBar = "Загрузка PDFScanner..."
    DoEvents

    exitCode = shell.Run(command, 0, True)
    Application.StatusBar = False

    If exitCode >= 8 Then
        Err.Raise vbObjectError + 1000, "DownloadPDFScanner", _
                  "Не удалось скопировать PDFScanner. Код robocopy: " & CStr(exitCode)
    End If

    If Not WaitForFile(expectedExePath, 10) Then
        Err.Raise vbObjectError + 1001, "DownloadPDFScanner", _
                  "После копирования не найден локальный файл: " & expectedExePath
    End If
End Sub

Private Sub LaunchExecutable(ByVal exePath As String)
    Dim shellApp As Object
    Dim parentFolder As String

    parentFolder = Left$(exePath, InStrRev(exePath, "\") - 1)

    Set shellApp = CreateObject("Shell.Application")
    shellApp.ShellExecute exePath, "", parentFolder, "open", 1
End Sub

Private Sub EnsureFolderExists(ByVal folderPath As String)
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
End Sub

Private Function FolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (Len(Dir$(folderPath, vbDirectory)) > 0)
    On Error GoTo 0
End Function

Private Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir$(filePath, vbNormal)) > 0)
    On Error GoTo 0
End Function

Private Function Quote(ByVal value As String) As String
    Quote = Chr$(34) & value & Chr$(34)
End Function

Private Function WaitForFile(ByVal filePath As String, ByVal timeoutSeconds As Long) As Boolean
    Dim startedAt As Single

    startedAt = Timer

    Do
        If FileExists(filePath) Then
            WaitForFile = True
            Exit Function
        End If

        DoEvents
        Application.Wait Now + TimeSerial(0, 0, 1)

        If Timer < startedAt Then
            startedAt = startedAt - 86400!
        End If
    Loop While Timer - startedAt < timeoutSeconds
End Function


