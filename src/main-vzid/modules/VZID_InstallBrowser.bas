Attribute VB_Name = "InstallBrowser"
Option Explicit

' ==========================================================================================
' === КОНСТАНТЫ ПУТЕЙ ======================================================================
' ==========================================================================================
Private Const FILE_EXT_TSV As String = "users_extensions.tsv"
Private Const FILE_PROFILE_TSV As String = "users_profile.tsv"
Private Const FILE_VER As String = "versionVZID.txt"
Private Const INSTALLER_VER_FILE As String = "version.txt"
Private Const MANIFEST_FILE As String = "manifest.json"

' Сетевые пути (корни)
Private Const PATH_ROOT_1 As String = "\\corp.vostok-electra.ru\Kgn\Отделы\Отдел взыскания по исполнительным документам\Зуйкевич Данил Иванович\Excel"
Private Const PATH_ROOT_2 As String = "\\corp.vostok-electra.ru\Tmn\Общая\ОВЗИД\Зуйкевич"
Private Const PATH_ROOT_3 As String = "\\Ekb-vpfs01\екатеринбург\Отделы\Отдел взыскания задолженности по исполнительным документам\26. Макрос"

Private Const SUFFIX_EXT As String = "\ВЗИД\Extensions"

Public ProcessingResultBrowser As String
Private glbCopiedCount As Long

' ==========================================================================================
' === ГЛАВНАЯ ТОЧКА ВХОДА ==================================================================
' ==========================================================================================
Public Sub Main_RunAll_Checks()
    On Error GoTo FinalCleanup

    ' ВАЖНО:
    ' frmLoading НЕ показываем тут.
    ' Он будет показан ТОЛЬКО:
    ' 1) когда реально обновляются install_*.exe (extension_installer)
    ' 2) когда реально запускается установка расширения (install_*.exe)

    ' 1) Обновляем файлы (копируем новые версии установщиков и/или StageTimer)
    Run_BrowserUpdate_Logic

    ' 2) Логируем запуск (AD запрос + запись в TSV)
    Log_Extension_Action "", "" ' просто обновляет дату

    ' 3) Сбор профиля из AD и запись
    CollectAndWrite_Profile_ToTsv

    ' 4) Автоматическая проверка и установка расширений
    AutoInstall_Extensions

FinalCleanup:
    On Error Resume Next
    Application.StatusBar = False
    Application.ScreenUpdating = True
    If frmLoading.Visible Then Unload frmLoading
End Sub

' ==========================================================================================
' === frmLoading: ПОКАЗ / СКРЫТИЕ ==========================================================
' ==========================================================================================
' ВАЖНО: UpdateLoadingStatus НЕ имеет права показывать форму.
' Форму показываем только через EnsureLoadingShown и только в нужных местах.

Private Sub EnsureLoadingShown(Optional ByVal firstText As String = vbNullString)
    On Error Resume Next
    If Not frmLoading.Visible Then frmLoading.Show vbModeless
    If Len(firstText) > 0 Then frmLoading.SetText firstText
    DoEvents
End Sub

Private Sub EnsureLoadingHidden()
    On Error Resume Next
    If frmLoading.Visible Then Unload frmLoading
End Sub

' Вспомогательная процедура для обновления текста на форме
Private Sub UpdateLoadingStatus(txt As String)
    ' НЕ показываем форму здесь!
    On Error Resume Next
    If frmLoading.Visible Then
        frmLoading.SetText txt
        DoEvents
    Else
        ' Можно оставить лёгкий фидбек в StatusBar без формы
        Application.StatusBar = txt
    End If
End Sub

' ==========================================================================================
' === АВТОМАТИЧЕСКАЯ УСТАНОВКА (С ПОДДЕРЖКОЙ ОТМЕНЫ) =======================================
' ==========================================================================================
Private Sub AutoInstall_Extensions()
    On Error Resume Next
    Dim login As String: login = Environ$("USERNAME")

    ' 1) Читаем статусы (без показа frmLoading)
    Dim statChrome As String, statEdge As String, statYandex As String
    GetExtensionStatus_MultiPath login, statChrome, statEdge, statYandex

    ' 2) Обрабатываем "Updated" (без принудительного показа формы)
    statChrome = ProcessPendingUpdate("Chrome", statChrome)
    statEdge = ProcessPendingUpdate("Edge", statEdge)
    statYandex = ProcessPendingUpdate("Yandex", statYandex)

    ' 3) Определяем наличие браузеров (без показа формы)
    Dim hasChrome As Boolean, hasEdge As Boolean, hasYandex As Boolean
    hasChrome = CheckBrowserExists("Chrome")
    hasEdge = CheckBrowserExists("Edge")
    hasYandex = CheckBrowserExists("Yandex")

    ' 4) Решаем: нужна ли вообще установка?
    Dim needInstallAny As Boolean
    needInstallAny = False
    If hasChrome And StrComp(statChrome, "Installed", vbTextCompare) <> 0 Then needInstallAny = True
    If hasEdge And StrComp(statEdge, "Installed", vbTextCompare) <> 0 Then needInstallAny = True
    If hasYandex And StrComp(statYandex, "Installed", vbTextCompare) <> 0 Then needInstallAny = True

    ' Если установка не нужна — frmLoading не показываем
    If Not needInstallAny Then Exit Sub

    ' === ВАЖНО: frmLoading показываем ТОЛЬКО тут, потому что дальше будет установка ===
    EnsureLoadingShown "Запуск модуля автоустановки..."
    frmLoading.IsCancelRequested = False
    frmLoading.btnCancel.Enabled = True
    UpdateLoadingStatus "Анализ состояния расширений..."

    ' 5) Логика по каждому браузеру
    If Not ProcessSingleBrowser("Chrome", hasChrome, statChrome) Then GoTo CANCEL_EXIT
    If Not ProcessSingleBrowser("Edge", hasEdge, statEdge) Then GoTo CANCEL_EXIT
    If Not ProcessSingleBrowser("Yandex", hasYandex, statYandex) Then GoTo CANCEL_EXIT

    UpdateLoadingStatus "Все проверки завершены."
    Application.Wait (Now + TimeValue("0:00:01"))
    EnsureLoadingHidden
    Exit Sub

CANCEL_EXIT:
    UpdateLoadingStatus "Операция прервана пользователем."
    Application.Wait (Now + TimeValue("0:00:02"))
    EnsureLoadingHidden
End Sub

' Возвращает False, если процесс был прерван пользователем
Private Function ProcessSingleBrowser(bName As String, exists As Boolean, currentStatus As String) As Boolean
    ProcessSingleBrowser = True ' По умолчанию продолжаем
    If Not exists Then Exit Function

    ' Проверка на отмену перед началом этапа
    If frmLoading.IsCancelRequested Then
        ProcessSingleBrowser = False
        Exit Function
    End If

    If StrComp(currentStatus, "Installed", vbTextCompare) = 0 Then
        UpdateLoadingStatus bName & ": Расширение уже установлено."
        Application.Wait (Now + TimeValue("0:00:01") / 4)
        Exit Function
    End If

    ' Запуск установки
    Dim res As Long
    res = RunInstallerWithStatus(bName, "install_" & LCase(bName) & ".exe")

    ' КОД 999 = ОТМЕНА ПОЛЬЗОВАТЕЛЕМ
    If res = 999 Then
        Log_Extension_Action bName, "Cancelled"
        ProcessSingleBrowser = False
        Exit Function
    End If

    Select Case res
        Case 0
            Log_Extension_Action bName, "Installed"
            OpenLinkInBrowser bName
        Case 10
            Log_Extension_Action bName, "Installed"
        Case 20
            Log_Extension_Action bName, "Failed"
        Case Else
            Log_Extension_Action bName, "Failed"
    End Select
End Function

Private Function RunInstallerWithStatus(browserName As String, exeName As String) As Long
    On Error Resume Next

    ' Тут установка — форму можно гарантированно держать открытой
    If Not frmLoading.Visible Then EnsureLoadingShown

    Dim localPath As String
    localPath = Environ$("APPDATA") & "\Microsoft\Excel\LocalCache\extension_installer\" & exeName

    If Dir(localPath) = "" Then
        RunInstallerWithStatus = -1
        Exit Function
    End If

    Dim statusFile As String
    statusFile = Environ$("TEMP") & "\install_status.txt"
    If Dir(statusFile) <> "" Then Kill statusFile

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")

    UpdateLoadingStatus "Запуск процесса " & exeName & "..." & vbCrLf & "Для остановки нажмите 'Отмена'"

    ' Отключаем TopMost, чтобы не мешать установке
    frmLoading.ToggleTopMost False

    ' Запускаем процесс
    Dim oExec As Object
    Set oExec = wsh.exec(Chr(34) & localPath & Chr(34))

    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim statusText As String

    ' === ЦИКЛ ОЖИДАНИЯ ===
    Do While oExec.status = 0
        ' 1) ПРОВЕРКА ОТМЕНЫ
        If frmLoading.IsCancelRequested Then
            wsh.Run "taskkill /F /T /PID " & oExec.ProcessID, 0, True
            RunInstallerWithStatus = 999
            frmLoading.ToggleTopMost True
            Set oExec = Nothing
            Set wsh = Nothing
            Exit Function
        End If

        ' 2) ЧТЕНИЕ СТАТУСА
        If fso.FileExists(statusFile) Then
            On Error Resume Next
            Set ts = fso.OpenTextFile(statusFile, 1, False, -1)
            statusText = ts.ReadAll
            ts.Close
            On Error Resume Next
        End If

        If Len(statusText) > 0 Then
            frmLoading.SetText "Установка в " & UCase(browserName) & "..." & vbCrLf & "Status: " & Right$(statusText, 100)
        Else
            frmLoading.SetText "Установка в " & UCase(browserName) & " (Запуск)..."
        End If

        DoEvents
        Application.Wait (Now + TimeValue("0:00:01") / 4)
    Loop

    RunInstallerWithStatus = oExec.exitCode

    frmLoading.ToggleTopMost True
    Set oExec = Nothing
    Set wsh = Nothing
    Set fso = Nothing
End Function

Private Sub GetExtensionStatus_MultiPath(login As String, ByRef stChrome As String, ByRef stEdge As String, ByRef stYandex As String)
    stChrome = "Not Installed": stEdge = "Not Installed": stYandex = "Not Installed"

    Dim paths(0 To 2) As String
    paths(0) = AddSlash_(PATH_ROOT_1) & FILE_EXT_TSV
    paths(1) = AddSlash_(PATH_ROOT_2) & FILE_EXT_TSV
    paths(2) = AddSlash_(PATH_ROOT_3) & FILE_EXT_TSV

    Dim p As Variant, lines() As String, parts() As String, i As Long
    For Each p In paths
        If FileExists_(CStr(p)) Then
            lines = ReadAllUtf16_(CStr(p))
            If (Not Not lines) <> 0 Then
                For i = LBound(lines) To UBound(lines)
                    If Len(lines(i)) > 0 And Not StartsWith_(lines(i), "Логин") Then
                        parts = Split(lines(i), vbTab)
                        If UBound(parts) >= 0 Then
                            If StrComp(parts(0), login, vbTextCompare) = 0 Then
                                If UBound(parts) >= 3 Then If Len(parts(3)) > 0 Then stChrome = Trim$(parts(3))
                                If UBound(parts) >= 4 Then If Len(parts(4)) > 0 Then stEdge = Trim$(parts(4))
                                If UBound(parts) >= 5 Then If Len(parts(5)) > 0 Then stYandex = Trim$(parts(5))
                            End If
                        End If
                    End If
                Next i
            End If
        End If
    Next p
End Sub

Private Function CheckBrowserExists(browserName As String) As Boolean
    ' Без показа frmLoading
    UpdateLoadingStatus "Проверка наличия " & browserName & "..."
    Dim paths(0 To 2) As String
    Dim i As Integer
    Select Case browserName
        Case "Chrome"
            paths(0) = Environ$("ProgramFiles") & "\Google\Chrome\Application\chrome.exe"
            paths(1) = Environ$("ProgramFiles(x86)") & "\Google\Chrome\Application\chrome.exe"
            paths(2) = Environ$("LOCALAPPDATA") & "\Google\Chrome\Application\chrome.exe"
        Case "Edge"
            paths(0) = Environ$("ProgramFiles") & "\Microsoft\Edge\Application\msedge.exe"
            paths(1) = Environ$("ProgramFiles(x86)") & "\Microsoft\Edge\Application\msedge.exe"
        Case "Yandex"
            paths(0) = Environ$("LOCALAPPDATA") & "\Yandex\YandexBrowser\Application\browser.exe"
            paths(1) = Environ$("ProgramFiles") & "\Yandex\YandexBrowser\Application\browser.exe"
            paths(2) = Environ$("ProgramFiles(x86)") & "\Yandex\YandexBrowser\Application\browser.exe"
    End Select
    For i = 0 To 2
        If paths(i) <> "" Then
            If Dir(paths(i)) <> "" Then CheckBrowserExists = True: Exit Function
        End If
    Next i
    CheckBrowserExists = False
End Function

Private Sub OpenLinkInBrowser(browserName As String)
    On Error Resume Next
    UpdateLoadingStatus "Открытие ссылки подтверждения в " & browserName & "..."
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    Dim sLogin As String: sLogin = Environ$("USERNAME")
    Dim targetUrl As String
    targetUrl = "https://script.google.com/macros/s/AKfycbzkP1L4n2Qc_GR1RigdEnX1kiG4Hw4eE5V3cDNrm3VV4ZYT8db8yTUUKLng1Pvj4Cp7/exec?baseName=CHECK&stageName=" & sLogin & "&duration=1"

    Dim browserCmd As String
    Select Case browserName
        Case "Chrome": browserCmd = "chrome.exe"
        Case "Edge":   browserCmd = "msedge.exe"
        Case "Yandex"
            Dim yandexPath As String: yandexPath = Environ$("LOCALAPPDATA") & "\Yandex\YandexBrowser\Application\browser.exe"
            If Dir(yandexPath) <> "" Then browserCmd = Chr(34) & yandexPath & Chr(34) Else browserCmd = "browser.exe"
    End Select

    wsh.Run browserCmd & " " & targetUrl, 1, False
    Set wsh = Nothing
End Sub

' ==========================================================================================
' === БЛОК 2: ОБНОВЛЕНИЕ ФАЙЛОВ ============================================================
' ==========================================================================================
Private Sub Run_BrowserUpdate_Logic()
    ProcessingResultBrowser = "CheckOnly"
    On Error Resume Next

    ' ВАЖНО: на этом этапе frmLoading не показываем.
    Application.StatusBar = "Поиск доступного сетевого источника..."

    Dim sourcePath As String: sourcePath = GetReadableSourcePath_Extensions()
    If sourcePath = "" Then
        Application.StatusBar = "ВНИМАНИЕ: Сетевые папки недоступны!"
        ProcessingResultBrowser = "NoSource"
        Application.Wait (Now + TimeValue("0:00:02"))
        Exit Sub
    End If

    ProcessingResultBrowser = ProcessExtensions_DeepCopy(sourcePath)
End Sub

Private Function ProcessExtensions_DeepCopy(srcPath As String) As String
    On Error GoTo CLEANUP

    ' Сверка версий делаем без показа формы
    Application.StatusBar = "Сверка версий файлов..."

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim localCachePath As String
    localCachePath = Environ$("APPDATA") & "\Microsoft\Excel\LocalCache"
    If Not fso.FolderExists(localCachePath) Then fso.CreateFolder localCachePath

    Dim needCopyTimer As Boolean: needCopyTimer = False
    Dim localStageTimer As String: localStageTimer = localCachePath & "\StageTimer"

    ' Сверка манифеста
    If Not fso.FolderExists(localStageTimer) Then
        needCopyTimer = True
    Else
        If GetJsonVersion_(localStageTimer & "\" & MANIFEST_FILE) <> GetJsonVersion_(srcPath & "\StageTimer\" & MANIFEST_FILE) Then needCopyTimer = True
    End If

    Dim needInstallBrowser As Boolean: needInstallBrowser = False
    Dim localInstaller As String: localInstaller = localCachePath & "\extension_installer"

    ' Сверка версии установщика (install_*.exe лежат тут)
    If Not fso.FolderExists(localInstaller) Then
        needInstallBrowser = True
    Else
        If ReadCleanVersion_(localInstaller & "\" & INSTALLER_VER_FILE) <> ReadCleanVersion_(srcPath & "\extension_installer\" & INSTALLER_VER_FILE) Then needInstallBrowser = True
    End If

    If (Not needCopyTimer) And (Not needInstallBrowser) Then
        Application.StatusBar = "Файлы актуальны. Обновление не требуется."
        ProcessExtensions_DeepCopy = "NoUpdates"
        GoTo CLEANUP
    End If

    glbCopiedCount = 0

    ' ВАЖНО: frmLoading показываем ТОЛЬКО если обновляется extension_installer (install_*.exe)
    Dim showedLoading As Boolean
    showedLoading = False

    If needInstallBrowser Then
        EnsureLoadingShown "Найдено обновление установщиков (install_*.exe). Копирование..."
        showedLoading = True
    Else
        ' StageTimer обновим без формы
        Application.StatusBar = "Найдено обновление StageTimer. Копирование..."
    End If

    If needCopyTimer Then
        On Error Resume Next
        If showedLoading Then
            UpdateLoadingStatus "Обновление расширения StageTimer..."
        Else
            Application.StatusBar = "Обновление StageTimer..."
        End If

        If fso.FolderExists(localStageTimer) Then fso.DeleteFolder localStageTimer, True
        On Error GoTo CLEANUP
        fso.CreateFolder localStageTimer
        RecursiveCopy fso, srcPath & "\StageTimer", localStageTimer
    End If

    If needInstallBrowser Then
        On Error Resume Next
        UpdateLoadingStatus "Обновление установщиков (extension_installer)..."
        If fso.FolderExists(localInstaller) Then fso.DeleteFolder localInstaller, True
        On Error GoTo CLEANUP
        fso.CreateFolder localInstaller
        RecursiveCopy fso, srcPath & "\extension_installer", localInstaller
    End If

    If showedLoading Then
        UpdateLoadingStatus "Установщики успешно обновлены."
        Application.Wait (Now + TimeValue("0:00:01"))
        EnsureLoadingHidden
    Else
        Application.StatusBar = "Файлы обновлены."
    End If

    ProcessExtensions_DeepCopy = "Updated"

CLEANUP:
    Set fso = Nothing
End Function

Private Sub RecursiveCopy(fso As Object, srcDir As String, destDir As String)
    On Error Resume Next
    Dim fldr As Object: Set fldr = fso.GetFolder(srcDir)
    Dim file As Object, subFldr As Object

    For Each file In fldr.files
        glbCopiedCount = glbCopiedCount + 1

        ' ВАЖНО: никаких показов формы тут нет.
        ' Просто обновляем текст, если форма уже показана.
        If frmLoading.Visible Then
            frmLoading.SetText "Копирование файла: " & file.name
            DoEvents
        End If

        fso.CopyFile file.path, destDir & "\" & file.name, True
    Next file

    For Each subFldr In fldr.SubFolders
        Dim newDest As String: newDest = destDir & "\" & subFldr.name
        If Not fso.FolderExists(newDest) Then fso.CreateFolder newDest
        RecursiveCopy fso, subFldr.path, newDest
    Next subFldr
End Sub

Private Function GetReadableSourcePath_Extensions() As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(PATH_ROOT_1 & SUFFIX_EXT) Then GetReadableSourcePath_Extensions = PATH_ROOT_1 & SUFFIX_EXT: Exit Function
    If fso.FolderExists(PATH_ROOT_2 & SUFFIX_EXT) Then GetReadableSourcePath_Extensions = PATH_ROOT_2 & SUFFIX_EXT: Exit Function
    If fso.FolderExists(PATH_ROOT_3 & SUFFIX_EXT) Then GetReadableSourcePath_Extensions = PATH_ROOT_3 & SUFFIX_EXT: Exit Function
    GetReadableSourcePath_Extensions = ""
End Function

' ==========================================================================================
' === БЛОК 3: ЛОГИРОВАНИЕ (LOGGING) ========================================================
' ==========================================================================================
Public Sub Log_Extension_Action(targetBrowser As String, status As String)
    On Error GoTo EH_EXT
    Dim login As String: login = Environ$("USERNAME")
    If Len(login) = 0 Then Exit Sub
    Dim fio As String: fio = "Unknown"

    ' Тут UpdateLoadingStatus не покажет форму. Это ок.
    UpdateLoadingStatus "Подключение к AD для проверки пользователя..."

    ' Получаем ФИО из AD
    On Error Resume Next
    Dim adConn As Object: Set adConn = CreateObject("ADODB.Connection")
    adConn.Provider = "ADsDSOObject": adConn.Open "Active Directory Provider"
    Dim t$, m$, c$, d$, ci$, adr$
    GetFromAD_ByLogin_ adConn, login, fio, t, m, c, d, ci, adr
    If fio = "" Then fio = login
    If Not adConn Is Nothing Then If adConn.State = 1 Then adConn.Close
    Set adConn = Nothing
    On Error GoTo EH_EXT

    If Len(targetBrowser) > 0 Then
        UpdateLoadingStatus "Запись результата (" & targetBrowser & "=" & status & ") в базу..."
    Else
        UpdateLoadingStatus "Обновление даты активности пользователя..."
    End If

    ' Пишем во все доступные файлы
    Dim filesToProcess As New Collection
    AddIfExists_File filesToProcess, PATH_ROOT_1, FILE_EXT_TSV
    AddIfExists_File filesToProcess, PATH_ROOT_2, FILE_EXT_TSV
    AddIfExists_File filesToProcess, PATH_ROOT_3, FILE_EXT_TSV

    If filesToProcess.count = 0 Then Exit Sub
    Dim fPath As Variant
    For Each fPath In filesToProcess
        UpdateSingleExtensionFile_MultiCol CStr(fPath), login, fio, targetBrowser, status
    Next fPath
    Exit Sub
EH_EXT:
End Sub

Private Sub UpdateSingleExtensionFile_MultiCol(path As String, login As String, fio As String, targetBrowser As String, status As String)
    On Error GoTo EH
    Dim header As String
    header = "Логин" & vbTab & "ФИО" & vbTab & "Дата проверки" & vbTab & "Chrome" & vbTab & "Edge" & vbTab & "Yandex"

    ' === 1. НОВЫЙ ФАЙЛ ===
    If GetFileSize_(path) = 0 Then
        Dim newArr(0 To 5) As String
        newArr(0) = login: newArr(1) = fio: newArr(2) = Format$(Now, "yyyy-MM-dd HH:mm:ss")
        newArr(3) = "Not Installed": newArr(4) = "Not Installed": newArr(5) = "Not Installed"

        If Len(targetBrowser) > 0 Then
            If targetBrowser = "Chrome" Then newArr(3) = status
            If targetBrowser = "Edge" Then newArr(4) = status
            If targetBrowser = "Yandex" Then newArr(5) = status
        End If

        WriteTextDirect_ path, header & vbCrLf & Join(newArr, vbTab)
        Exit Sub
    End If

    ' === 2. ЧТЕНИЕ ===
    Dim lines() As String: lines = ReadAllUtf16_(path)
    Dim dataRows As New Collection
    Dim parts() As String
    Dim i As Long
    Dim foundIndex As Long: foundIndex = -1
    Dim currentRow() As String

    If (Not Not lines) <> 0 Then
        For i = LBound(lines) To UBound(lines)
            Dim ln As String: ln = lines(i)
            If Len(Trim$(ln)) > 0 And Not StartsWith_(ln, "Логин") Then
                parts = Split(ln, vbTab)
                If UBound(parts) < 5 Then ReDim Preserve parts(0 To 5)
                dataRows.Add parts
                If foundIndex = -1 And StrComp(parts(0), login, vbTextCompare) = 0 Then foundIndex = dataRows.count
            End If
        Next i
    End If

    ' === 3. ОБРАБОТКА ===
    If foundIndex > 0 Then
        currentRow = dataRows(foundIndex)
        dataRows.Remove foundIndex
    Else
        ReDim currentRow(0 To 5)
        currentRow(3) = "Not Installed": currentRow(4) = "Not Installed": currentRow(5) = "Not Installed"
    End If

    currentRow(0) = login
    currentRow(1) = fio
    currentRow(2) = Format$(Now, "yyyy-MM-dd HH:mm:ss")

    If Len(targetBrowser) > 0 Then
        If targetBrowser = "Chrome" Then currentRow(3) = status
        If targetBrowser = "Edge" Then currentRow(4) = status
        If targetBrowser = "Yandex" Then currentRow(5) = status
    End If

    dataRows.Add currentRow

    ' === 4. СОРТИРОВКА И ЗАПИСЬ ===
    Dim sortArr() As Variant, count As Long: count = dataRows.count
    If count > 0 Then
        ReDim sortArr(1 To count)
        For i = 1 To count: sortArr(i) = dataRows(i): Next i
        If count > 1 Then
            Dim j As Long, temp As Variant
            For i = 1 To count - 1
                For j = i + 1 To count
                    If IsArray(sortArr(j)) And IsArray(sortArr(i)) Then
                        If UBound(sortArr(j)) >= 2 And UBound(sortArr(i)) >= 2 Then
                            If sortArr(j)(2) > sortArr(i)(2) Then
                                temp = sortArr(i): sortArr(i) = sortArr(j): sortArr(j) = temp
                            End If
                        End If
                    End If
                Next j
            Next i
        End If
    End If

    Dim outLines() As String
    ReDim outLines(0 To count)
    outLines(0) = header
    If count > 0 Then
        For i = 1 To count: outLines(i) = Join(sortArr(i), vbTab): Next i
    End If

    Application.Wait (Now + TimeValue("0:00:01") / 20)
    WriteAllUtf16_ path, outLines
    Exit Sub
EH: Err.Clear
End Sub

' ==========================================================================================
' === ЧАСТЬ 4: СБОР ПРОФИЛЯ ================================================================
' ==========================================================================================
Public Sub CollectAndWrite_Profile_ToTsv()
    On Error GoTo EH_MAIN

    ' Без показа формы
    UpdateLoadingStatus "Запрос профиля в Active Directory..."

    Dim login As String: login = Environ$("USERNAME")
    If LenB(login) = 0 Then Exit Sub
    Dim adConn As Object: Set adConn = CreateObject("ADODB.Connection")
    adConn.Provider = "ADsDSOObject": adConn.Open "Active Directory Provider"
    Dim isDev As Boolean: isDev = False
    Dim myData(0 To 9) As String
    Dim fio$, title$, mail$, company$, dept$, city$, addr$
    GetFromAD_ByLogin_ adConn, login, fio, title, mail, company, dept, city, addr
    myData(0) = login: myData(1) = fio: myData(2) = title: myData(3) = mail
    myData(4) = company: myData(5) = dept: myData(6) = city: myData(7) = addr
    myData(8) = Format$(Now, "yyyy-MM-dd HH:mm:ss"): myData(9) = GetUpdateStatus_(isDev)

    Dim filesToProcess As New Collection
    AddIfExists_File filesToProcess, PATH_ROOT_1, FILE_PROFILE_TSV
    AddIfExists_File filesToProcess, PATH_ROOT_2, FILE_PROFILE_TSV
    AddIfExists_File filesToProcess, PATH_ROOT_3, FILE_PROFILE_TSV

    If filesToProcess.count = 0 Then GoTo CLEANUP
    Dim fPath As Variant
    For Each fPath In filesToProcess
        UpdateLoadingStatus "Обновление файла профиля: " & fPath
        ProcessSingleProfileFile_ CStr(fPath), adConn, isDev, myData, login
    Next fPath
CLEANUP:
    If Not adConn Is Nothing Then If adConn.State = 1 Then adConn.Close
    Set adConn = Nothing
    Exit Sub
EH_MAIN: Resume CLEANUP
End Sub

Private Sub ProcessSingleProfileFile_(ByVal targetPath As String, ByVal adConn As Object, _
                               ByVal isDev As Boolean, ByRef myData() As String, ByVal myLogin As String)
    On Error GoTo EH_FILE
    Dim header As String
    header = "Логин" & vbTab & "ФИО" & vbTab & "Должность" & vbTab & "Почта" & vbTab & _
             "Компания" & vbTab & "Департамент" & vbTab & "Город" & vbTab & "Адрес" & vbTab & _
             "Дата последнего использования" & vbTab & "Статус обновления"
    If GetFileSize_(targetPath) = 0 Then
        WriteTextDirect_ targetPath, header & vbCrLf & Join(myData, vbTab)
        Exit Sub
    End If
    Dim lines() As String: lines = ReadAllUtf16_(targetPath)
    Dim dataRows As New Collection
    Dim parts() As String
    Dim i As Long
    Dim foundMyIndex As Long: foundMyIndex = -1
    If (Not Not lines) <> 0 Then
        For i = LBound(lines) To UBound(lines)
            Dim ln As String: ln = lines(i)
            If Len(Trim$(ln)) > 0 And Not StartsWith_(ln, "Логин") Then
                parts = Split(ln, vbTab)
                If UBound(parts) < 9 Then ReDim Preserve parts(0 To 9)
                dataRows.Add parts
                If foundMyIndex = -1 And StrComp(parts(0), myLogin, vbTextCompare) = 0 Then foundMyIndex = dataRows.count
            End If
        Next i
    End If
    If foundMyIndex > 0 Then
        Dim oldRow() As String: oldRow = dataRows(foundMyIndex)
        Dim sOld As String, sNew As String
        sOld = Norm8_(oldRow(1), oldRow(2), oldRow(3), oldRow(4), oldRow(5), oldRow(6), oldRow(7))
        sNew = Norm8_(myData(1), myData(2), myData(3), myData(4), myData(5), myData(6), myData(7))
        If StrComp(sOld, sNew, vbTextCompare) <> 0 Then
            dataRows.Add myData
        Else
            dataRows.Remove foundMyIndex
            dataRows.Add myData
        End If
    Else
        dataRows.Add myData
    End If
    Dim sortArr() As Variant, count As Long: count = dataRows.count
    If count > 0 Then
        ReDim sortArr(1 To count)
        For i = 1 To count: sortArr(i) = dataRows(i): Next i
        If count > 1 Then
            Dim j As Long, temp As Variant
            For i = 1 To count - 1
                For j = i + 1 To count
                    If IsArray(sortArr(j)) And IsArray(sortArr(i)) Then
                        If UBound(sortArr(j)) >= 8 And UBound(sortArr(i)) >= 8 Then
                            If sortArr(j)(8) > sortArr(i)(8) Then
                                temp = sortArr(i): sortArr(i) = sortArr(j): sortArr(j) = temp
                            End If
                        End If
                    End If
                Next j
            Next i
        End If
    End If
    Dim outLines() As String
    ReDim outLines(0 To count)
    outLines(0) = header
    If count > 0 Then
        For i = 1 To count: outLines(i) = Join(sortArr(i), vbTab): Next i
    End If
    Application.Wait (Now + TimeValue("0:00:01") / 10)
    WriteAllUtf16_ targetPath, outLines
    Exit Sub
EH_FILE: Err.Clear
End Sub

' ==========================================================================================
' === ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ==============================================================
' ==========================================================================================
Private Sub AddIfExists_File(coll As Collection, pathBase As String, fname As String)
    Dim full As String: full = AddSlash_(pathBase) & fname
    If FileExists_(full) Then coll.Add full
End Sub

Private Function GetUpdateStatus_(ByVal isDev As Boolean) As String
    On Error Resume Next
    If isDev Then GetUpdateStatus_ = "+": Exit Function
    Dim localVerPath As String: localVerPath = Environ$("APPDATA") & "\Microsoft\Excel\LocalCache\" & FILE_VER
    Dim localVer As String: localVer = NormalizeVer_(ReadAllTextSimple_(localVerPath))
    If Len(localVer) = 0 Then GetUpdateStatus_ = "Не обновился": Exit Function
    Dim remoteVerPath As String: remoteVerPath = ""
    If FileExists_(AddSlash_(PATH_ROOT_1) & FILE_VER) Then remoteVerPath = AddSlash_(PATH_ROOT_1) & FILE_VER
    If remoteVerPath = "" And FileExists_(AddSlash_(PATH_ROOT_2) & FILE_VER) Then remoteVerPath = AddSlash_(PATH_ROOT_2) & FILE_VER
    If remoteVerPath = "" And FileExists_(AddSlash_(PATH_ROOT_3) & FILE_VER) Then remoteVerPath = AddSlash_(PATH_ROOT_3) & FILE_VER
    If Len(remoteVerPath) = 0 Then GetUpdateStatus_ = "+": Exit Function
    Dim remoteVer As String: remoteVer = NormalizeVer_(ReadAllTextSimple_(remoteVerPath))
    If StrComp(localVer, remoteVer, vbTextCompare) <> 0 Then GetUpdateStatus_ = "Не обновился" Else GetUpdateStatus_ = "+"
End Function

Private Function CheckUserExistsInAD_(ByVal conn As Object, ByVal sam As String, ByVal fio As String) As Boolean
    On Error GoTo EH
    sam = Replace(sam, "(", ""): sam = Replace(sam, ")", ""): sam = Trim$(sam)
    fio = Replace(fio, "(", ""): fio = Replace(fio, ")", ""): fio = Trim$(fio)
    If Len(sam) = 0 And Len(fio) = 0 Then GoTo EH
    Dim root As Object, cmd As Object, rs As Object
    Set root = GetObject("LDAP://rootDSE")
    Set cmd = CreateObject("ADODB.Command"): Set cmd.ActiveConnection = conn
    Dim filter As String: filter = "(|"
    If Len(sam) > 0 Then filter = filter & "(sAMAccountName=" & sam & ")"
    If Len(fio) > 0 Then filter = filter & "(displayName=" & fio & ")"
    filter = filter & ")"
    cmd.CommandText = "<LDAP://" & root.Get("defaultNamingContext") & ">;(&(objectClass=user)" & filter & ");sAMAccountName;subtree"
    Set rs = cmd.Execute
    If rs.EOF Then CheckUserExistsInAD_ = False Else CheckUserExistsInAD_ = True
    rs.Close: Exit Function
EH: CheckUserExistsInAD_ = False
End Function

Private Function GetFromAD_ByLogin_(ByVal conn As Object, ByVal sam As String, _
    ByRef fio As String, ByRef title As String, ByRef mail As String, _
    ByRef company As String, ByRef dept As String, ByRef city As String, ByRef addr As String) As Boolean
    On Error GoTo EH
    Dim root As Object, cmd As Object, rs As Object
    Set root = GetObject("LDAP://rootDSE"): Set cmd = CreateObject("ADODB.Command"): Set cmd.ActiveConnection = conn
    cmd.CommandText = "<LDAP://" & root.Get("defaultNamingContext") & _
                      ">;(&(objectClass=user)(sAMAccountName=" & sam & "));displayName,title,mail,company,department,l,streetAddress;subtree"
    Set rs = cmd.Execute
    If rs.EOF Then GoTo EH
    fio = NzFld_(rs, "displayName"): title = NzFld_(rs, "title"): mail = NzFld_(rs, "mail")
    company = NzFld_(rs, "company"): dept = NzFld_(rs, "department")
    city = NzFld_(rs, "l"): addr = NzFld_(rs, "streetAddress")
    rs.Close: GetFromAD_ByLogin_ = True: Exit Function
EH: GetFromAD_ByLogin_ = False
End Function

Private Function Norm8_(fio$, title$, mail$, company$, dept$, city$, addr$) As String
    Norm8_ = LCase$(Trim$(fio)) & "|" & LCase$(Trim$(title)) & "|" & LCase$(Trim$(mail)) & "|" & _
             LCase$(Trim$(company)) & "|" & LCase$(Trim$(dept)) & "|" & LCase$(Trim$(city)) & "|" & LCase$(Trim$(addr))
End Function

Private Function NzFld_(rs As Object, name$) As String
    On Error Resume Next
    If rs.fields(name).value & "" <> "" Then NzFld_ = CStr(rs.fields(name).value) Else NzFld_ = ""
End Function

Private Function FileExists_(path$) As Boolean
    On Error Resume Next: FileExists_ = (Len(Dir$(path, vbNormal)) > 0)
End Function

Private Function GetFileSize_(path$) As Long
    On Error Resume Next: Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(path) Then GetFileSize_ = fso.GetFile(path).Size Else GetFileSize_ = 0
End Function

Private Function PickWritablePath_(p1$, p2$, p3$, fname$) As String
    Dim t1$, t2$, t3$: t1 = AddSlash_(p1) & fname: t2 = AddSlash_(p2) & fname: t3 = AddSlash_(p3) & fname
    If CanAppend_(t1) Then PickWritablePath_ = t1: Exit Function
    If CanAppend_(t2) Then PickWritablePath_ = t2: Exit Function
    If CanAppend_(t3) Then PickWritablePath_ = t3: Exit Function
End Function

Private Function AddSlash_(p$) As String
    If Right$(p, 1) = "\" Then AddSlash_ = p Else AddSlash_ = p & "\"
End Function

Private Function CanAppend_(path$) As Boolean
    On Error GoTo EH: Dim f%: f = FreeFile: Open path For Append As #f: Close #f: CanAppend_ = True: Exit Function
EH: CanAppend_ = False
End Function

Private Sub WriteTextDirect_(path As String, txt As String)
    On Error Resume Next: Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "unicode": stm.Open: stm.WriteText txt: stm.SaveToFile path, 2: stm.Close
End Sub

Private Function ReadAllUtf16_(ByVal filePath As String) As String()
    On Error GoTo EH
    If GetFileSize_(filePath) = 0 Then ReadAllUtf16_ = Split(vbNullString): Exit Function
    Dim stm As Object, txt As String: Set stm = CreateObject("ADODB.Stream")
    With stm: .Type = 2: .Charset = "unicode": .Open: .LoadFromFile filePath: txt = .ReadText(-1): .Close: End With
    If Len(txt) = 0 Then ReadAllUtf16_ = Split(vbNullString) Else ReadAllUtf16_ = Split(Replace(txt, vbCrLf, vbLf), vbLf)
    Exit Function
EH: ReadAllUtf16_ = Split(vbNullString)
End Function

Private Sub WriteAllUtf16_(path$, lines() As String)
    On Error GoTo EH: Dim stm As Object: Set stm = CreateObject("ADODB.Stream"): Dim txt$, i&
    For i = LBound(lines) To UBound(lines): If i > LBound(lines) Then txt = txt & vbCrLf
    txt = txt & lines(i): Next: stm.Type = 2: stm.Charset = "unicode": stm.Open: stm.WriteText txt: stm.SaveToFile path, 2: stm.Close: Exit Sub
EH:
End Sub

Private Function StartsWith_(s$, p$) As Boolean
    StartsWith_ = (Left$(s, Len(p)) = p)
End Function

Private Function ReadAllTextSimple_(path$) As String
    On Error Resume Next: Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(path) Then ReadAllTextSimple_ = fso.OpenTextFile(path, 1).ReadAll
End Function

Private Function NormalizeVer_(s$) As String
    NormalizeVer_ = Trim$(Replace(Replace(Replace(Replace(s, ChrW(&HFEFF), ""), vbCr, ""), vbLf, ""), vbTab, ""))
End Function

Private Function ReadCleanVersion_(filePath As String) As String
    On Error Resume Next: ReadCleanVersion_ = ""
    Dim fso As Object, txt As String: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(filePath) Then
        txt = fso.OpenTextFile(filePath, 1).ReadAll: txt = Replace(txt, vbCr, ""): txt = Replace(txt, vbLf, ""): txt = Replace(txt, vbTab, "")
        ReadCleanVersion_ = Trim$(txt)
    End If
End Function

Private Function GetJsonVersion_(jsonPath As String) As String
    On Error Resume Next: GetJsonVersion_ = ""
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(jsonPath) Then Exit Function
    Dim txt As String: txt = fso.OpenTextFile(jsonPath, 1).ReadAll
    Dim p1 As Long, p2 As Long
    p1 = InStr(1, txt, """version""", vbTextCompare): If p1 = 0 Then Exit Function
    p1 = InStr(p1, txt, ":"): If p1 = 0 Then Exit Function
    p1 = InStr(p1, txt, """"): If p1 = 0 Then Exit Function
    p2 = InStr(p1 + 1, txt, """"): If p2 = 0 Then Exit Function
    GetJsonVersion_ = Mid$(txt, p1 + 1, p2 - p1 - 1)
End Function

' ==========================================================================================
' === БЛОК РЕСТАРТА БРАУЗЕРОВ (RESTORE SESSION) ============================================
' ==========================================================================================
Private Function ProcessPendingUpdate(browserName As String, currentStatus As String) As String
    ProcessPendingUpdate = currentStatus
    If StrComp(currentStatus, "Updated", vbTextCompare) <> 0 Then Exit Function

    Dim procName As String, exePath As String
    Select Case browserName
        Case "Chrome": procName = "chrome.exe"
        Case "Edge":   procName = "msedge.exe"
        Case "Yandex": procName = "browser.exe"
    End Select

    ' 1) Проверяем, запущен ли браузер
    If Not IsProcessRunning(procName) Then
        Log_Extension_Action browserName, "Installed"
        ProcessPendingUpdate = "Installed"
        UpdateLoadingStatus browserName & ": Браузер закрыт. Статус обновлен."
        Exit Function
    End If

    ' 2) Если запущен — перезапуск
    UpdateLoadingStatus "ПЕРЕЗАПУСК " & browserName & "..."

    exePath = GetBrowserPath(browserName)
    If exePath = "" Then Exit Function

    KillProcess procName
    Application.Wait (Now + TimeValue("0:00:02"))

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    On Error Resume Next
    wsh.Run Chr(34) & exePath & Chr(34) & " --restore-last-session", 1, False
    Set wsh = Nothing
    On Error GoTo 0

    Log_Extension_Action browserName, "Installed"
    ProcessPendingUpdate = "Installed"
    Application.Wait (Now + TimeValue("0:00:02"))
End Function

Private Function IsProcessRunning(processName As String) As Boolean
    Dim objWMIService As Object, colProcess As Object
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name = '" & processName & "'")
    IsProcessRunning = (colProcess.count > 0)
End Function

Private Sub KillProcess(processName As String)
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")

    wsh.Run "taskkill /IM " & processName, 0, True

    Dim i As Integer
    For i = 1 To 3
        If Not IsProcessRunning(processName) Then
            Set wsh = Nothing
            Exit Sub
        End If
        Application.Wait (Now + TimeValue("0:00:01"))
    Next i

    If IsProcessRunning(processName) Then
        wsh.Run "taskkill /F /IM " & processName, 0, True
    End If
    Set wsh = Nothing
End Sub

Private Function GetBrowserPath(browserName As String) As String
    Dim paths(0 To 2) As String
    Dim i As Integer
    Select Case browserName
        Case "Chrome"
            paths(0) = Environ$("ProgramFiles") & "\Google\Chrome\Application\chrome.exe"
            paths(1) = Environ$("ProgramFiles(x86)") & "\Google\Chrome\Application\chrome.exe"
            paths(2) = Environ$("LOCALAPPDATA") & "\Google\Chrome\Application\chrome.exe"
        Case "Edge"
            paths(0) = Environ$("ProgramFiles") & "\Microsoft\Edge\Application\msedge.exe"
            paths(1) = Environ$("ProgramFiles(x86)") & "\Microsoft\Edge\Application\msedge.exe"
        Case "Yandex"
            paths(0) = Environ$("LOCALAPPDATA") & "\Yandex\YandexBrowser\Application\browser.exe"
            paths(1) = Environ$("ProgramFiles") & "\Yandex\YandexBrowser\Application\browser.exe"
            paths(2) = Environ$("ProgramFiles(x86)") & "\Yandex\YandexBrowser\Application\browser.exe"
    End Select

    For i = 0 To 2
        If paths(i) <> "" Then
            If Dir(paths(i)) <> "" Then
                GetBrowserPath = paths(i)
                Exit Function
            End If
        End If
    Next i
    GetBrowserPath = ""
End Function


