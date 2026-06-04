Attribute VB_Name = "PDFSigner"
Option Explicit

' ==========================================================================================
' === КОНСТАНТЫ ПУТЕЙ И ФАЙЛОВ =============================================================
' ==========================================================================================
Private Const APP_EXE_NAME As String = "PDFSigner.exe"
Private Const APP_VER_FILE As String = "version.txt"
Private Const APP_TEMPLATE_FILE As String = "Шаблон_подписи.pdf"

' Пути (относительно корня диска и AppData)
Private Const LOCAL_APP_PATH_SUFFIX As String = "\PDFSigner"
Private Const REMOTE_APP_PATH_SUFFIX As String = "\ВЗИД\Extensions\PDFSigner"

' Сетевые пути (корни)
Private Const PATH_ROOT_1 As String = "\\corp.vostok-electra.ru\Kgn\Отделы\Отдел взыскания по исполнительным документам\Зуйкевич Данил Иванович\Excel"
Private Const PATH_ROOT_2 As String = "\\corp.vostok-electra.ru\Tmn\Общая\ОВЗИД\Зуйкевич"
Private Const PATH_ROOT_3 As String = "\\Ekb-vpfs01\екатеринбург\Отделы\Отдел взыскания задолженности по исполнительным документам\26. Макрос"

' ==========================================================================================
' === ОСНОВНОЙ МАКРОС ЗАПУСКА ==============================================================
' ==========================================================================================
Sub LaunchPDFSigner()
    On Error GoTo GlobalError
    
    ' 1. Инициализация формы загрузки
    If Not frmLoading.Visible Then frmLoading.Show vbModeless
    UpdateLoadingStatus "Подготовка к запуску PDFSigner..."
    
    ' 2. Проверка и установка обновлений (вызов перенесенной логики)
    CheckAndUpdate_PDFSigner
    
    ' 3. Определение локальных путей для запуска
    Dim appData As String
    Dim folderPath As String
    Dim exePath As String
    
    appData = Environ$("APPDATA")
    folderPath = appData & "\Microsoft\Excel\LocalCache" & LOCAL_APP_PATH_SUFFIX
    exePath = folderPath & "\" & APP_EXE_NAME
    
    ' 4. Финальная проверка перед запуском
    If Dir(exePath) = "" Then
        Unload frmLoading
        MsgBox "Не удалось найти файл приложения после обновления!" & vbCrLf & vbCrLf & _
               "Путь: " & exePath, vbCritical, "Ошибка"
        Exit Sub
    End If
    
    ' 5. Запуск приложения
    UpdateLoadingStatus "Запуск приложения..."
    
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    
    ' Важно: задаем рабочую директорию, чтобы PDFSigner видел шаблон рядом с собой
    wsh.CurrentDirectory = folderPath
    
    ' Запуск (1 = vbNormalFocus, False = не ждать завершения)
    wsh.Run Chr(34) & exePath & Chr(34), 1, False
    
    ' 6. Завершение
    Set wsh = Nothing
    Application.StatusBar = "PDFSigner запущен..."
    Application.OnTime Now + TimeValue("00:00:03"), "ClearStatusBar"
    
    ' Небольшая пауза, чтобы пользователь увидел "Готово"
    UpdateLoadingStatus "Готово!"
    Application.Wait (Now + TimeValue("0:00:01") / 2)
    
    Unload frmLoading
    Exit Sub

GlobalError:
    Unload frmLoading
    MsgBox "Критическая ошибка запуска: " & Err.Description, vbCritical
End Sub

' ==========================================================================================
' === ЛОГИКА ОБНОВЛЕНИЯ (Перенесено из InstallBrowser) =====================================
' ==========================================================================================
Private Sub CheckAndUpdate_PDFSigner()
    On Error GoTo EH
    
    ' 1. Поиск источника
    UpdateLoadingStatus "Поиск сервера обновлений..."
    Dim sourcePath As String
    sourcePath = GetReadableSourcePath_PDFSigner()
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim localAppPath As String
    localAppPath = Environ$("APPDATA") & "\Microsoft\Excel\LocalCache" & LOCAL_APP_PATH_SUFFIX
    
    ' Если папки нет, создаем
    If Not fso.FolderExists(localAppPath) Then fso.CreateFolder localAppPath
    
    ' Если сети нет, но локальный файл есть — просто выходим (работаем оффлайн)
    If sourcePath = "" Then
        If fso.FileExists(localAppPath & "\" & APP_EXE_NAME) Then
            UpdateLoadingStatus "Работа в автономном режиме..."
            Application.Wait (Now + TimeValue("0:00:01") / 2)
            Exit Sub
        Else
            ' Сети нет и файла нет — беда
            UpdateLoadingStatus "Ошибка: Нет доступа к сети и файл не установлен."
            Application.Wait (Now + TimeValue("0:00:02"))
            Exit Sub
        End If
    End If
    
    ' 2. Определение путей к файлам
    Dim remoteVersionPath As String: remoteVersionPath = sourcePath & "\" & APP_VER_FILE
    Dim localVersionPath As String: localVersionPath = localAppPath & "\" & APP_VER_FILE
    
    Dim remoteExePath As String: remoteExePath = sourcePath & "\" & APP_EXE_NAME
    Dim localExePath As String: localExePath = localAppPath & "\" & APP_EXE_NAME
    
    Dim remoteTplPath As String: remoteTplPath = sourcePath & "\" & APP_TEMPLATE_FILE
    Dim localTplPath As String: localTplPath = localAppPath & "\" & APP_TEMPLATE_FILE
    
    ' 3. Чтение версий
    Dim remoteVersion As String: remoteVersion = ReadCleanVersion_(remoteVersionPath)
    Dim localVersion As String: localVersion = "0.0.0"
    
    If fso.FileExists(localVersionPath) Then
        localVersion = ReadCleanVersion_(localVersionPath)
    End If
    
    ' 4. Проверка условий обновления
    Dim needUpdate As Boolean: needUpdate = False
    
    ' А. Версии отличаются
    If remoteVersion <> "" And remoteVersion <> localVersion Then needUpdate = True
    ' Б. Нет самого EXE файла локально
    If Not fso.FileExists(localExePath) Then needUpdate = True
    
    ' 5. Выполнение обновления
    If needUpdate Then
        UpdateLoadingStatus "Доступна версия " & remoteVersion & ". Обновление..."
        
        If fso.FileExists(remoteExePath) Then
            ' Копируем EXE
            fso.CopyFile remoteExePath, localExePath, True
            ' Копируем версию
            fso.CopyFile remoteVersionPath, localVersionPath, True
            ' Копируем шаблон (если он есть на сервере)
            If fso.FileExists(remoteTplPath) Then
                fso.CopyFile remoteTplPath, localTplPath, True
            End If
            
            UpdateLoadingStatus "Приложение обновлено до версии " & remoteVersion
            Application.Wait (Now + TimeValue("0:00:01"))
        Else
            UpdateLoadingStatus "Ошибка обновления: Файл не найден на сервере."
            Application.Wait (Now + TimeValue("0:00:01"))
        End If
    Else
        ' Версия актуальна, но проверяем, есть ли ШАБЛОН локально
        ' Если EXE есть, а шаблона нет — докачиваем его тихо
        If Not fso.FileExists(localTplPath) And fso.FileExists(remoteTplPath) Then
            UpdateLoadingStatus "Загрузка файла шаблона..."
            fso.CopyFile remoteTplPath, localTplPath, True
        Else
            ' Все отлично
            UpdateLoadingStatus "Версия актуальна (" & localVersion & ")."
            ' Короткая пауза, чтобы глаз успел заметить
            ' Application.Wait (Now + TimeValue("0:00:01") / 5)
        End If
    End If
    
    Set fso = Nothing
    Exit Sub
    
EH:
    UpdateLoadingStatus "Ошибка при проверке обновлений."
    Set fso = Nothing
    ' Не выходим с фатальной ошибкой, пытаемся запустить то что есть
End Sub

' ==========================================================================================
' === ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ==============================================================
' ==========================================================================================

Private Function GetReadableSourcePath_PDFSigner() As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(PATH_ROOT_1 & REMOTE_APP_PATH_SUFFIX) Then GetReadableSourcePath_PDFSigner = PATH_ROOT_1 & REMOTE_APP_PATH_SUFFIX: Exit Function
    If fso.FolderExists(PATH_ROOT_2 & REMOTE_APP_PATH_SUFFIX) Then GetReadableSourcePath_PDFSigner = PATH_ROOT_2 & REMOTE_APP_PATH_SUFFIX: Exit Function
    If fso.FolderExists(PATH_ROOT_3 & REMOTE_APP_PATH_SUFFIX) Then GetReadableSourcePath_PDFSigner = PATH_ROOT_3 & REMOTE_APP_PATH_SUFFIX: Exit Function
    GetReadableSourcePath_PDFSigner = ""
End Function

Private Function ReadCleanVersion_(filePath As String) As String
    On Error Resume Next
    ReadCleanVersion_ = ""
    Dim fso As Object, txt As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(filePath) Then
        txt = fso.OpenTextFile(filePath, 1).ReadAll
        txt = Replace(txt, vbCr, "")
        txt = Replace(txt, vbLf, "")
        txt = Replace(txt, vbTab, "")
        ReadCleanVersion_ = Trim$(txt)
    End If
End Function

Private Sub UpdateLoadingStatus(txt As String)
    If Not frmLoading.Visible Then frmLoading.Show vbModeless
    frmLoading.SetText txt
    DoEvents
End Sub

Private Sub ClearStatusBar()
    Application.StatusBar = False
End Sub

