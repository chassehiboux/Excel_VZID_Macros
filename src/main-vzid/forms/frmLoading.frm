VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLoading 
   Caption         =   "Обновление/Установка"
   ClientHeight    =   2475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12480
   OleObjectBlob   =   "frmLoading.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================================
' === API ДЛЯ РЕЖИМА "ПОВЕРХ ВСЕХ ОКОН" ====================================================
' ==========================================================================================
#If VBA7 Then
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#Else
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

' Флаг отмены
Public IsCancelRequested As Boolean

' ==========================================================================================
' === ЛОГИКА ФОРМЫ =========================================================================
' ==========================================================================================

Private Sub UserForm_Initialize()
    IsCancelRequested = False
    MakeAlwaysOnTop
End Sub

Private Sub UserForm_Activate()
    MoveToBottomCenter
End Sub

' --- ОБРАБОТЧИК КНОПКИ ОТМЕНА ---
Private Sub btnCancel_Click()
    IsCancelRequested = True
    Me.lblStatus.Caption = "Прерывание процесса..." & vbCrLf & "Пожалуйста, подождите..."
    Me.btnCancel.Enabled = False ' Чтобы не кликали много раз
    DoEvents
End Sub

Private Sub MoveToBottomCenter()
    On Error Resume Next
    Dim appLeft As Double, appTop As Double, appWidth As Double, appHeight As Double
    
    appLeft = Application.Left: appTop = Application.Top
    appWidth = Application.width: appHeight = Application.Height
    
    If Application.WindowState = xlMaximized Then
        appLeft = 0: appTop = 0
    End If
    
    Me.Left = appLeft + (appWidth - Me.width) / 2
    Me.Top = appTop + appHeight - Me.Height - 45
    
    If Me.Top < 0 Then Me.Top = 0
    If Me.Left < 0 Then Me.Left = 0
    DoEvents
End Sub

Public Sub ToggleTopMost(enable As Boolean)
    Dim hwnd As Variant, flag As Long
    hwnd = FindWindow("ThunderDFrame", Me.Caption)
    If hwnd <> 0 Then
        If enable Then flag = HWND_TOPMOST Else flag = HWND_NOTOPMOST
        SetWindowPos hwnd, flag, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
    DoEvents
End Sub

Private Sub MakeAlwaysOnTop()
    ToggleTopMost True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        ' Если нажали крестик - тоже считаем за отмену
        IsCancelRequested = True
        Cancel = True
    End If
End Sub

Public Sub SetText(txt As String)
    ' Если нажата отмена, не обновляем текст статусом из модуля, чтобы не сбивать пользователя
    If Not IsCancelRequested Then
        Me.lblStatus.Caption = txt
    End If
    Me.Repaint
    DoEvents
End Sub
