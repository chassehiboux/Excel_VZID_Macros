Attribute VB_Name = "AccessPolicy"
Option Explicit

Public Function AccessPolicy_CurrentUser() As String
    AccessPolicy_CurrentUser = LCase$(Trim$(Environ$("USERNAME")))
End Function

Public Function AccessPolicy_IsCommandEnabled(ByVal commandId As String) As Boolean
    Dim regionId As String
    Dim commandKey As String

    regionId = RegionState_SelectedId()

    If Not CommandRegistry_IsVisible(commandId, regionId) Then Exit Function
    If Not CommandRegistry_IsSupported(commandId, regionId) Then Exit Function

    If AccessPolicy_HasFullAccess() Then
        AccessPolicy_IsCommandEnabled = True
        Exit Function
    End If

    commandKey = AccessPolicy_CommandKey(commandId)
    If LenB(commandKey) = 0 Then Exit Function

    AccessPolicy_IsCommandEnabled = AccessPolicy_UserInCsv( _
        MainConfig_ReadValue(commandKey, AccessPolicy_CommandDefaultCsv(commandId)), _
        AccessPolicy_CurrentUser())
End Function

Private Function AccessPolicy_HasFullAccess() As Boolean
    AccessPolicy_HasFullAccess = AccessPolicy_UserInCsv(MainConfig_ReadValue("fullAccessUsersCsv", ""), AccessPolicy_CurrentUser())
End Function

Private Function AccessPolicy_CommandKey(ByVal commandId As String) As String
    Select Case UCase$(commandId)
        Case "DOC_PACKETS"
            AccessPolicy_CommandKey = "commandAccessDocPacketsCsv"
        Case "MAKE_COVER_LETTERS"
            AccessPolicy_CommandKey = "commandAccessMakeCoverLettersCsv"
        Case "MAKE_COVER_LETTERS_BY_RECIPIENT"
            AccessPolicy_CommandKey = "commandAccessMakeCoverLettersByRecipientCsv"
        Case "ZAKAZNYE_CREATE"
            AccessPolicy_CommandKey = "commandAccessZakaznyeCreateCsv"
        Case "MANUAL_PROC_REPORT"
            AccessPolicy_CommandKey = "commandAccessManualProcReportCsv"
        Case "PDF_SIGNER"
            AccessPolicy_CommandKey = "commandAccessPdfSignerCsv"
        Case "PDF_SCANNER"
            AccessPolicy_CommandKey = "commandAccessPdfScannerCsv"
        Case "OSP_SELECT"
            AccessPolicy_CommandKey = "commandAccessOspSelectCsv"
        Case "HOTKEY_SETUP"
            AccessPolicy_CommandKey = "commandAccessHotkeySetupCsv"
    End Select
End Function

Private Function AccessPolicy_CommandDefaultCsv(ByVal commandId As String) As String
    Select Case UCase$(commandId)
        Case "DOC_PACKETS", _
             "MAKE_COVER_LETTERS", _
             "MAKE_COVER_LETTERS_BY_RECIPIENT", _
             "ZAKAZNYE_CREATE", _
             "MANUAL_PROC_REPORT", _
             "PDF_SIGNER", _
             "PDF_SCANNER", _
             "OSP_SELECT", _
             "HOTKEY_SETUP"
            AccessPolicy_CommandDefaultCsv = "*"
    End Select
End Function

Private Function AccessPolicy_UserInCsv(ByVal csvValue As String, ByVal userName As String) As Boolean
    Dim normalized As String

    normalized = LCase$(Trim$(csvValue))
    If LenB(normalized) = 0 Then Exit Function
    If normalized = "*" Then
        AccessPolicy_UserInCsv = True
        Exit Function
    End If

    normalized = Replace(normalized, ",", ";")
    normalized = Replace(normalized, " ", "")

    AccessPolicy_UserInCsv = (InStr(1, ";" & normalized & ";", ";" & LCase$(userName) & ";", vbTextCompare) > 0)
End Function
