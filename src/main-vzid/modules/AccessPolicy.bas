Attribute VB_Name = "AccessPolicy"
Option Explicit

Public Function AccessPolicy_CurrentUser() As String
    AccessPolicy_CurrentUser = LCase$(Trim$(Environ$("USERNAME")))
End Function

Public Function AccessPolicy_IsCommandEnabled(ByVal commandId As String) As Boolean
    Dim regionId As String
    regionId = RegionState_SelectedId()

    If Not CommandRegistry_IsSupported(commandId, regionId) Then Exit Function

    If AccessPolicy_HasFullAccess() Then
        AccessPolicy_IsCommandEnabled = True
        Exit Function
    End If

    AccessPolicy_IsCommandEnabled = AccessPolicy_UserInCsv(MainConfig_ReadValue(AccessPolicy_CommandKey(commandId), ""), AccessPolicy_CurrentUser())
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
        Case "RIC_ZAKAZNYE_CREATE"
            AccessPolicy_CommandKey = "commandAccessRicZakaznyeCreateCsv"
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
