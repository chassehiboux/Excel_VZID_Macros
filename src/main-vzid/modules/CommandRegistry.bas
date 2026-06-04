Attribute VB_Name = "CommandRegistry"
Option Explicit

Public Function CommandRegistry_IsSupported(ByVal commandId As String, ByVal regionId As String) As Boolean
    CommandRegistry_IsSupported = (LenB(CommandRegistry_GetMacroName(commandId, regionId)) > 0)
End Function

Public Function CommandRegistry_GetMacroName(ByVal commandId As String, ByVal regionId As String) As String
    Select Case UCase$(commandId)
        Case "DOC_PACKETS"
            Select Case UCase$(regionId)
                Case "KGN", "TMN", "EKB"
                    CommandRegistry_GetMacroName = "Doc_Packets"
            End Select

        Case "MAKE_COVER_LETTERS"
            Select Case UCase$(regionId)
                Case "KGN"
                    CommandRegistry_GetMacroName = "Make_Cover_Letters"
                Case "TMN"
                    CommandRegistry_GetMacroName = "TMN_Make_Cover_Letters_likeKGN"
            End Select

        Case "MAKE_COVER_LETTERS_BY_RECIPIENT"
            Select Case UCase$(regionId)
                Case "KGN"
                    CommandRegistry_GetMacroName = "Make_Cover_Letters_ByRecipient"
                Case "TMN"
                    CommandRegistry_GetMacroName = "TMN_Make_Cover_Letters_ByRecipient_likeKGN"
            End Select

        Case "RIC_ZAKAZNYE_CREATE"
            Select Case UCase$(regionId)
                Case "EKB"
                    CommandRegistry_GetMacroName = "RIC_Zakaznye_Create"
            End Select
    End Select
End Function
