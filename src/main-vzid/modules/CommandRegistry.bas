Attribute VB_Name = "CommandRegistry"
Option Explicit

Public Function CommandRegistry_IsSupported(ByVal commandId As String, ByVal regionId As String) As Boolean
    CommandRegistry_IsSupported = (LenB(CommandRegistry_GetMacroName(commandId, regionId)) > 0)
End Function

Public Function CommandRegistry_IsVisible(ByVal commandId As String, ByVal regionId As String) As Boolean
    Select Case UCase$(commandId)
        Case "OSP_SELECT", "HOTKEY_SETUP"
            CommandRegistry_IsVisible = (StrComp(UCase$(regionId), "TMN", vbTextCompare) = 0)
        Case Else
            CommandRegistry_IsVisible = True
    End Select
End Function

Public Function CommandRegistry_GetMacroName(ByVal commandId As String, ByVal regionId As String) As String
    Select Case UCase$(commandId)
        Case "DOC_PACKETS"
            Select Case UCase$(regionId)
                Case "KGN", "TMN", "EKB", "CHLB"
                    CommandRegistry_GetMacroName = "Doc_Packets"
            End Select

        Case "MAKE_COVER_LETTERS"
            Select Case UCase$(regionId)
                Case "KGN"
                    CommandRegistry_GetMacroName = "Make_Cover_Letters"
                Case "TMN"
                    CommandRegistry_GetMacroName = "TMN_Make_Cover_Letters_likeKGN"
                Case "CHLB"
                    CommandRegistry_GetMacroName = "CHLB_Make_Cover_Letters_ByRecipient"
            End Select

        Case "MAKE_COVER_LETTERS_BY_RECIPIENT"
            Select Case UCase$(regionId)
                Case "KGN"
                    CommandRegistry_GetMacroName = "Make_Cover_Letters_ByRecipient"
                Case "TMN"
                    CommandRegistry_GetMacroName = "TMN_Make_Cover_Letters_ByRecipient_likeKGN"
                Case "CHLB"
                    CommandRegistry_GetMacroName = "CHLB_Make_Cover_Letters_ByRecipient"
            End Select

        Case "ZAKAZNYE_CREATE"
            Select Case UCase$(regionId)
                Case "KGN"
                    CommandRegistry_GetMacroName = "KGN_Zakaznye_Create"
                Case "TMN"
                    CommandRegistry_GetMacroName = "TMN_Zakaznye_Create"
                Case "EKB"
                    CommandRegistry_GetMacroName = "RIC_Zakaznye_Create"
            End Select

        Case "MANUAL_PROC_REPORT"
            Select Case UCase$(regionId)
                Case "KGN"
                    CommandRegistry_GetMacroName = "CommandHandlers_ShowManualProcReportKGN"
                Case "TMN"
                    CommandRegistry_GetMacroName = "CommandHandlers_ShowManualProcReportTMN"
                Case "EKB"
                    CommandRegistry_GetMacroName = "CommandHandlers_ShowManualProcReportEKB"
            End Select

        Case "PDF_SIGNER"
            Select Case UCase$(regionId)
                Case "KGN", "TMN", "EKB", "CHLB"
                    CommandRegistry_GetMacroName = "LaunchPDFSigner"
            End Select

        Case "PDF_SCANNER"
            Select Case UCase$(regionId)
                Case "KGN", "TMN", "EKB", "CHLB"
                    CommandRegistry_GetMacroName = "RunPDFScanner"
            End Select

        Case "OSP_SELECT"
            Select Case UCase$(regionId)
                Case "TMN"
                    CommandRegistry_GetMacroName = "OspSelect_Run_TMN"
            End Select

        Case "HOTKEY_SETUP"
            Select Case UCase$(regionId)
                Case "TMN"
                    CommandRegistry_GetMacroName = "CommandHandlers_ShowHotkeySetupTMN"
            End Select
    End Select
End Function
