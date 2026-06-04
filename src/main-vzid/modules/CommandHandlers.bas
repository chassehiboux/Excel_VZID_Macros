Attribute VB_Name = "CommandHandlers"
Option Explicit

Public Sub CommandHandlers_ShowManualProcReportKGN()
    frmVZID_KGN_ManualProcReport.Show
End Sub

Public Sub CommandHandlers_ShowManualProcReportTMN()
    frmVZID_TMN_ManualProcReport.Show
End Sub

Public Sub CommandHandlers_ShowManualProcReportEKB()
    frmVZID_EKB_ManualProcReport.Show
End Sub

Public Sub CommandHandlers_ShowHotkeySetupTMN()
    modHotkeyManager.Hotkey_ShowSetup "OspSelect_Run"
End Sub
