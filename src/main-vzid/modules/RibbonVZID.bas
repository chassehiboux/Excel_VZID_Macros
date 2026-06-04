Attribute VB_Name = "RibbonVZID"
Option Explicit

Private mRibbon As IRibbonUI

Public Sub RibbonVZID_OnLoad(ByVal ribbon As IRibbonUI)
    Set mRibbon = ribbon
    MainEntryPoint_Initialize
End Sub

Public Sub RibbonVZID_GetRegionCount(ByVal control As IRibbonControl, ByRef returnedVal)
    returnedVal = RegionState_Count()
End Sub

Public Sub RibbonVZID_GetRegionId(ByVal control As IRibbonControl, ByVal index As Integer, ByRef returnedVal)
    returnedVal = RegionState_Id(index)
End Sub

Public Sub RibbonVZID_GetRegionLabel(ByVal control As IRibbonControl, ByVal index As Integer, ByRef returnedVal)
    returnedVal = RegionState_Label(index)
End Sub

Public Sub RibbonVZID_GetSelectedRegionIndex(ByVal control As IRibbonControl, ByRef returnedVal)
    returnedVal = RegionState_SelectedIndex()
End Sub

Public Sub RibbonVZID_OnRegionChanged(ByVal control As IRibbonControl, ByVal selectedId As String, ByVal selectedIndex As Integer)
    RegionState_SetSelectedId selectedId
    RibbonVZID_Invalidate
End Sub

Public Sub RibbonVZID_RunCommand(ByVal control As IRibbonControl)
    CommandDispatcher_Run control.Tag
End Sub

Public Sub RibbonVZID_GetCommandEnabled(ByVal control As IRibbonControl, ByRef returnedVal)
    returnedVal = AccessPolicy_IsCommandEnabled(control.Tag)
End Sub

Public Sub RibbonVZID_GetUpdateStatusLabel(ByVal control As IRibbonControl, ByRef returnedVal)
    returnedVal = UpdateStatus_RibbonText()
End Sub

Public Sub RibbonVZID_CheckUpdates(ByVal control As IRibbonControl)
    MainEntryPoint_RunCheckUpdates
End Sub

Public Sub RibbonVZID_DownloadUpdate(ByVal control As IRibbonControl)
    MainEntryPoint_RunDownloadUpdate
End Sub

Public Sub RibbonVZID_GetDownloadUpdateEnabled(ByVal control As IRibbonControl, ByRef returnedVal)
    returnedVal = UpdateStatus_CanDownloadUpdate()
End Sub

Public Sub RibbonVZID_InstallFromFile(ByVal control As IRibbonControl)
    MainEntryPoint_InstallUpdateFromFile
End Sub

Public Sub RibbonVZID_Invalidate()
    If Not mRibbon Is Nothing Then mRibbon.Invalidate
End Sub
