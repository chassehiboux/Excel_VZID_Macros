Attribute VB_Name = "RegionState"
Option Explicit

Private mSelectedRegion As String

Public Sub RegionState_Initialize()
    Dim regionId As String

    regionId = UCase$(Trim$(MainConfig_ReadValue("selectedRegion", "KGN")))
    If Not RegionState_IsKnown(regionId) Then regionId = "KGN"

    mSelectedRegion = regionId
End Sub

Public Function RegionState_Count() As Integer
    RegionState_Count = 4
End Function

Public Function RegionState_Label(ByVal index As Integer) As String
    Select Case index
        Case 0: RegionState_Label = "Курган"
        Case 1: RegionState_Label = "Тюмень"
        Case 2: RegionState_Label = "Екатеринбург"
        Case 3: RegionState_Label = "ЧЭС"
    End Select
End Function

Public Function RegionState_Id(ByVal index As Integer) As String
    Select Case index
        Case 0: RegionState_Id = "KGN"
        Case 1: RegionState_Id = "TMN"
        Case 2: RegionState_Id = "EKB"
        Case 3: RegionState_Id = "CHLB"
    End Select
End Function

Public Function RegionState_SelectedId() As String
    If LenB(mSelectedRegion) = 0 Then RegionState_Initialize
    RegionState_SelectedId = mSelectedRegion
End Function

Public Function RegionState_SelectedIndex() As Integer
    Dim index As Integer

    If LenB(mSelectedRegion) = 0 Then RegionState_Initialize

    For index = 0 To RegionState_Count() - 1
        If StrComp(RegionState_Id(index), mSelectedRegion, vbTextCompare) = 0 Then
            RegionState_SelectedIndex = index
            Exit Function
        End If
    Next index
End Function

Public Sub RegionState_SetSelectedId(ByVal regionId As String)
    regionId = UCase$(Trim$(regionId))
    If Not RegionState_IsKnown(regionId) Then Exit Sub

    mSelectedRegion = regionId
    MainConfig_WriteValue "selectedRegion", regionId
End Sub

Public Function RegionState_IsKnown(ByVal regionId As String) As Boolean
    Select Case UCase$(regionId)
        Case "KGN", "TMN", "EKB", "CHLB"
            RegionState_IsKnown = True
    End Select
End Function
