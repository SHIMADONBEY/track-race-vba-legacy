Attribute VB_Name = "LaneSettingFactory"
'namespace=vba-files/track-race/module/event/model
Option Explicit
Option Private Module

Public Function GenerateByLanesCount(LanesCount As Long) As LaneSettingModel
    Dim vData As LaneSettingModel: Set vData = New LaneSettingModel
    Select Case LanesCount 
        Case 9
            Call vData.Initialize(LanesCount, "6,5,7,4", "8,9", "3,2", "1")
        Case 8
            Call vData.Initialize(LanesCount, "5,4,6,3", "7,8", "2,1", "")
        Case 6
            Call vData.Initialize(LanesCount, "4,3,5,2", "6", "1", "")
        Case 4
            Call vData.Initialize(LanesCount, "3,2", "4", "1", "")
        Case Else
            Set vData = Nothing
    End Select

    Set GenerateByLanesCount = vData
End Function
