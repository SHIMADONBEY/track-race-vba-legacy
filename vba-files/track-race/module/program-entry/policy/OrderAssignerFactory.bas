Attribute VB_Name = "OrderAssignerFactory"
'namespace=vba-files/track-race/modules/program-entry/policy
Option Explicit
Option Private Module

Public Function Generate(AssignerType As Long) As IOrderAssignable
    Select Case AssignerType
    Case 0
        If (CompetitionRepository.ReadSettinng().PersonPerGroup > CompetitionRepository.ReadSettinng().LaneSetting.LaneCount) Then
            Set Generate = New RandomAssigner
        Else
            Set Generate = New RandomLaneAssigner
        End If
    Case 1
        Set Generate = New ThreeLaneAssigner
    Case 2
        Set Generate = New AscendOrderAssigner
    Case 3
        Set Generate = New DescendOrderAssigner
    Case Else
        Set Generate = Nothing
    End Select

End Function
