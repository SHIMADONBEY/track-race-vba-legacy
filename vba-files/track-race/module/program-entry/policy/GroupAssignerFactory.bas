Attribute VB_Name = "GroupAssignerFactory"
'namespace=vba-files/track-race/modules/program-entry/policy
Option Explicit
Option Private Module

Public Function Generate(AssignerType As Long) As IGroupAssignable
    Select Case AssignerType
        Case 1
            Set Generate = New EquallyGroupAssigner
        Case 2
            Set Generate = New AscendGroupAssigner
        Case 3
            Set Generate = New DescendGroupAssigner
        Case Else
            Set Generate = Nothing
    End Select
End Function
