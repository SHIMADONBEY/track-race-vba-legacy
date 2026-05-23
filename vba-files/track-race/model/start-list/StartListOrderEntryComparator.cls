VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "StartListOrderEntryComparator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/model/start-list
Option Explicit
Implements IObjectComparator

Private Function IObjectComparator_Compare(LeftHandObject As Object, RightHandObject As Object) As Integer
    Dim pLeftHand As StartListOrderModel    : Set pLeftHand = LeftHandObject
    Dim pRightHand As StartListOrderModel   : Set pRightHand = RightHandObject

    IObjectComparator_Compare = Sgn(CalculateStartingStatus(pLeftHand) - CalculateStartingStatus(pRightHand))
End Function

Private Function CalculateStartingStatus(TargetObject As StartListOrderModel) As Long
    CalculateStartingStatus = 2147483647
    If (TargetObject Is Nothing) Then
        Exit Function
    End If

    CalculateStartingStatus = IIf(TargetObject.DemoEntry, 2 ^ 0, 0)
    CalculateStartingStatus = IIf(TargetObject.NotStarted, 2 ^ 1, 0) + CalculateStartingStatus
End Function
