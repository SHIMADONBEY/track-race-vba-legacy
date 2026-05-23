VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "TrackResultOrderComparator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/model/result/track
Option Explicit
Implements IObjectComparator

Private  Function IObjectComparator_Compare(LeftHandObject As Object, RightHandObject As Object) As Integer
    Dim leftHand As TrackResultOrderModel: Set leftHand = leftHandObject
    Dim rightHand As TrackResultOrderModel: Set rightHand = RightHandObject

    IObjectComparator_Compare = Sgn(leftHand.RoundId - rightHand.RoundId)
    If (IObjectComparator_Compare <> 0) Then
        Exit Function
    End If

    IObjectComparator_Compare = Sgn(leftHand.Group - rightHand.Group)
    If (IObjectComparator_Compare <> 0) Then
        Exit Function
    End If

    IObjectComparator_Compare = Sgn(leftHand.Order - rightHand.Order)
End Function
