VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "StartListGroupOrderComparator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/model/start-list
Option Explicit
Implements IObjectComparator

Private Function IObjectComparator_Compare(LeftHandObject As Object, RightHandObject As Object) As Integer
    Dim leftHand As StartListOrderModel: Set leftHand = LeftHandObject
    Dim rightHand As StartListOrderModel: Set rightHand = RightHandObject

    IObjectComparator_Compare = Sgn(leftHand.Group - rightHand.Group)
    If IObjectComparator_Compare <> 0 Then Exit Function
    IObjectComparator_Compare = Sgn(leftHand.Order - rightHand.Order)
End Function