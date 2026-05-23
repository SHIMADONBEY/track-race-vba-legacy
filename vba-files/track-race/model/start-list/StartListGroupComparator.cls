VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "StartListGroupComparator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/model/start-list
Option Explicit
Implements IObjectComparator

Private Function IObjectComparator_Compare(LeftHandObject As Object, RightHandObject As Object) As Integer
    Dim leftHand As StartListGroupModel: Set leftHand = LeftHandObject
    Dim rightHand As StartListGroupModel: Set rightHand = RightHandObject

    IObjectComparator_Compare = Sgn(leftHand.GroupNumber - rightHand.GroupNumber)
End Function
