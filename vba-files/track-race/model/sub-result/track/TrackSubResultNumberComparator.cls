VERSION 1.0 CLASS
  BEGIN
  MultiUse = -1  'True
  END
  Attribute VB_Name = "TrackSubResultNumberComparator"
  Attribute VB_GlobalNameSpace = False
  Attribute VB_Creatable = False
  Attribute VB_PredeclaredId = False
  Attribute VB_Exposed = False
'namespace=vba-files/track-race/model/sub-result/track
Option Explicit
Implements IObjectComparator

Private Function IObjectComparator_Compare(LeftHandObject As Object, RightHandObject As Object) As Integer
    Dim leftHand As TrackSubResultOrderModel: Set leftHand = LeftHandObject
    Dim rightHand As TrackSubResultOrderModel: Set rightHand = RightHandObject

    IObjectComparator_Compare = Sgn(leftHand.RoundId - rightHand.RoundId)
    If IObjectComparator_Compare <> 0 Then Exit Function
    IObjectComparator_Compare = Sgn(leftHand.GroupNumber - rightHand.GroupNumber)
    If IObjectComparator_Compare <> 0 Then Exit Function
    IObjectComparator_Compare = Sgn(leftHand.SubNumber - rightHand.SubNumber)
End Function