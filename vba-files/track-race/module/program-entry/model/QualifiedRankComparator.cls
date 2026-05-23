VERSION 1.0 CLASS
  BEGIN
  MultiUse = -1  'True
  END
  Attribute VB_Name = "QualifiedRankComparator"
  Attribute VB_GlobalNameSpace = False
  Attribute VB_Creatable = False
  Attribute VB_PredeclaredId = False
  Attribute VB_Exposed = False
'namespace=vba-files/track-race/modules/program-entry/model
Option Explicit
Implements IObjectComparator

Private Function IObjectComparator_Compare(LeftHandObject As Object, RightHandObject As Object) As Integer
    Dim leftHand As QualifiedAthleteModel: Set leftHand = LeftHandObject
    Dim rightHand As QualifiedAthleteModel: Set rightHand = RightHandObject

    IObjectComparator_Compare = IIf(leftHand.DemoEntry, 1, 0) - IIf(rightHand.DemoEntry, 1, 0)
    If IObjectComparator_Compare <> 0 Then Exit Function

    Dim thisRank As Long: thisRank = IIf(leftHand.QualifiedRank < 1, 2147483647, leftHand.QualifiedRank)
    Dim thatRank As Long: thatRank = IIf(rightHand.QualifiedRank < 1, 2147483647, rightHand.QualifiedRank)
    IObjectComparator_Compare = Sgn(thisRank - thatRank)
    If IObjectComparator_Compare <> 0 Then Exit Function

    IObjectComparator_Compare = leftHand.Qualified1.Compare(rightHand.Qualified1)
End Function
