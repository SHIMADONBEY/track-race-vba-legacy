VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "TrackResultRankComparator"
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

    Dim leftRankGrade As Long: leftRankGrade = IIf(leftHand.DemoEntry, 1, 0) + IIf(leftHand.NotFinished, 2, 0) + IIf(leftHand.Disqualified, 4, 0) + IIf(leftHand.NotStarted, 8, 0)
    Dim rightRankGrade As Long: rightRankGrade = IIf(rightHand.DemoEntry, 1, 0) + IIf(rightHand.NotFinished, 2, 0) + IIf(rightHand.Disqualified, 4, 0) + IIf(rightHand.NotStarted, 8, 0)

    IObjectComparator_Compare = Sgn(leftRankGrade - rightRankGrade)
    If (IObjectComparator_Compare <> 0) Then
        Exit Function
    End If

    IObjectComparator_Compare = Sgn(leftHand.Rank - rightHand.Rank)
    If (IObjectComparator_Compare <> 0) Then
        Exit Function
    End If

    IObjectComparator_Compare = Sgn(leftHand.Order - rightHand.Order)
End Function
