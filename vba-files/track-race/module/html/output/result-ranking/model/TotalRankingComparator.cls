VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "TotalRankingComparator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/modules/html/output/result-ranking/model
Option Explicit
Implements IObjectComparator

Private Function IObjectComparator_Compare(LeftHandObject As Object, RightHandObject As Object) As Integer
    Dim leftHand As ResultRankingModel: Set leftHand = LeftHandObject
    Dim rightHand As ResultRankingModel: Set rightHand = RightHandObject

    IObjectComparator_Compare = IIf(leftHand.NotStarted, 1, 0) - IIf(rightHand.NotStarted, 1, 0)
    If (IObjectComparator_Compare <> 0) Then Exit Function

    IObjectComparator_Compare = IIf(leftHand.NotFinished, 1, 0) - IIf(rightHand.NotFinished, 1, 0)
    If (IObjectComparator_Compare <> 0) Then Exit Function

    IObjectComparator_Compare = IIf(leftHand.Disqualified, 1, 0) - IIf(rightHand.Disqualified, 1, 0)
    If (IObjectComparator_Compare <> 0) Then Exit Function

    IObjectComparator_Compare = IIf(leftHand.DemoEntry, 1, 0) - IIf(rightHand.DemoEntry, 1, 0)
    If (IObjectComparator_Compare <> 0) Then Exit Function

    IObjectComparator_Compare = Sgn(leftHand.TotalRank - rightHand.TotalRank)
    If (IObjectComparator_Compare <> 0) Then Exit Function

    IObjectComparator_Compare = Sgn(leftHand.ResultRank - rightHand.ResultRank)
    If (IObjectComparator_Compare <> 0) Then Exit Function

    IObjectComparator_Compare = Sgn(leftHand.GroupNumber - rightHand.GroupNumber)
    If (IObjectComparator_Compare <> 0) Then Exit Function

    IObjectComparator_Compare = Sgn(leftHand.Rank - rightHand.Rank)
End Function