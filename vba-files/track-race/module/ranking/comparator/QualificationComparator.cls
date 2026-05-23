VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "QualificationComparator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/modules/ranking/comparator
Option Explicit
Implements IObjectComparator

Private Function IObjectComparator_Compare(LeftHandObject As Object, RightHandObject As Object) As Integer
    Dim leftHand As TrackResultRankingModel: Set leftHand = LeftHandObject
    Dim rightHand As TrackResultRankingModel: Set rightHand = RightHandObject

    If (leftHand.QualifiedByRank Or rightHand.QualifiedByRank) Then
        IObjectComparator_Compare = Sgn(leftHand.Rank - rightHand.Rank)
        If (IObjectComparator_Compare <> 0) Then
            Exit Function
        End If
    End If

    IObjectComparator_Compare = leftHand.RealResult.Compare(rightHand.RealResult)
End Function
