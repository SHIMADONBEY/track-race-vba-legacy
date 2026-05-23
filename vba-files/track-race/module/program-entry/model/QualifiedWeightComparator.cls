VERSION 1.0 CLASS
  BEGIN
  MultiUse = -1  'True
  END
  Attribute VB_Name = "QualifiedWeightComparator"
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

    IObjectComparator_Compare = Sgn(leftHand.Weight - rightHand.Weight)
End Function
