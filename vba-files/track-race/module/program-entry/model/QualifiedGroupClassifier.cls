VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "QualifiedGroupClassifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/module/program-entry/model
Option Explicit
Implements IObjectClassifier

Private Const PRIFIX As String = "Group__"

Private Function IObjectClassifier_Classify(Element As Object) As Variant
    Dim recordItem As QualifiedAthleteModel: Set recordItem = Element
    IObjectClassifier_Classify = PRIFIX & recordItem.Group
End Function

Public Function KeyFromGroup(Key As Variant) As Long
    KeyFromGroup = CLng(Replace(Key, PREFIX, ""))
End Function
