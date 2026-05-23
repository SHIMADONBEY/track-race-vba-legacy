VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "StartListGroupClassifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/model/start-list
Option Explicit
Implements IObjectClassifier

Private Function IObjectClassifier_Classify(Element As Object) As Variant
    Dim recordItem As StartListOrderModel: Set recordItem = Element

    If Not (recordItem.IsOrdered()) Then
        IObjectClassifier_Classify = ""
        Exit Function
    End If

    IObjectClassifier_Classify = recordItem.RoundId & "__" & recordItem.Group
End Function

Public Function RoundId(ElementKey As Variant) As Long
    Dim splitValues() As String: splitValues = Split(CStr(ElementKey), "__")
    RoundId = CInt(splitValues(0))
End Function

Public Function Group(ElementKey As Variant) As Long
    Dim splitValues() As String: splitValues = Split(CStr(ElementKey), "__")
    Group = CInt(splitValues(1))
End Function
