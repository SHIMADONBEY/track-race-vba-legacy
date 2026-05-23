VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "TrackSubResultOrderClassifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/model/sub-result/track
Option Explicit
Implements IObjectClassifier

Private Function IObjectClassifier_Classify(Element As Object) As Variant
    Dim recordItem As TrackSubResultOrderModel: Set recordItem = Element
    IObjectClassifier_Classify = recordItem.ResultId
End Function
