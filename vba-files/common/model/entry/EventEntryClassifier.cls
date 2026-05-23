VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "EventEntryClassifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/common/model/entry
Option Explicit
Implements IObjectClassifier

Private Function IObjectClassifier_Classify(Element As Object) As Variant
    Dim recordItem As PersonalEntryModel: Set recordItem = Element
    IObjectClassifier_Classify = recordItem.EventId
End Function
