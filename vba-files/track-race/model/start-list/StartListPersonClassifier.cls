VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "StartListPersonClassifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/model/start-list
Option Explicit
Implements IObjectClassifier

Private Function IObjectClassifier_Classify(Element As Object) As Variant
    Dim pElement As StartListOrderModel: Set pElement = Element
    IObjectClassifier_Classify = pElement.Person.Id
End Function
