VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "TrackGameRecordsJsonParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'namespace=vba-files/track-race/model/game-record
Option Explicit

Implements IJsonParser

Private m_ElementParser As IJsonParser

Private Sub Class_Initialize()
    Set m_ElementParser = New TrackGameRecordJsonParser
End Sub

Private Sub Class_Terminate()
    Set m_ElementParser = Nothing
End Sub

Private Function IJsonParser_Parse(JsonString As String, ByVal NewObject As Object) As Object
    Dim pNewObject As GameRecordModels: Set pNewObject = NewObject
    Dim vElements As Collection : Set vElements = JsonConverter.ParseJson(JsonString)
    Dim vRecords As Collection  : Set vRecords = New Collection

    Dim vElement As Object
    For Each vElement In vElements
        Dim vRecord As GameRecordModel: Set vRecord = JsonConverter.ParseJsonTo(JsonConverter.ConvertToJson(vElement), New GameRecordModel, m_ElementParser)
        Call vRecords.Add(vRecord)
    Next vElement

    Call pNewObject.AddRange(vRecords)
    Set IJsonParser_Parse = pNewObject
End Function
