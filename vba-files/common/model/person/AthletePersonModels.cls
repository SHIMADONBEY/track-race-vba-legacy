VERSION 1.0 CLASS
BEGIN
MultiUse = -1 ' True
END
Attribute VB_Name = "AthletePersonModels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/common/model/person
Option Explicit
Implements IJsonConvertable
Implements IJsonParsable

Private m_Records As Object

Private Sub Class_Initialize()
    Set m_Records = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set m_Records = Nothing
End Sub

Public Sub Add(Item As AthletePersonModel)
    Call m_Records.Add(Item.Id, Item)
End Sub

Public Sub AddRange(Items As Collection)
    Dim recordItem As AthletePersonModel
    For Each recordItem In Items
        Call Me.Add(recordItem)
    Next recordItem
End Sub

Public Function All() As Collection
    Dim recordItems As Collection: Set recordItems = New Collection
    Dim recordId As Variant
    For Each recordId In m_Records
        Dim recordItem As AthletePersonModel: Set recordItem = m_Records.Item(recordId)
        Call recordItems.Add(recordItem)
    Next recordId

    Set All = recordItems
End Function

Public Function Count() As Long
    Count = m_Records.Count()
End Function

Public Function Item(Id As Variant) As AthletePersonModel
    If Not (m_Records.Exists(Id)) Then
        Set Item = Nothing
        Exit Function
    End If

    Set Item = m_Records.Item(Id)
End Function

Private Function IJsonConvertable_ToJson() As String
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(Me.All)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function

Public Sub FromJson(JsonString As String)
    Call IJsonParsable_FromJson(JsonString)
End Sub

Private Sub IJsonParsable_FromJson(JsonString As String)
    Dim vElements As Collection: Set vElements = JsonConverter.ParseJson(JsonString)
    Dim vRecords As Collection: Set vRecords = New Collection
    Dim vElement As Object
    For Each vElement In vElements
        Dim vRecord As AthletePersonModel: Set vRecord = JsonConverter.ParseJsonTo(JsonConverter.ConvertToJson(vElement), New AthletePersonModel)
        Call vRecords.Add(vRecord)
    Next vElement 

    Call Me.AddRange(vRecords)
End Sub
