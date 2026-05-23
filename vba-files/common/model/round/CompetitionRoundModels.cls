VERSION 1.0 CLASS
BEGIN
MultiUse = -1 ' True
END
Attribute VB_Name = "CompetitionRoundModels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/common/model/round
Option Explicit
Implements IJsonConvertable
Implements IJsonParsable

Private m_Records As Object
Private m_EventRoundIndexes As Object

Private Sub Class_Initialize()
    Set m_Records = CreateObject("Scripting.Dictionary")
    Set m_EventRoundIndexes = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set m_Records = Nothing
    Set m_EventRoundIndexes = Nothing
End Sub

Public Sub Add(Item As CompetitionRoundModel)
    Dim eventRoundIndex As Variant: eventRoundIndex = EventRoundIdKey(Item)
    ' 種目とラウンドID被りは入力ミスなのでエラーとする.
    Call m_EventRoundIndexes.Add(eventRoundIndex, Item.Id)
    Call m_Records.Add(Item.Id, Item)
End Sub

Public Sub AddRange(Items As Collection)
    Dim recordItem As CompetitionRoundModel
    For Each recordItem In Items
        Call Me.Add(recordItem)
    Next recordItem
End Sub

Public Function All() As Collection
    Dim recordItems As Collection: Set recordItems = New Collection
    Dim recordId As Variant
    For Each recordId In m_Records
        Dim recordItem As CompetitionRoundModel: Set recordItem = m_Records.Item(recordId)
        Call recordItems.Add(recordItem)
    Next recordId

    Set All = recordItems
End Function

Public Function Count() As Long
    Count = m_Records.Count()
End Function

Public Function Item(Id As Variant) As CompetitionRoundModel
    If Not (m_Records.Exists(Id)) Then
        Set Item = Nothing
        Exit Function
    End If

    Set Item = m_Records.Item(Id)
End Function

Public Function FindByEventRound(EventId As Long, RoundId As Long) As CompetitionRoundModel
    If Not (m_EventRoundIndexes.Exists(EventId & "__" & RoundId)) Then
        Set FindByEventRound = Nothing
        Exit Function
    End If

    Dim recordId As Variant: recordId = m_EventRoundIndexes.Item(EventId & "__" & RoundId)
    Set FindByEventRound = Me.Item(recordId)
End Function

Public Function GroupByEvents() As Object
    Dim vGroups As Object: Set vGroups = CollectionUtil.GroupingBy(Me.All(), New EventRoundClassifier)

    Dim vEventId As Variant
    Dim rGroupByEvents As Object: Set rGroupByEvents = CreateObject("Scripting.Dictionary")
    For Each vEventId In vGroups
        Dim vRounds As CompetitionRoundModels: Set vRounds = New CompetitionRoundModels
        Call vRounds.AddRange(vGroups.Item(vEventId))
        Call rGroupByEvents.Add(vEventId, vRounds)
    Next vEventId

    Set GroupByEvents = rGroupByEvents
End Function

Private Function IJsonConvertable_ToJson() As String
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(Me.All)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function

Private Function EventRoundIdKey(Item As CompetitionRoundModel) As Variant
    EventRoundIdKey = Item.EventId & "__" & Item.RoundId
End Function

Public Sub FromJson(JsonString As String)
    Call IJsonParsable_FromJson(JsonString)
End Sub

Private Sub IJsonParsable_FromJson(JsonString As String)
    Dim vElements As Collection: Set vElements = JsonConverter.ParseJson(JsonString)
    Dim vRecords As Collection: Set vRecords = New Collection
    Dim vElement As Object
    For Each vElement In vElements
        Dim vRecord As CompetitionRoundModel: Set vRecord = JsonConverter.ParseJsonTo(JsonConverter.ConvertToJson(vElement), New CompetitionRoundModel)
        Call vRecords.Add(vRecord)
    Next vElement 

    Call Me.AddRange(vRecords)
End Sub
