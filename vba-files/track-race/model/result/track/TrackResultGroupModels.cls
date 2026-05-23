VERSION 1.0 CLASS
BEGIN
MultiUse = -1 ' True
END
Attribute VB_Name = "TrackResultGroupModels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/model/result/track
Option Explicit
Implements IJsonConvertable

Private m_Records As Object
Private m_RoundGroupIndex As Object

Private Sub Class_Initialize()
    Set m_Records = CreateObject("Scripting.Dictionary")
    Set m_RoundGroupIndex = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set m_Records = Nothing
    Set m_RoundGroupIndex = Nothing
End Sub

Public Sub Add(Item As TrackResultGroupModel)
    Dim recordIndex As Variant: recordIndex = Item.RoundId & "__" & Item.GroupNumber
    Call m_RoundGroupIndex.Add(recordIndex, Item.Id)
    Call m_Records.Add(Item.Id, Item)
End Sub

Public Sub AddRange(Items As Collection)
    Dim recordItem As TrackResultGroupModel
    For Each recordItem In Items
        Call Me.Add(recordItem)
    Next recordItem
End Sub

Public Function All() As Collection
    Dim recordItems As Collection: Set recordItems = New Collection
    Dim recordId As Variant
    For Each recordId In m_Records
        Dim recordItem As TrackResultGroupModel: Set recordItem = m_Records.Item(recordId)
        Call recordItems.Add(recordItem)
    Next recordId

    Set All = recordItems
End Function

Public Function Count() As Long
    Count = m_Records.Count()
End Function

Public Function Item(Id As Variant) As TrackResultGroupModel
    If Not (m_Records.Exists(Id)) Then
        Set Item = Nothing
        Exit Function
    End If

    Set Item = m_Records.Item(Id)
End Function

Public Function FindByRoundGroup(RoundId As Long, GroupNumber As Long) As TrackResultGroupModel
    Dim recordIndex As Variant: recordIndex = RoundId & "__" & GroupNumber
    If Not (m_RoundGroupIndex.Exists(recordIndex)) Then
        Set FindByRoundGroup = Nothing
        Exit Function
    End If

    Dim recordId As Variant: recordId = m_RoundGroupIndex.Item(recordIndex)
    Set FindByRoundGroup = Me.Item(recordId)
End Function

Public Sub AddResults(RoundId As Long, GroupNumber As Long, Results As Collection)
    Dim recordItem As TrackResultGroupModel: Set recordItem = Me.FindByRoundGroup(RoundId, GroupNumber)

    If (recordItem Is Nothing) Then
        Exit Sub
    End If

    Call recordItem.AddResults(Results)
End Sub

Public Sub AddSplitTimes(RoundId As Long, GroupNumber As Long, SplitTimes As Collection)
    Dim recordItem As TrackResultGroupModel: Set recordItem = Me.FindByRoundGroup(RoundId, GroupNumber)

    If (recordItem Is Nothing) Then
        Exit Sub
    End If

    Call recordItem.AddSplitTimes(SplitTimes)
End Sub

Private Function IJsonConvertable_ToJson() As String
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(Me.All)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function
