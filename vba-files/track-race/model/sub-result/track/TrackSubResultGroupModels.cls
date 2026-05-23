VERSION 1.0 CLASS
BEGIN
MultiUse = -1 ' True
END
Attribute VB_Name = "TrackSubResultGroupModels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/model/sub-result/track
Option Explicit
Implements IJsonConvertable

Private m_Records As Object
Private m_GroupIndexes As Object

Private Sub Class_Initialize()
    Set m_Records = CreateObject("Scripting.Dictionary")
    Set m_GroupIndexes = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set m_Records = Nothing
    Set m_GroupIndexes = Nothing
End Sub

Public Sub Add(Item As TrackSubResultGroupModel)
    Dim recordIndex As Variant: recordIndex = Item.RoundId & "__" & Item.Group
    Call m_GroupIndexes.Add(recordIndex, Item.Id)
    Call m_Records.Add(Item.Id, Item)
End Sub

Public Sub AddRange(Items As Collection)
    Dim recordItem As TrackSubResultGroupModel
    For Each recordItem In Items
        Call Me.Add(recordItem)
    Next recordItem
End Sub

Public Function All() As Collection
    Dim recordItems As Collection: Set recordItems = New Collection
    Dim recordId As Variant

    For Each recordId In m_Records
        Dim recordItem As TrackSubResultGroupModel: Set recordItem = m_Records.Item(recordId)
        Call recordItems.Add(recordItem)
    Next recordId

    Set All = recordItems
End Function

Public Function Count() As Long
    Count = m_Records.Count()
End Function

Public Function Item(Id As Variant) As TrackSubResultGroupModel
    If Not (m_Records.Exists(Id)) Then
        Set Item = Nothing
        Exit Function
    End If

    Set Item = m_Records.Item(Id)
End Function

Public Function FindByGroup(RoundId As Long, Group As Long) As TrackSubResultGroupModel
    Dim recordIndex As Variant: recordIndex = RoundId & "__" & Group
    If Not (m_GroupIndexes.Exists(recordIndex)) Then
        Set FindByGroup = Nothing
        Exit Function
    End If

    Dim recordId As Variant: recordId = m_GroupIndexes.Item(recordIndex)
    Set FindByGroup = m_Records.Item(recordId)
End Function

Public Function FilterByRound(RoundId As Long) As TrackSubResultGroupModels
    Dim filtered As TrackSubResultGroupModels: Set filtered = New TrackSubResultGroupModels

    Dim recordId As Variant
    For Each recordId In m_Records
        Dim recordItem As TrackSubResultGroupModel: Set recordItem = m_Records.Item(recordId)
        If (recordItem.RoundId = RoundId) Then
            Call filtered.Add(recordItem)
        End If
    Next recordId 

    Set FilterByRound = filtered
End Function

Private Function IJsonConvertable_ToJson() As String
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(Me.All())
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function
