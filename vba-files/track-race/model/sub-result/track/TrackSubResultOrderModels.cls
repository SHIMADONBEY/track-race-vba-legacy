VERSION 1.0 CLASS
BEGIN
MultiUse = -1 ' True
END
Attribute VB_Name = "TrackSubResultOrderModels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/model/sub-result/track
Option Explicit
Implements IJsonConvertable

Private m_Records As Object
Private m_ResultSubIndexes As Object

Private Sub Class_Initialize()
    Set m_Records = CreateObject("Scripting.Dictionary")
    Set m_ResultSubIndexes = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set m_Records = Nothing
    Set m_ResultSubIndexes = Nothing
End Sub

Public Sub Add(Item As TrackSubResultOrderModel)
    If Not (Item.ResultId = "") Then
        Call m_ResultSubIndexes.Add(Item.ResultId & "__" & Item.SubNumber, Item.Id)
    End If
    Call m_Records.Add(Item.Id, Item)
End Sub

Public Sub AddRange(Items As Collection)
    Dim recordItem As TrackSubResultOrderModel
    For Each recordItem In Items
        Call Me.Add(recordItem)
    Next recordItem
End Sub

Public Function All() As Collection
    Dim recordItems As Collection: Set recordItems = New Collection
    Dim recordId As Variant

    For Each recordId In m_Records
        Dim recordItem As TrackSubResultOrderModel: Set recordItem = m_Records.Item(recordId)
        Call recordItems.Add(recordItem)
    Next recordId

    Set All = recordItems
End Function

Public Function Count() As Long
    Count = m_Records.Count()
End Function

Public Function Item(Id As Variant) As TrackSubResultOrderModel
    If Not (m_Records.Exists(Id)) Then
        Set Item = Nothing
        Exit Function
    End If

    Set Item = m_Records.Item(Id)
End Function

Public Sub MergeStartList(StartList As StartListOrderModels)
    Dim recordId As Variant
    For Each recordId In m_Records
        Dim recordItem As TrackSubResultOrderModel: Set recordItem = m_Records.Item(recordId)
        Dim vStartRecord As StartListOrderModel: Set vStartRecord = StartList.FindOrder(recordItem.RoundId, recordItem.GroupNumber, recordItem.Order)
        If Not (vStartRecord Is Nothing) Then
            Set recordItem.Person = vStartRecord.Person
        End If
    Next recordId 
End Sub

Public  Function FilterByRound(RoundId As Long) As TrackSubResultOrderModels
    Dim recordList As TrackSubResultOrderModels: Set recordList = New TrackSubResultOrderModels
    Dim recordId As Variant
    For Each recordId In m_Records
        Dim recordItem As TrackSubResultOrderModel: Set recordItem = m_Records.Item(recordId)
        If (recordItem.RoundId = RoundId) Then
            Call recordList.Add(recordItem)
        End If
    Next recordId 

    Set FilterByRound = recordList
End Function

Public Function FilterByResultId(ResultId As String) As TrackSubResultOrderModels
    Dim recordList As TrackSubResultOrderModels: Set recordList = New TrackSubResultOrderModels
    Dim recordId As Variant
    For Each recordId In m_Records
        Dim recordItem As TrackSubResultOrderModel: Set recordItem = m_Records.Item(recordId)
        If (recordItem.ResultId = ResultId) Then
            Call recordList.Add(recordItem)
        End If
    Next recordId 

    Set FilterByResultId = recordList
End Function

Public Function FilterLapTimeByRound(TargetRoundId As Long) As TrackSubResultOrderModels
    Dim recordList As Collection: Set recordList = New Collection
    Dim recordItem As TrackSubResultOrderModel
    For Each recordItem In Me.FilterByRound(TargetRoundId).All()
        If (recordItem.SubNumber = 0) Then
            ' DO NOTHING
        Else
            Call recordList.Add(recordItem)
        End If
    Next recordItem
    
    Set FilterLapTimeByRound = New TrackSubResultOrderModels
    Call FilterLapTimeByRound.AddRange(CollectionUtil.SortObjectCollection(recordList, New TrackSubResultNumberComparator))
End Function
    
Public Function FilterLapTimeByGroup(TargetRoundId As Long, TargetGroup As Long) As TrackSubResultOrderModels
    Dim recordList As Collection: Set recordList = New Collection
    Dim recordItem As TrackSubResultOrderModel
    For Each recordItem In Me.FilterByGroup(TargetRoundId, TargetGroup).All()
        If (recordItem.SubNumber = 0) Then
            ' DO NOTHING
        Else
            Call recordList.Add(recordItem)
        End If
    Next recordItem
    
    Set FilterLapTimeByGroup = New TrackSubResultOrderModels
    Call FilterLapTimeByGroup.AddRange(CollectionUtil.SortObjectCollection(recordList, New TrackSubResultNumberComparator))
End Function

Public Function FilterByGroup(TargetRoundId As Long, TargetGroup As Long) As TrackSubResultOrderModels
    Dim recordList As Collection: Set recordList = New Collection
    Dim recordId As Variant
    For Each recordId In m_Records
        Dim recordItem As TrackSubResultOrderModel: Set recordItem = m_Records.Item(recordId)
        If (recordItem.RoundId <> TargetRoundId) Then 
            ' DO NOTHING
        ElseIf (recordItem.GroupNumber <> TargetGroup) Then
            ' DO NOTHING
        Else
            Call recordList.Add(recordItem)
        End If
    Next recordId
    
    Set FilterByGroup = New TrackSubResultOrderModels
    Call FilterByGroup.AddRange(CollectionUtil.SortObjectCollection(recordList, New TrackSubResultNumberComparator))
End Function

Public Function FindByResultId(ResultId As String, Optional SubNumber As Long = 0) As TrackSubResultOrderModel
    Dim recordIndex As String: recordIndex = ResultId & "__" & SubNumber
    If (Not m_ResultSubIndexes.Exists(recordIndex)) Then
        Set FindByResultId = Nothing
        Exit Function
    End If

    Dim recordId As Variant: recordId = m_ResultSubIndexes.Item(recordIndex)
    Set FindByResultId = Me.Item(recordId)
End Function

Public Function FindByGroupSubNumber(TargetRoundId As Long, TargetGroup As Long, Optional SubNumber As Long = 0) As TrackSubResultOrderModel
    Dim recordItem As TrackSubResultOrderModel
    Set FindByGroupSubNumber = Nothing
    
    For Each recordItem In Me.FilterByGroup(TargetRoundId, TargetGroup).All()
        If (recordItem.SubNumber = SubNumber) Then
            Set FindByGroupSubNumber = recordItem
            Exit Function
        End If
    Next recordItem 
End Function

Private Function IJsonConvertable_ToJson() As String
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(Me.All)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function
