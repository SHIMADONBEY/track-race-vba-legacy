VERSION 1.0 CLASS
BEGIN
MultiUse = -1 ' True
END
Attribute VB_Name = "TrackResultOrderModels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/model/result/track
Option Explicit
Implements IJsonConvertable

Private m_Records As Object
Private m_OrderKeys As Object

Private Sub Class_Initialize()
    Set m_Records = CreateObject("Scripting.Dictionary")
    Set m_OrderKeys = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set m_Records = Nothing
    Set m_OrderKeys = Nothing
End Sub

Public Sub Add(Item As TrackResultOrderModel)
    Dim orderKey As String: orderKey = Item.RoundId & "__" & Item.Group & "__" & Item.Order
    Call m_OrderKeys.Add(orderKey, Item.Id)
    Call m_Records.Add(Item.Id, Item)
End Sub

Public Sub AddRange(Items As Collection)
    Dim recordItem As TrackResultOrderModel
    For Each recordItem In Items
        Call Me.Add(recordItem)
    Next recordItem
End Sub

Public Function All() As Collection
    Dim recordItems As Collection: Set recordItems = New Collection
    Dim recordId As Variant
    For Each recordId In m_Records
        Dim recordItem As TrackResultOrderModel: Set recordItem = m_Records.Item(recordId)
        Call recordItems.Add(recordItem)
    Next recordId

    Set All = recordItems
End Function

Public Function Count() As Long
    Count = m_Records.Count()
End Function

Public Function Item(Id As Variant) As TrackResultOrderModel
    If Not (m_Records.Exists(Id)) Then
        Set Item = Nothing
        Exit Function
    End If

    Set Item = m_Records.Item(Id)
End Function


Public Function SortByOrder() As TrackResultOrderModels
    Set SortByOrder = SortMe(New TrackResultOrderComparator)
End Function

Public Function SortByRank() As TrackResultOrderModels
    Set SortByRank = SortMe(New TrackResultRankComparator)
End Function

Public Sub MergeStartList(StartList As StartListOrderModels)
    Dim resultId As Variant
    For Each resultId In m_Records
        Dim resultRecord As TrackResultOrderModel: Set resultRecord = m_Records.Item(resultId)
        Dim startRecord As StartListOrderModel: Set startRecord = StartList.FindOrder(resultRecord.RoundId, resultRecord.Group, resultRecord.Order)
        If Not (startRecord Is Nothing) Then
            Set resultRecord.Accumaltive = startRecord.Accumaltive
            Set resultRecord.Person = startRecord.Person
        End If
    Next resultId 
End Sub

Public Function FilterByRound(RoundId As Long) As TrackResultOrderModels
    Dim filtered As Collection: Set filtered = New Collection
    Dim recordId As Variant
    For Each recordId In m_Records
        Dim record As TrackResultOrderModel: Set record = m_Records.Item(recordId)
        If (record.RoundId = RoundId) Then
            Call filtered.Add(record)
        End If
    Next recordId 

    Set FilterByRound = New TrackResultOrderModels
    Call FilterByRound.AddRange(filtered)
End Function

Public Function FilterByRoundGroup(RoundId As Long, Group As Long) As TrackResultOrderModels
    Dim filtered As Collection: Set filtered = New Collection
    Dim recordId As Variant
    For Each recordId In m_Records
        Dim record As TrackResultOrderModel: Set record = m_Records.Item(recordId)
        If (record.RoundId = RoundId And record.Group = Group) Then
            Call filtered.Add(record)
        End If
    Next recordId 

    Set FilterByRoundGroup = New TrackResultOrderModels
    Call FilterByRoundGroup.AddRange(filtered)
End Function

Public Function FilterQualified(RoundId As Long) As TrackResultOrderModels
    Dim filtered As Collection: Set filtered = New Collection
    Dim recordId As Variant
    For Each recordId In m_Records
        Dim record As TrackResultOrderModel: Set record = m_Records.Item(recordId)
        If ((record.RoundId = RoundId) And (record.Qualified())) Then
            Call filtered.Add(record)
        End If
    Next recordId 
    Set FilterQualified = New TrackResultOrderModels
    Call FilterQualified.AddRange(filtered)
End Function

Public Function FilterNeedToDraw(RoundId As Long) As TrackResultOrderModels
    Dim filtered As Collection: Set filtered = New Collection
    Dim recordId As Variant
    For Each recordId In m_Records
        Dim record As TrackResultOrderModel: Set record = m_Records.Item(recordId)
        If ((record.RoundId = RoundId) And (record.NeedToDraw())) Then
            Call filtered.Add(record)
        End If
    Next recordId 
    Set FilterNeedToDraw = New TrackResultOrderModels
    Call FilterNeedToDraw.AddRange(filtered)
End Function
    
Public Function FindByOrder(RoundId As Long, Group As Long, Order As Long) As TrackResultOrderModel
    Dim orderKey As String: orderKey = RoundId & "__" & Group & "__" & Order
    If Not(m_OrderKeys.Exists(orderKey)) Then
        Set FindByOrder = Nothing
        Exit Function
    End If

    Dim recordId As Variant: recordId = m_OrderKeys.Item(orderKey)
    Set FindByOrder = Me.Item(recordId)
End Function

Public Function Self() 
    Set Self = Me
End Function

Private Function IJsonConvertable_ToJson() As String
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(Me.All)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function

Private Function SortMe(ObjectComparator As IObjectComparator) As TrackResultOrderModels
    Dim sortedRecords As Collection: Set sortedRecords = CollectionUtil.SortObjectCollection(Me.All(), ObjectComparator)
    Dim sorted As TrackResultOrderModels: Set sorted = New TrackResultOrderModels
    Call sorted.AddRange(sortedRecords)
    Set SortMe = sorted
End Function
