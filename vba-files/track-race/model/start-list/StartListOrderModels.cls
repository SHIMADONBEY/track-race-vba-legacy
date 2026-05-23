VERSION 1.0 CLASS
BEGIN
MultiUse = -1 ' True
END
Attribute VB_Name = "StartListOrderModels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/model/start-list
Option Explicit
Implements IJsonConvertable

Private m_Records As Object
Private m_OrderDictionary As Object

Private Sub Class_Initialize()
    Set m_Records = CreateObject("Scripting.Dictionary")
    Set m_OrderDictionary = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set m_Records = Nothing
    Set m_OrderDictionary = Nothing
End Sub

Public Sub Add(Item As StartListOrderModel)
    If (Item.IsOrdered()) Then
        Dim orderKey As String: orderKey = Item.RoundId & "__" & Item.Group & "__" & Item.Order 
        Call m_OrderDictionary.Add(orderKey, Item.Id)
    End If

    Call m_Records.Add(Item.Id, Item)
End Sub

Public Sub AddRange(Items As Collection)
    Dim recordItem As StartListOrderModel
    For Each recordItem In Items
        Call Me.Add(recordItem)
    Next recordItem
End Sub

Public Function All() As Collection
    Dim recordItems As Collection: Set recordItems = New Collection
    Dim recordId As Variant

    For Each recordId In m_Records
        Dim recordItem As StartListOrderModel: Set recordItem = m_Records.Item(recordId)
        Call recordItems.Add(recordItem)
    Next recordId

    Set All = recordItems
End Function

Public Function Count() As Long
    Count = m_Records.Count()
End Function

Public Function Item(Id As Variant) As StartListOrderModel
    If Not (m_Records.Exists(Id)) Then
        Set Item = Nothing
        Exit Function
    End If

    Set Item = m_Records.Item(Id)
End Function

Public Function FilterByRound(TargetRoundId As Long) As StartListOrderModels
    Dim filtered As Collection: Set filtered = New Collection
    Dim recordId As Variant
    For Each recordId In m_Records
        Dim recordItem As StartListOrderModel: Set recordItem = m_Records.Item(recordId)
        If (recordItem.RoundId = TargetRoundId) Then
            Call filtered.Add(recordItem)
        End If
    Next recordId 

    Set FilterByRound = New StartListOrderModels
    Call FilterByRound.AddRange(filtered)
End Function

Public Function FilterByGroup(TargetGroup As Long, Optional TargetRoundId As Long = -2147483648) As StartListOrderModels
    Dim filtered As Collection: Set filtered = New Collection
    Dim recordId As Variant
    For Each recordId In m_Records
        Dim recordItem As StartListOrderModel: Set recordItem = m_Records.Item(recordId)
        If (TargetRoundId <= 0 Or recordItem.RoundId = TargetRoundId) Then
            If (recordItem.Group = TargetGroup) Then
                Call filtered.Add(recordItem)
            End If
        End If
    Next recordId 

    Set FilterByGroup = New StartListOrderModels
    Call FilterByGroup.AddRange(filtered)
End Function

Public Function FindOrder(RoundId As Long, Group As Long, Order As Long) As StartListOrderModel
    Dim orderKey As Variant: orderKey = RoundId & "__" & Group & "__" & Order
    If Not (m_OrderDictionary.Exists(orderKey)) Then
        Set FindOrder = Nothing
        Exit Function
    End If

    Dim recordId As Variant: recordId = m_OrderDictionary.Item(orderKey)
    Set FindOrder = Me.Item(recordId)
End Function

Public Sub UpdateOrderKeys()
    Dim recordId As Variant
    Set m_OrderDictionary = CreateObject("Scripting.Dictionary")
    For Each recordId In m_Records
        Dim recordItem As StartListOrderModel: Set recordItem = m_Records.Item(recordId)
        If (recordItem.IsOrdered()) Then
            Dim orderKey As String: orderKey = recordItem.RoundId & "__" & recordItem.Group & "__" & recordItem.Order
            Call m_OrderDictionary.Add(orderKey, recordId)
        End If
    Next recordId 
End Sub

Public Function SortByRankable() As StartListOrderModels
    Set SortByRankable = Sort(New StartListOrderEntryComparator, False)
End Function

Public Function SortByOrder() As StartListOrderModels
    Set SortByOrder = Sort(New StartListGroupOrderComparator, False)
End Function

Private Function IJsonConvertable_ToJson() As String
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(Me.All)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function

Private Function Sort(Comparator As IObjectComparator, Descending As Boolean) As StartListOrderModels
    Dim vSorted As Collection: Set vSorted = CollectionUtil.SortObjectCollection(Me.All(), Comparator, Descending)
    Dim vItems As StartListOrderModels: Set vItems = New StartListOrderModels
    Call vItems.AddRange(vSorted)
    Set Sort = vItems
End Function
