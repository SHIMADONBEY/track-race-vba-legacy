VERSION 1.0 CLASS
BEGIN
MultiUse = -1 ' True
END
Attribute VB_Name = "PersonalEntryModels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/common/model/entry
Option Explicit
Implements IJsonConvertable

Private m_Records As Object

Private Sub Class_Initialize()
    Set m_Records = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set m_Records = Nothing
End Sub

Public Sub Add(Item As PersonalEntryModel)
    Dim recordId As Variant: recordId = Item.PersonalId & "__" & Item.EntryId
    Call m_Records.Add(recordId, Item)
End Sub

Public Sub AddRange(Items As Collection)
    Dim recordItem As PersonalEntryModel
    For Each recordItem In Items
        Call Me.Add(recordItem)
    Next recordItem
End Sub

Public Function All() As Collection
    Dim recordItems As Collection: Set recordItems = New Collection
    Dim recordId As Variant
    For Each recordId In m_Records
        Dim recordItem As PersonalEntryModel: Set recordItem = m_Records.Item(recordId)
        Call recordItems.Add(recordItem)
    Next recordId

    Set All = recordItems
End Function

Public Function Count() As Long
    Count = m_Records.Count()
End Function

Public Function Item(PersonalId As String, EntryId As Long) As PersonalEntryModel
    Dim recordId As Variant: recordId = PersonalId & "__" & EntryId
    If Not (m_Records.Exists(recordId)) Then
        Set Item = Nothing
        Exit Function
    End If

    Set Item = m_Records.Item(recordId)
End Function

Public Function FilterByEvent(EventId As Long) As PersonalEntryModels
    Dim filteredList As PersonalEntryModels: Set filteredList = New PersonalEntryModels
    Dim recordId As Variant
    For Each recordId In m_Records
        Dim recordItem As PersonalEntryModel: Set recordItem = m_Records.Item(recordId)
        If (recordItem.EventId = EventId) Then
            Call filteredList.Add(recordItem)
        End If
    Next recordId

    Set FilterByEvent = filteredList
End Function

Public Function GroupByPerson() As Object
    Set GroupByPerson = GroupBy(New PersonEntryClassifier)
End Function

Public Function GroupByEvent() As Object
    Set GroupByEvent = GroupBy(New EventEntryClassifier)
End Function

Private Function IJsonConvertable_ToJson() As String
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(Me.All)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function

Private Function GroupBy(Classifier As IObjectClassifier)
    Dim vGroupedCollection As Object: Set vGroupedCollection = CollectionUtil.GroupingBy(Me.All(), Classifier)

    Dim vGroupedRecords As Object: Set vGroupedRecords = CreateObject("Scripting.Dictionary")
    Dim keyValue As Variant
    For Each keyValue In vGroupedCollection
        Dim vGroup As PersonalEntryModels: Set vGroup = New PersonalEntryModels
        Call vGroup.AddRange(vGroupedCollection.Item(keyValue))
        Call vGroupedRecords.Add(keyValue, vGroup)
    Next keyValue

    Set GroupBy = vGroupedRecords
End Function
