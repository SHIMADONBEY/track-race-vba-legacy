VERSION 1.0 CLASS
BEGIN
MultiUse = -1 ' True
END
Attribute VB_Name = "QualifiedAthleteModels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/modules/program-entry/model
Option Explicit
Implements IJsonConvertable


Private m_Records As Object

Private Sub Class_Initialize()
    Set m_Records = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set m_Records = Nothing
End Sub

Public Sub Add(Item As QualifiedAthleteModel)
    Call m_Records.Add(Item.Id, Item)
End Sub

Public Sub AddRange(Items As Collection)
    Dim recordItem As QualifiedAthleteModel
    For Each recordItem In Items
        Call Me.Add(recordItem)
    Next recordItem
End Sub

Public Function All() As Collection
    Dim recordItems As Collection: Set recordItems = New Collection
    Dim recordId As Variant
    For Each recordId In m_Records
        Dim recordItem As QualifiedAthleteModel: Set recordItem = m_Records.Item(recordId)
        Call recordItems.Add(recordItem)
    Next recordId

    Set All = recordItems
End Function

Public Function Count() As Long
    Count = m_Records.Count()
End Function

Public Function Item(Id As Variant) As QualifiedAthleteModel
    If Not (m_Records.Exists(Id)) Then
        Set Item = Nothing
        Exit Function
    End If

    Set Item = m_Records.Item(Id)
End Function

Public Function IsEmpty() As Boolean
    IsEmpty = (Me.Count < 1)
End Function

Public Function Sort(Optional Descending As Boolean = False) As QualifiedAthleteModels
    Dim vSortedList As QualifiedAthleteModels: Set vSortedList = New QualifiedAthleteModels
    Call vSortedList.AddRange(CollectionUtil.SortObjectCollection(Me.All(), New QualifiedRankComparator, Descending))
    Set Sort = vSortedList
End Function

Public Function GroupByGroup() As Object
    Dim vClassifier As QualifiedGroupClassifier: Set vClassifier = New QualifiedGroupClassifier
    Dim vGroupedList As Object: Set vGroupedList = CollectionUtil.GroupingBy(Me.All(), vClassifier)

    Dim vGroupedRecords As Object: Set vGroupedRecords = CreateObject("Scripting.Dictionary")
    Dim vGroupKey As Variant
    For Each vGroupKey In vGroupedList
        Dim vSortedGroupList As QualifiedAthleteModels: Set vSortedGroupList = New QualifiedAthleteModels
        Call vSortedGroupList.AddRange(CollectionUtil.SortObjectCollection(vGroupedList.Item(vGroupKey), New QualifiedOrderComparator))
        Call vGroupedRecords.Add(vGroupKey, vSortedGroupList)
    Next vGroupKey

    Set GroupByGroup = vGroupedRecords
End Function

Private Function IJsonConvertable_ToJson() As String
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(Me.All())
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function
