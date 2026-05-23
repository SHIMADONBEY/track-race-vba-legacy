VERSION 1.0 CLASS
BEGIN
MultiUse = -1 ' True
END
Attribute VB_Name = "GameRecordModels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/common/model/game-record
Option Explicit
Implements IJsonConvertable

Private m_Records As Object

Private Sub Class_Initialize()
    Set m_Records = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set m_Records = Nothing
End Sub

Public Sub Add(Item As GameRecordModel)
    Call m_Records.Add(Item.Id, Item)
End Sub

Public Sub AddRange(Items As Collection)
    Dim recordItem As GameRecordModel
    For Each recordItem In Items
        Call Me.Add(recordItem)
    Next recordItem
End Sub

Public Function All() As Collection
    Dim recordItems As Collection: Set recordItems = New Collection
    Dim recordId As Variant
    For Each recordId In m_Records
        Dim recordItem As GameRecordModel: Set recordItem = m_Records.Item(recordId)
        Call recordItems.Add(recordItem)
    Next recordId

    Set All = recordItems
End Function

Public Function Count() As Long
    Count = m_Records.Count()
End Function

Public Function Item(Id As Variant) As GameRecordModel
    If Not (m_Records.Exists(Id)) Then
        Set Item = Nothing
        Exit Function
    End If

    Set Item = m_Records.Item(Id)
End Function

Public Function GroupBy(ByVal Classifier As IObjectClassifier) As Object
    Dim vGrouped As Object: Set vGrouped = CreateObject("Scripting.Dictionary")

    Dim vGroupedCollection As Object: Set vGroupedCollection = CollectionUtil.GroupingBy(Me.All(), Classifier)
    Dim vGroupKey As Variant
    For Each vGroupKey In vGroupedCollection
        Dim vRecords As GameRecordModels: Set vRecords = New GameRecordModels
        Call vRecords.AddRange(vGroupedCollection.Item(vGroupKey))
        Call vGrouped.Add(vGroupKey, vRecords)
    Next vGroupKey

    Set GroupBy = vGrouped
End Function

Private Function IJsonConvertable_ToJson() As String
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(Me.All)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function
