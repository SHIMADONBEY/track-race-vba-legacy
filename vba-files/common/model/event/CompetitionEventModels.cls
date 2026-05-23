VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "CompetitionEventModels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/common/model/event
Option Explicit
Implements IJsonConvertable

Private m_EventNameIndex As Object
Private m_Records As Object

Private Sub Class_Initialize()
    Set m_Records = CreateObject("Scripting.Dictionary")
    Set m_EventNameIndex = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set m_Records = Nothing
    Set m_EventNameIndex = Nothing
End Sub

Public Sub Add(Item As CompetitionEventModel)
    Call m_EventNameIndex.Add(Item.EventKey, Item.Id)
    Call m_Records.Add(Item.Id, Item)
End Sub

Public Sub AddRange(Items As Collection)
    Dim recordItem As CompetitionEventModel
    For Each recordItem In Items
        Call Me.Add(recordItem)
    Next recordItem
End Sub

Public Function All() As Collection
    Dim recordItems As Collection: Set recordItems = New Collection
    Dim recordId As Variant
    For Each recordId In m_Records
        Dim recordItem As CompetitionEventModel: Set recordItem = m_Records.Item(recordId)
        Call recordItems.Add(recordItem)
    Next recordId

    Set All = recordItems
End Function

Public Function Count() As Long
    Count = m_Records.Count()
End Function

Public Function Item(Id As Variant) As CompetitionEventModel
    If Not (m_Records.Exists(Id)) Then
        Set Item = Nothing
        Exit Function
    End If

    Set Item = m_Records.Item(Id)
End Function

Public Function FindByEvent(Optional EventKey As Variant = Empty _
        , Optional Category As String = "" _
        , Optional Sex As String = "" _
        , Optional EventName As String = "" _
        , Optional Supecification As String = "" _
    ) As CompetitionEventModel

    Dim pEventKey As Variant
    If (IsEmpty(EventKey)) Then 
        pEventKey = Category & Sex & EventName & IIf(Supecification = "", "", " (" & Supecification & ")")
    Else
        pEventKey = EventKey
    End If

    If Not (m_EventNameIndex.Exists(pEventKey)) Then
        Set FindByEvent = Nothing
        Exit Function
    End If

    Dim recordId As Variant: recordId = m_EventNameIndex.Item(pEventKey)
    Set FindByEvent = Me.Item(recordId)
End Function

Private Function IJsonConvertable_ToJson() As String
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(Me.All)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function
