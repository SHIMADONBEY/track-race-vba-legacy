VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "StartListRoundModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/model/start-list
Option Explicit
Implements IJsonConvertable

Private m_Groups As StartListGroupModels
Private m_RoundId As Long

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
    Set m_Groups = New StartListGroupModels 
End Sub

Private Sub Class_Terminate()
    Set m_Groups = Nothing
End Sub

Public Sub Initialize(GroupCount As Long, RoundId As Long, Optional StartOrders As StartListOrderModels = Nothing)
    If m_Initialized Then Err.Raise(91)

    Dim groupKey As Variant
    Dim keyGenerator As StartListGroupClassifier: Set keyGenerator = New StartListGroupClassifier
    Dim registeredCount As Long: registeredCount = 0

    m_RoundId = RoundId

    If Not (StartOrders Is Nothing) Then
        Dim groupedOrders As Object: Set groupedOrders = CollectionUtil.GroupingBy(StartOrders.All(), keyGenerator)

        For Each groupKey In groupedOrders
            If (keyGenerator.RoundId(groupKey) = RoundId) Then
                With New StartListGroupModel
                    Call .Initialize(RoundId, keyGenerator.Group(groupKey))
                    Call .AddRange(groupedOrders.Item(groupKey))
                    Call m_Groups.Add(.Self())
                End With

                registeredCount = registeredCount + 1
            End If
        Next groupKey 
    End If

    Dim i As Long
    For i = (registeredCount + 1) To GroupCount
        With New StartListGroupModel
            Call .Initialize(RoundId, i)
            Call m_Groups.Add(.Self())
        End With
    Next i

    m_Initialized = True
End Sub

Public Property Get RoundId()
    If Not m_Initialized Then Err.Raise(91)
    RoundId = m_RoundId
End Property

Public Property Get Groups() As StartListGroupModels
    If Not m_Initialized Then Err.Raise(91)
    Set Groups = New StartListGroupModels
    Call Groups.AddRange(CollectionUtil.SortObjectCollection(m_Groups.All(), New StartListGroupComparator))
End Property

Public Function Group(GroupNumber As Long) As StartListGroupModel
    Set Group = m_Groups.FindByRoundGroup(Me.RoundId, GroupNumber)
End Function

Public Sub AddOrder(Order As StartListOrderModel)
    Dim groupIndex As Long: groupIndex = Order.Group
    Dim groupItem As StartListGroupModel: Set groupItem = Me.Group(groupIndex)

    If (groupItem Is Nothing) Then
        Set groupItem = New StartListGroupModel
        Call groupItem.Initialize(m_RoundId, groupIndex)
        Call Me.AddGroup(groupItem)
    End If

    Call groupItem.Add(Order)  
End Sub

Public Sub AddGroup(Group As StartListGroupModel)
    Call m_Groups.Add(Group)
End Sub

Public Function EmptyOrders() As Boolean
    EmptyOrders = m_Groups.EmptyOrdersAll()
End Function

Private Function IJsonConvertable_ToJson() As String
    Dim jsonDictionary As Object: Set jsonDictionary = CreateObject("Scripting.Dictionary")
    
    Call jsonDictionary.Add("round_id", Me.RoundId)
    Call jsonDictionary.Add("items", Me.Groups)

    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(jsonDictionary)
End Function

Public Function ToJson()
    ToJson = IJsonConvertable_ToJson()
End Function
