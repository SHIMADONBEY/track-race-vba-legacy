VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "StartListGroupModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/model/start-list
Option Explicit
Implements IJsonConvertable
Private m_Id As String
Private m_RoundId As Long
Private m_GroupNumber As Long
Private m_Orders As StartListOrderModels

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
    ' Write Process To Initialize
    Set m_Orders = New StartListOrderModels
End Sub

Private Sub Class_Terminate()
    ' Write Process To Terminate
    Set m_Orders = Nothing
End Sub

Public Sub Initialize(RoundId As Long, GroupNumber As Long, Optional OrderItems As StartListOrderModels = Nothing)
    If m_Initialized Then Err.Raise(91)
    m_Id = UuidFactory.GenerateUuid()
    m_RoundId = RoundId
    m_GroupNumber = GroupNumber

    If Not(OrderItems Is Nothing) Then
        Call m_Orders.AddRange(OrderItems.All())
    End If

    m_Initialized = True
End Sub

Public Property Get Id() As String
    If Not m_Initialized Then Err.Raise(91)
    Id = m_Id
End Property

Public Property Get RoundId()
    If Not m_Initialized Then Err.Raise(91)
    RoundId = m_RoundId
End Property

Public Property Get GroupNumber() As Long
    If Not m_Initialized Then Err.Raise(91)
    GroupNumber = m_GroupNumber
End Property

Public Property Get Orders() As StartListOrderModels
    If Not m_Initialized Then Err.Raise(91)

    Set Orders = New StartListOrderModels
    Call Orders.AddRange(CollectionUtil.SortObjectCollection(m_Orders.All(), New StartListGroupOrderComparator))
End Property

Public Function CountOrders() As Long
    If Not m_Initialized Then Err.Raise(91)
    CountOrders = m_Orders.Count()
End Function

Public Sub Add(OrderItem As StartListOrderModel)
    If Not m_Initialized Then Err.Raise(91)
    Call m_Orders.Add(OrderItem)
End Sub

Public Sub AddRange(Items As Collection)
    If Not m_Initialized Then Err.Raise(91)
    Call m_Orders.AddRange(Items)
End Sub

Public Function FindByOrder(Order As Long) As StartListOrderModel
    Set FindByOrder = m_Orders.FindOrder(Me.RoundId, Me.GroupNumber, Order)
End Function

Public Function EmptyOrders() As Boolean
    EmptyOrders = Me.CountOrders() < 1
End Function

Public Function Self() As StartListGroupModel
    If Not m_Initialized Then Err.Raise(91)
    Set Self = Me
End Function

Private Function IJsonConvertable_ToJson() As String
    Dim jsonDictionary As Object: Set jsonDictionary = CreateObject("Scripting.Dictionary")
    
    Call jsonDictionary.Add("id", Me.Id)
    Call jsonDictionary.Add("round_id", Me.RoundId)
    Call jsonDictionary.Add("group_number", Me.GroupNumber)
    Call jsonDictionary.Add("orders", Me.Orders)

    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(jsonDictionary)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function

Private Function GenerateOrderKey(OrderNumber As Long)
    GenerateOrderKey = "__" & OrderNumber
End Function