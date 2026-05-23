VERSION 1.0 CLASS
BEGIN
MultiUse = -1 ' True
END
Attribute VB_Name = "CompetitionRoundModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/common/model/round
Option Explicit
Implements IJsonConvertable
Implements IJsonParsable

Private m_EventId As Long
Private m_RoundId As Long
Private m_Name As String
Private m_StartDateTime As Date
Private m_OrderCount As Long
Private m_GroupCount As Long
Private m_GroupStrategyType As Long
Private m_OrderStrategyType As Long
Private m_PlaceEachGroup As Long
Private m_QuarifyingStandard As Variant
Private m_AdditionCount As Long
Private m_EventArea As String
Private m_EndDateTime As Date
Private m_GroupsPerCall As Long
Private m_Id As String

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
    ' TODO: Write Process To Initialize
End Sub

Private Sub Class_Terminate()
    ' TODO: Write Process To Terminate
End Sub

Public Sub Initialize( _
        EventId As Long _
        , RoundId As Long _
        , Name As String _
        , StartDateTime As Date _
        , OrderCount As Long _
        , GroupCount As Long _
        , GroupStrategyType As Long _
        , OrderStrategyType As Long _
        , PlaceEachGroup As Long _
        , QuarifyingStandard As Variant _
        , AdditionCount As Long _
        , EventArea As String _
        , EndDateTime As Date _
        , GroupsPerCall As Long _
        , Optional Id As String = "" _
    )
    If m_Initialized Then Err.Raise(91)
    m_EventId = EventId
    m_RoundId = RoundId
    m_Name = Name
    m_StartDateTime = StartDateTime
    m_OrderCount = OrderCount
    m_GroupCount = GroupCount
    m_GroupStrategyType = GroupStrategyType
    m_OrderStrategyType = OrderStrategyType
    m_PlaceEachGroup = PlaceEachGroup
    m_QuarifyingStandard = QuarifyingStandard
    m_AdditionCount = AdditionCount
    m_EventArea = EventArea
    m_EndDateTime = EndDateTime
    m_GroupsPerCall = GroupsPerCall

    If (Id = "") Then 
        m_Id = UuidFactory.GenerateUuid()
    Else
        m_Id = Id
    End If

    m_Initialized = True
End Sub

Public Property Get EventId() As Long
    If Not m_Initialized Then Err.Raise(91)
    EventId = m_EventId
End Property

Public Property Get RoundId() As Long
    If Not m_Initialized Then Err.Raise(91)
    RoundId = m_RoundId
End Property

Public Property Get Name() As String
    If Not m_Initialized Then Err.Raise(91)
    Name = m_Name
End Property

Public Property Get StartDateTime() As Date
    If Not m_Initialized Then Err.Raise(91)
    StartDateTime = m_StartDateTime
End Property

Public Property Get OrderCount() As Long
    If Not m_Initialized Then Err.Raise(91)
    OrderCount = m_OrderCount
End Property

Public Property Get GroupCount() As Long
    If Not m_Initialized Then Err.Raise(91)
    GroupCount = m_GroupCount
End Property

Public Property Get GroupStrategyType() As Long
    If Not m_Initialized Then Err.Raise(91)
    GroupStrategyType = m_GroupStrategyType
End Property

Public Property Get OrderStrategyType() As Long
    If Not m_Initialized Then Err.Raise(91)
    OrderStrategyType = m_OrderStrategyType
End Property

Public Property Get PlaceEachGroup() As Long
    If Not m_Initialized Then Err.Raise(91)
    PlaceEachGroup = m_PlaceEachGroup
End Property

Public Property Get QuarifyingStandard() As Variant
    If Not m_Initialized Then Err.Raise(91)
    QuarifyingStandard = m_QuarifyingStandard
End Property

Public Property Get AdditionCount() As Long
    If Not m_Initialized Then Err.Raise(91)
    AdditionCount = m_AdditionCount
End Property

Public Property Get EventArea() As String
    If Not m_Initialized Then Err.Raise(91)
    EventArea = m_EventArea
End Property

Public Property Get EndDateTime() As Date
    If Not m_Initialized Then Err.Raise(91)
    EndDateTime = m_EndDateTime
End Property

Public Property Get GroupsPerCall() As Long
    If Not m_Initialized Then Err.Raise(91)
    GroupsPerCall = m_GroupsPerCall
End Property

Public Property Get Id() As String
    If Not m_Initialized Then Err.Raise(91)
    Id = m_Id
End Property

Public Function Self() 
    If Not m_Initialized Then Err.Raise(91)
    Set Self = Me
End Function

Private Function IJsonConvertable_ToJson() As String
    Dim jsonDictionary As Object: Set jsonDictionary = CreateObject("Scripting.Dictionary")
    Call jsonDictionary.Add("event_id", Me.EventId)
    Call jsonDictionary.Add("round_id", Me.RoundId)
    Call jsonDictionary.Add("name", Me.Name)
    Call jsonDictionary.Add("start_date_time", Me.StartDateTime)
    Call jsonDictionary.Add("order_count", Me.OrderCount)
    Call jsonDictionary.Add("group_count", Me.GroupCount)
    Call jsonDictionary.Add("group_strategy_type", Me.GroupStrategyType)
    Call jsonDictionary.Add("order_strategy_type", Me.OrderStrategyType)
    Call jsonDictionary.Add("place_each_group", Me.PlaceEachGroup)
    Call jsonDictionary.Add("quarifying_standard", Me.QuarifyingStandard)
    Call jsonDictionary.Add("addition_count", Me.AdditionCount)
    Call jsonDictionary.Add("event_area", Me.EventArea)
    Call jsonDictionary.Add("end_date_time", Me.EndDateTime)
    Call jsonDictionary.Add("groups_per_call", Me.GroupsPerCall)
    Call jsonDictionary.Add("id", Me.Id)
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(jsonDictionary)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function

Public Sub FromJson(JsonString As String)
    Call IJsonParsable_FromJson(JsonString)
End Sub

Private Sub IJsonParsable_FromJson(JsonString As String)
    Debug.Print JsonString

    With JsonConverter.ParseJson(JsonString)
        Call Me.Initialize( _
                .Item("event_id") _
                , .Item("round_id") _
                , .Item("name") _
                , JsonConverter.ParseIso(.Item("start_date_time")) _
                , .Item("order_count") _
                , .Item("group_count") _
                , .Item("group_strategy_type") _
                , .Item("order_strategy_type") _
                , .Item("place_each_group") _
                , .Item("quarifying_standard") _
                , .Item("addition_count") _
                , .Item("event_area") _
                , JsonConverter.ParseIso(.Item("end_date_time")) _
                , .Item("groups_per_call") _
                , .Item("id") _
        )
    End With
End Sub
