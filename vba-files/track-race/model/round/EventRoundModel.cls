VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "EventRoundModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba_files/track-race/model/round
Option Explicit
Implements IJsonConvertable

Private m_Initialized As Boolean
Private m_Id As Long
Private m_Name As String
Private m_StartDateTime As Date
Private m_OrderCount As Long
Private m_GroupCount As Long
Private m_GroupStrategyType As Long
Private m_OrderStrategyType As Long
Private m_PlaceEachGroup As Long
Private m_AdditionCount As Long

Private  Sub Class_Initialize()
    m_Initialized = false
End Sub

Public Sub Initialize( _
        Id As Long _
        , Name As String _
        , StartDateTime As Date _
        , OrderCount As Integer _
        , GroupCount As Integer _
        , GroupStrategyType As Integer _
        , OrderStrategyType As Integer _
        , PlaceEachGroup As Integer _
        , AdditionCount As Integer _
    )

    If m_Initialized Then Err.Raise(91)
    m_Id = Id
    m_Name = Name
    m_StartDateTime = StartDateTime
    m_OrderCount = OrderCount
    m_GroupCount = GroupCount
    m_GroupStrategyType = GroupStrategyType
    m_OrderStrategyType = OrderStrategyType
    m_PlaceEachGroup = PlaceEachGroup
    m_AdditionCount = AdditionCount
    m_Initialized = True
End Sub

Public Sub FromCompetitionRound(Information As CompetitionRoundModel)
    If m_Initailized Then Err.Raise(91)

    m_Id = Information.EventId
    m_Name = Information.Name
    m_StartDateTime = Information.StartDateTime
    m_OrderCount = Information.OrderCount
    m_GroupCount = Information.GroupCount
    m_GroupStrategyType = Information.GroupStrategyType
    m_OrderStrategyType = Information.OrderStrategyType
    m_PlaceEachGroup = Information.PlaceEachGroup
    m_AdditionCount = Information.AdditionCount

    m_Initailized = True
End Sub

Public Property Get Id() As Long
    If Not m_Initialized Then Err.Raise(91)
    Id = m_Id
End Property

Public Property Get Name() As String
    If Not m_Initialized Then Err.Raise(91)
    Name = m_Name
End Property

Public Property Get StartDateTime() As Date
    If Not m_Initialized Then Err.Raise(91)
    StartDateTime = m_StartDateTime
End Property

Public Property Get OrderCount()
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

Public Property Get AdditionCount() As Long
    If Not m_Initialized Then Err.Raise(91)
    AdditionCount = m_AdditionCount
End Property

Public Function NextQualifiersCount() As Long
    NextQualifiersCount = Me.GroupCount * Me.PlaceEachGroup + Me.AdditionCount
End Function

Public Function Self() As EventRoundModel
    Set Self = Me
End Function

Public Function ToJson() As String
    Dim jsonDict As Object: Set jsonDict = CreateObject("Scripting.Dictionary")

    Call jsonDict.Add("id", Me.Id)
    Call jsonDict.Add("name", Me.Name)
    Call jsonDict.Add("start_date_time", Me.StartDateTime)
    Call jsonDict.Add("order_count", Me.OrderCount)
    Call jsonDict.Add("group_count", Me.GroupCount)
    Call jsonDict.Add("group_strategy_type", Me.GroupStrategyType)
    Call jsonDict.Add("order_strategy_type", Me.OrderStrategyType)
    Call jsonDict.Add("place_each_group", Me.PlaceEachGroup)
    Call jsonDict.Add("addition_count", Me.AdditionCount)

    ToJson = JsonConverter.ConvertToJson(jsonDict)
End Function

Private Function IJsonConvertable_ToJson() As String
    IJsonConvertable_ToJson = Me.ToJson()
End Function
    
