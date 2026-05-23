VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "TrackResultOrderModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/model/result/track
Option Explicit
Implements IJsonConvertable
Private m_Id As String
Private m_RoundId As Long
Private m_GroupNumber As Long
Private m_Order As Long
Private m_Rank As Long
Private m_RealResult As TimeResultValue
Private m_Result As TimeResultValue
Private m_Remark As String
Private m_ReactionTime As Single
Private m_NotStarted As Boolean
Private m_NotFinished As Boolean
Private m_DisqualifiedReason As String
Private m_AdvancedNextRound As String
Private m_DemoEntry As Boolean
Private m_TotalRank As Long
Private m_SameResult As Long
Private m_ResultRank As Long
Private m_NextQualified As String
Private m_Score As Long
Private m_UpdatedAt As Date

' îCà”Ç≈ñÑÇﬂçûÇ›
Private m_Accumaltive As Collection
Private m_Person As AthletePersonModel


Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
    ' TODO: Write Process To Initialize
End Sub

Private Sub Class_Terminate()
    ' TODO: Write Process To Terminate
End Sub

Public Sub Initialize( _
        RoundId As Long _
        , Group As Long _
        , Order As Long _
        , Rank As Long _
        , RealResult As TimeResultValue _
        , Result As TimeResultValue _
        , Remark As String _
        , ReactionTime As Single _
        , NotStarted As Boolean _
        , NotFinished As Boolean _
        , DisqualifiedReason As String _
        , AdvancedNextRound As String _
        , DemoEntry As Boolean _
        , Score As Long _
        , UpdatedAt As Date _
        , Id As String _
        , Optional TotalRank As Long = -2147483648 _
        , Optional SameResult As Long = 0 _
        , Optional ResultRank As Long = -2147483648 _
        , Optional NextQualified As String = "" _
)
    If m_Initialized Then Err.Raise(91)

    If (Id = "") Then
        Call Err.Raise(5)
    Else
        m_Id = Id
    End If

    m_RoundId = RoundId
    m_GroupNumber = Group
    m_Order = Order
    m_Rank = Rank
    Set m_RealResult = RealResult
    Set m_Result = Result
    m_Remark = Remark
    m_ReactionTime = ReactionTime
    m_NotStarted = NotStarted
    m_NotFinished = NotFinished
    m_DisqualifiedReason = DisqualifiedReason
    m_AdvancedNextRound = AdvancedNextRound
    m_DemoEntry = DemoEntry
    m_TotalRank = TotalRank
    m_SameResult = SameResult
    m_ResultRank = ResultRank
    m_NextQualified = NextQualified
    m_Score = Score
    m_UpdatedAt = UpdatedAt

    m_Initialized = True
End Sub

Public Property Get Id() As String
    If Not m_Initialized Then Err.Raise(91)
    Id = m_Id
End Property

Public Property Get RoundId() As Long
    If Not m_Initialized Then Err.Raise(91)
    RoundId = m_RoundId
End Property

Public Property Get Group() As Long
    If Not m_Initialized Then Err.Raise(91)
    Group = m_GroupNumber
End Property

Public Property Get Order() As Long
    If Not m_Initialized Then Err.Raise(91)
    Order = m_Order
End Property

Public Property Get Rank() As Long
    If Not m_Initialized Then Err.Raise(91)
    Rank = m_Rank
End Property

Public Property Get RealResult() As TimeResultValue
    If Not m_Initialized Then Err.Raise(91)
    Set RealResult = m_RealResult
End Property

Public Property Get Result() As TimeResultValue
    If Not m_Initialized Then Err.Raise(91)
    Set Result = m_Result
End Property

Public Property Get Remark() As String
    If Not m_Initialized Then Err.Raise(91)
    Remark = m_Remark
End Property

Public Property Get ReactionTime() As Single
    If Not m_Initialized Then Err.Raise(91)
    ReactionTime = m_ReactionTime
End Property

Public Property Get NotStarted() As Boolean
    If Not m_Initialized Then Err.Raise(91)
    NotStarted = m_NotStarted
End Property

Public Property Get NotFinished() As Boolean
    If Not m_Initialized Then Err.Raise(91)
    NotFinished = m_NotFinished
End Property

Public Property Get Disqualified() As Boolean
    If Not m_Initialized Then Err.Raise(91)
    Disqualified = Not (m_DisqualifiedReason = "")
End Property

Public Property Get DisqualifiedReason() As String
    If Not m_Initialized Then Err.Raise(91)
    DisqualifiedReason = m_DisqualifiedReason
End Property

Public Property Get AdvancedNextRound() As String
    If Not m_Initialized Then Err.Raise(91)
    AdvancedNextRound = m_AdvancedNextRound
End Property

Public Function DemoEntry() As Boolean
    If Not m_Initialized Then Err.Raise(91)
    DemoEntry = m_DemoEntry
End Function

Public Property Get TotalRank() As Long
    If Not m_Initialized Then Err.Raise(91)
    TotalRank = m_TotalRank
End Property

Public Property Let TotalRank(Value As Long)
    If Not m_Initialized Then Err.Raise(91)
    m_TotalRank = Value
End Property

Public Property Get SameResult() As Long
    If Not m_Initialized Then Err.Raise(91)
    SameResult = m_SameResult
End Property

Public Property Let SameResult(Value As Long)
    If Not m_Initialized Then Err.Raise(91)
    m_SameResult = Value
End Property
    
Public Property Get ResultRank() As Long
    If Not m_Initialized Then Err.Raise(91)
    ResultRank = m_ResultRank
End Property

Public Property Let ResultRank(Value As Long)
    If Not m_Initialized Then Err.Raise(91)
    m_ResultRank = Value
End Property

Public Property Get NextQualified() As String
    If Not m_Initialized Then Err.Raise(91)
    NextQualified = m_NextQualified
End Property

Public Property Let NextQualified(Value As String)
    If Not m_Initialized Then Err.Raise(91)
    m_NextQualified = Value
End Property

Public Property Get Score() As Long
    If Not m_Initialized Then Err.Raise(91)
    Score = m_Score
End Property

Public Property Get UpdatedAt() As Date
    If Not m_Initialized Then Err.Raise(91)
    UpdatedAt = m_UpdatedAt
End Property

Public Property Get Accumaltive() As Collection
    If Not m_Initialized Then Err.Raise(91)
    Set Accumaltive = m_Accumaltive
End Property

Public Property Set Accumaltive(Value As Collection)
    If Not m_Initialized Then Err.Raise(91)
    Set m_Accumaltive = Value
End Property

Public Property Get Person() As AthletePersonModel
    If Not m_Initialized Then Err.Raise(91)
    Set Person = m_Person
End Property

Public Property Set Person(Value As AthletePersonModel)
    If Not m_Initialized Then Err.Raise(91)
    Set m_Person = Value
End Property
    
Public Function Self() 
    If Not m_Initialized Then Err.Raise(91)
    Set Self = Me
End Function

Public Function Qualified() As Boolean
    Qualified = (Me.NextQualified = "Q") Or (Me.NextQualified = "q")    
End Function

Public Function NeedToDraw() As Boolean
    NeedToDraw = (Me.NextQualified = "?")
End Function

Public Function IsRankable()
    If (Me.DemoEntry Or Me.NotStarted Or Me.NotFinished Or Not(Me.DisqualifiedReason = "")) Then
        IsRankable = False
        Exit Function
    End If

    IsRankable = Me.RealResult.Valid()
End Function

Private Function IJsonConvertable_ToJson() As String
    Dim jsonDict As Object: Set jsonDict = CreateObject("Scripting.Dictionary")

    Call jsonDict.Add("id", Me.Id)
    Call jsonDict.Add("round_id", Me.RoundId)
    Call jsonDict.Add("group_number", Me.Group)
    Call jsonDict.Add("order", Me.Order)

    If (Me.IsRankable()) Then
        Call jsonDict.Add("rank", Me.Rank)
    Else
        Call jsonDict.Add("rank", "")
    End If

    Call jsonDict.Add("real_result", Me.RealResult.ToString())
    Call jsonDict.Add("result", Me.Result.ToString())
    Call jsonDict.Add("remark", Me.Remark)
    Call jsonDict.Add("reaction_time", Me.ReactionTime)
    Call jsonDict.Add("not_started", Me.NotStarted)
    Call jsonDict.Add("not_finished", Me.NotFinished)
    Call jsonDict.Add("disqualified", Me.Disqualified)
    Call jsonDict.Add("disqualified_reason", Me.DisqualifiedReason)
    Call jsonDict.Add("advanced_next_round", Me.AdvancedNextRound)
    Call jsonDict.Add("demo_entry", Me.DemoEntry)
    Call jsonDict.Add("total_rank", Me.TotalRank)
    Call jsonDict.Add("same_result", Me.SameResult)
    Call jsonDict.Add("result_rank", Me.ResultRank)
    Call jsonDict.Add("next_qualified", Me.NextQualified)
    Call jsonDict.Add("score", Me.Score)
    Call jsonDict.Add("updated_at", Me.UpdatedAt)

    If Not (Me.Accumaltive Is Nothing) Then
        Call jsonDict.Add("accumaltive", Me.Accumaltive)        
    End If

    If Not (Me.Person Is Nothing) Then
        Call jsonDict.Add("person", Me.Person)
    End If

    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(jsonDict)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function
