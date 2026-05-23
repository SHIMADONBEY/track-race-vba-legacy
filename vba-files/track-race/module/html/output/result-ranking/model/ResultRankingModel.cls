VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "ResultRankingModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/modules/html/result-ranking/model
Option Explicit
Implements IJsonConvertable

Private m_Id As String
Private m_RoundId As Long
Private m_GroupNumber As Long
Private m_Order As Long
Private m_Rank As Long
Private m_Result As TimeResultValue
Private m_RealResult As TimeResultValue
Private m_Remark As String
Private m_ReactionTime As Single
Private m_NotStarted As Boolean
Private m_NotFinished As Boolean
Private m_Disqualified As Boolean
Private m_DisqualifiedReason As String
Private m_AdvancedNextRound As String
Private m_DemoEntry As Boolean
Private m_TotalRank As Long
Private m_SameResult As Long
Private m_ResultRank As Long
Private m_NextQualified As String
Private m_Score As Long
Private m_Accumaltive As Collection
Private m_Person As AthletePersonModel
Private m_Wind As WindValue

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
    ' TODO: Write Process To Initialize
End Sub

Private Sub Class_Terminate()
    ' TODO: Write Process To Terminate
End Sub

Public Sub Initialize(Id As String, RoundId As Long, GroupNumber As Long, Order As Long, Rank As Long, Result As TimeResultValue, RealResult As TimeResultValue, Remark As String, ReactionTime As Single, NotStarted As Boolean, NotFinished As Boolean, DisqualifiedReason As String, AdvancedNextRound As String, DemoEntry As Boolean, TotalRank As Long, SameResult As Long, ResultRank As Long, NextQualified As String, Score As Long, Accumaltive As Collection, Person As AthletePersonModel, Wind As WindValue)
    If m_Initialized Then Err.Raise(91)
    m_Id = Id
    m_RoundId = RoundId
    m_GroupNumber = GroupNumber
    m_Order = Order
    m_Rank = Rank
    Set m_Result = Result
    Set m_RealResult = RealResult
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
    Set m_Accumaltive = Accumaltive
    Set m_Person = Person
    Set m_Wind = Wind
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

Public Property Get GroupNumber() As Long
    If Not m_Initialized Then Err.Raise(91)
    GroupNumber = m_GroupNumber
End Property

Public Property Get Order() As Long
    If Not m_Initialized Then Err.Raise(91)
    Order = m_Order
End Property

Public Property Get Rank() As Long
    If Not m_Initialized Then Err.Raise(91)
    Rank = m_Rank
End Property

Public Property Get Result() As TimeResultValue
    If Not m_Initialized Then Err.Raise(91)
    Set Result = m_Result
End Property

Public Property Get RealResult() As TimeResultValue
    If Not m_Initialized Then Err.Raise(91)
    Set RealResult = m_RealResult
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

Public Property Get DemoEntry() As Boolean
    If Not m_Initialized Then Err.Raise(91)
    DemoEntry = m_DemoEntry
End Property

Public Property Get TotalRank() As Long
    If Not m_Initialized Then Err.Raise(91)
    TotalRank = m_TotalRank
End Property

Public Property Get SameResult() As Long
    If Not m_Initialized Then Err.Raise(91)
    SameResult = m_SameResult
End Property
    
Public Property Get ResultRank() As Long
    If Not m_Initialized Then Err.Raise(91)
    ResultRank = m_ResultRank
End Property

Public Property Get NextQualified() As String
    If Not m_Initialized Then Err.Raise(91)
    NextQualified = m_NextQualified
End Property

Public Property Get Score() As Long
    If Not m_Initialized Then Err.Raise(91)
    Score = m_Score
End Property

Public Property Get Accumaltive() As Collection
    If Not m_Initialized Then Err.Raise(91)
    Set Accumaltive = m_Accumaltive    
End Property

Public Property Get Person() As AthletePersonModel
    If Not m_Initialized Then Err.Raise(91)
    Set Person = m_Person
End Property

Public Property Get Wind() As WindValue
    If Not m_Initialized Then Err.Raise(91)
    Set Wind = m_Wind
End Property

Public Function Self() As ResultRankingModel
    If Not m_Initialized Then Err.Raise(91)
    Set Self = Me
End Function

Private Function IJsonConvertable_ToJson() As String
    Dim jsonDictionary As Object: Set jsonDictionary = CreateObject("Scripting.Dictionary")
    Call jsonDictionary.Add("id", Me.Id)
    Call jsonDictionary.Add("round_id", Me.RoundId)
    Call jsonDictionary.Add("group_number", Me.GroupNumber)
    Call jsonDictionary.Add("order", Me.Order)
    Call jsonDictionary.Add("rank", Me.Rank)
    Call jsonDictionary.Add("result", Me.Result.ToString())
    Call jsonDictionary.Add("real_result", Me.RealResult.ToString())
    Call jsonDictionary.Add("remark", Me.Remark)
    Call jsonDictionary.Add("reaction_time", Me.ReactionTime)
    Call jsonDictionary.Add("not_started", Me.NotStarted)
    Call jsonDictionary.Add("not_finished", Me.NotFinished)
    Call jsonDictionary.Add("disqualified", Me.Disqualified)
    Call jsonDictionary.Add("disqualified_reason", Me.DisqualifiedReason)
    Call jsonDictionary.Add("advanced_next_round", Me.AdvancedNextRound)
    Call jsonDictionary.Add("demo_entry", Me.DemoEntry)
    Call jsonDictionary.Add("total_rank", Me.TotalRank)
    Call jsonDictionary.Add("same_result", Me.SameResult)
    Call jsonDictionary.Add("result_rank", Me.ResultRank)
    Call jsonDictionary.Add("next_qualified", Me.NextQualified)
    Call jsonDictionary.Add("score", Me.Score)
    Call jsonDictionary.Add("accumaltive", Me.Accumaltive)
    Call jsonDictionary.Add("person", Me.Person)

    If Not (Me.Wind Is Nothing) Then
        Call jsonDictionary.Add("wind", Me.Wind.Value)
    End If

    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(jsonDictionary)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function
