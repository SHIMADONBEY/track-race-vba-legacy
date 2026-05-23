VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "TrackResultRankingModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/modules/ranking/model
Option Explicit
Implements IJsonConvertable

Private m_Id As String
Private m_RoundId As Long
Private m_Group As Long
Private m_Rank As Long
Private m_RealResult As TimeResultValue
Private m_OfficialResult As TimeResultValue
Private m_QualifiedByRank As Boolean
Private m_QualifiedByResult As Boolean
Private m_DrawToQualify As Boolean
Private m_TotalRanking As Long
Private m_SameResult As Long

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
    ' TODO: Write Process To Initialize
End Sub

Private Sub Class_Terminate()
    ' TODO: Write Process To Terminate
End Sub

Public Sub Initialize( _
        Id As String _
        , RoundId As Long _
        , Group As Long _
        , Rank As Long _
        , RealResult As TimeResultValue _
        , OfficialResult As TimeResultValue _
        , Optional QualifiedByRank As Boolean = False _
        , Optional QualifiedByResult As Boolean = False _
        , Optional DrawToQualify As Boolean = False _
        , Optional TotalRanking As Long = -2147483648 _
        , Optional SameResult As Long = 0 _
)
    If m_Initialized Then Err.Raise(91)
    m_Id = Id
    m_RoundId = RoundId
    m_Group = Group
    m_Rank = Rank
    Set m_RealResult = RealResult
    Set m_OfficialResult = OfficialResult
    m_QualifiedByRank = QualifiedByRank
    m_QualifiedByResult = QualifiedByResult
    m_DrawToQualify = DrawToQualify
    m_TotalRanking = TotalRanking
    m_SameResult = SameResult
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
    Group = m_Group
End Property

Public Property Get Rank() As Long
    If Not m_Initialized Then Err.Raise(91)
    Rank = m_Rank
End Property

Public Property Get RealResult() As TimeResultValue
    If Not m_Initialized Then Err.Raise(91)
    Set RealResult = m_RealResult
End Property

Public Property Get OfficialResult() As TimeResultValue
    If Not m_Initialized Then Err.Raise(91)
    Set OfficialResult = m_OfficialResult
End Property

Public Property Get QualifiedByRank() As Boolean
    If Not m_Initialized Then Err.Raise(91)
    QualifiedByRank = m_QualifiedByRank
End Property

Public Property Let QualifiedByRank(Value As Boolean)
    If Not m_Initialized Then Err.Raise(91)
    m_QualifiedByRank = Value
End Property

Public Property Get QualifiedByResult() As Boolean
    If Not m_Initialized Then Err.Raise(91)
    QualifiedByResult = m_QualifiedByResult
End Property

Public Property Let QualifiedByResult(Value As Boolean)
    If Not m_Initialized Then Err.Raise(91)
    m_QualifiedByResult = Value
End Property

Public Property Get DrawToQualify() As Boolean
    If Not m_Initialized Then Err.Raise(91)
    DrawToQualify = m_DrawToQualify
End Property

Public Property Let DrawToQualify(Value As Boolean)
    If Not m_Initialized Then Err.Raise(91)
    m_DrawToQualify = Value
End Property

Public Property Get TotalRanking() As Long
    If Not m_Initialized Then Err.Raise(91)
    TotalRanking = m_TotalRanking
End Property

Public Property Let TotalRanking(Value As Long)
    If Not m_Initialized Then Err.Raise(91)
    m_TotalRanking = Value
End Property

Public Property Get SameResult() As Long
    If Not m_Initialized Then Err.Raise(91)
    SameResult = m_SameResult
End Property

Public Property Let SameResult(Value As Long)
    If Not m_Initialized Then Err.Raise(91)
    m_SameResult = Value
End Property

Public Function Self() 
    If Not m_Initialized Then Err.Raise(91)
    Set Self = Me
End Function

Private Function IJsonConvertable_ToJson() As String
    Dim jsonDictionary As Object: Set jsonDictionary = CreateObject("Scripting.Dictionary")

    Call jsonDictionary.Add("id", Me.Id)
    Call jsonDictionary.Add("round_id", Me.RoundId)
    Call jsonDictionary.Add("group", Me.Group)
    Call jsonDictionary.Add("rank", Me.Rank)
    Call jsonDictionary.Add("real_result", Me.RealResult)
    Call jsonDictionary.Add("official_result", Me.OfficialResult)
    Call jsonDictionary.Add("qualified_by_rank", Me.QualifiedByRank)
    Call jsonDictionary.Add("qualified_by_result", Me.QualifiedByResult)
    Call jsonDictionary.Add("draw_to_qualify", Me.DrawToQualify)
    Call jsonDictionary.Add("total_ranking", Me.TotalRanking)
    Call jsonDictionary.Add("same_result", Me.SameResult)

    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(jsonDictionary)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function

Public Function Clone() As TrackResultRankingModel
    If Not m_Initialized Then Err.Raise(91)
    Set Clone = New TrackResultRankingModel
    Call Clone.Initialize( _
            Me.Id _
            , Me.RoundId _
            , Me.Group _
            , Me.Rank _
            , Me.RealResult _
            , Me.OfficialResult _
            , Me.QualifiedByRank _
            , Me.QualifiedByResult _
            , Me.DrawToQualify _
            , Me.TotalRanking _
            , Me.SameResult _
    )
End Function
