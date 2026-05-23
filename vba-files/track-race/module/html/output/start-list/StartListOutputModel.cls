VERSION 1.0 CLASS
  BEGIN
  MultiUse = -1  'True
  END
  Attribute VB_Name = "StartListOutputModel"
  Attribute VB_GlobalNameSpace = False
  Attribute VB_Creatable = False
  Attribute VB_PredeclaredId = False
  Attribute VB_Exposed = False
'namespace=vba-files/track-race/modules/html/output/start-list
Option Explicit
Implements IJsonConvertable

Private m_TrackEvent As TrackEventSettingModel
Private m_Rounds As EventRoundModels
Private m_GameRecords As GameRecordModels
Private m_CurrentRound As EventRoundModel
Private m_Groups As StartListRoundModel

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
    ' TODO: Write Process To Initialize
End Sub

Private Sub Class_Terminate()
    ' TODO: Write Process To Terminate
End Sub

Public Sub Initialize(RoundName As String, TrackEvent As TrackEventSettingModel, Rounds As EventRoundModels, GameRecords As GameRecordModels, Groups As StartListRoundModel)
    If m_Initialized Then Err.Raise(91)
    Set m_TrackEvent = TrackEvent
    Set m_Rounds = Rounds
    Set m_GameRecords = GameRecords
    Set m_CurrentRound = Rounds.FindByName(RoundName)
    Set m_Groups = Groups
    m_Initialized = True
End Sub

Public Property Get TrackEvent() As TrackEventSettingModel
    If Not m_Initialized Then Err.Raise(91)
    Set TrackEvent = m_TrackEvent
End Property

Public Property Get Rounds() As EventRoundModels
    If Not m_Initialized Then Err.Raise(91)
    Set Rounds = m_Rounds
End Property

Public Property Get GameRecords() As GameRecordModels
    If Not m_Initialized Then Err.Raise(91)
    Set GameRecords = m_GameRecords
End Property

Public Property Get CurrentRound() As EventRoundModel
    If Not m_Initialized Then Err.Raise(91)
    Set CurrentRound = m_CurrentRound
End Property

Public Property Get Groups() As StartListRoundModel
    If Not m_Initialized Then Err.Raise(91)
    Set Groups = m_Groups
End Property

Public Function Self() As StartListOutputModel
    If Not m_Initialized Then Err.Raise(91)
    Set Self = Me
End Function

Public Function ToJson() As String
    Dim jsonDict As Object: Set jsonDict = CreateObject("Scripting.Dictionary")

    Call jsonDict.Add("track_event", Me.TrackEvent)
    Call jsonDict.Add("rounds", Me.Rounds)
    Call jsonDict.Add("game_records", Me.GameRecords)
    Call jsonDict.Add("current_round", Me.CurrentRound)
    Call jsonDict.Add("groups", Me.Groups)

    ToJson = JsonConverter.ConvertToJson(jsonDict)
End Function

Private Function IJsonConvertable_ToJson() As String
    IJsonConvertable_ToJson = Me.ToJson()
End Function
