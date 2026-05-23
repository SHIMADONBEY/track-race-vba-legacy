VERSION 1.0 CLASS
BEGIN
MultiUse = -1 ' True
END
Attribute VB_Name = "GameRecordModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/common/model/game-record
Option Explicit
Implements IJsonConvertable

Private m_EventId As Long
Private m_Serial As Long
Private m_Name As String
Private m_Abbrebiation As String
Private m_Record As IResultValue
Private m_AthleteName As String
Private m_TeamName As String
Private m_TeamPlace As String
Private m_TeamCountry As String
Private m_RecordedYear As String
Private m_RecordedMonth As String
Private m_RecordedDay As String
Private m_RecordScale As Long
Private m_TargetGrade As String
Private m_Comment As String
Private m_RecordedYmd As String
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
        , Serial As Long _
        , Name As String _
        , Abbrebiation As String _
        , Record As IResultValue _
        , AthleteName As String _
        , TeamName As String _
        , TeamPlace As String _
        , TeamCountry As String _
        , RecordedYear As String _
        , RecordedMonth As String _
        , RecordedDay As String _
        , RecordScale As Long _
        , TargetGrade As String _
        , Comment As String _
        , RecordedYmd As String _
        , Optional Id As String = "" _
    )

    If m_Initialized Then Err.Raise(91)

    If (Name = "" Or TeamName = "") Then
        Err.Raise 5, "GameRecordModel.Initialize"
    End If

    If (Record Is Nothing) Then 
        Err.Raise 5, "GameRecordModel.Initialize"
    ElseIf (Record.Valid = False) Then
        Err.Raise 5, "GameRecordModel.Initialize"
    End If

    Dim pId As String
    If (Id = "") Then 
        pId = UuidFactory.GenerateUuid()
    Else
        pId = Id
    End If

    m_EventId = EventId
    m_Serial = Serial
    m_Name = Name
    m_Abbrebiation = Abbrebiation
    Set m_Record = Record
    m_AthleteName = AthleteName
    m_TeamName = TeamName
    m_TeamPlace = TeamPlace
    m_TeamCountry = TeamCountry
    m_RecordedYear = RecordedYear
    m_RecordedMonth = RecordedMonth
    m_RecordedDay = RecordedDay
    m_RecordScale = RecordScale
    m_TargetGrade = TargetGrade
    m_Comment = Comment
    m_RecordedYmd = RecordedYmd
    m_Id = pId
    m_Initialized = True
End Sub

Public Property Get EventId() As Long
    If Not m_Initialized Then Err.Raise(91)
    EventId = m_EventId
End Property

Public Function Serial() As Long
    If Not m_Initialized Then Err.Raise(91)
    Serial = m_Serial
End Function

Public Property Get Name() As String
    If Not m_Initialized Then Err.Raise(91)
    Name = m_Name
End Property

Public Property Get Abbrebiation() As String
    If Not m_Initialized Then Err.Raise(91)
    Abbrebiation = m_Abbrebiation
End Property

Public Property Get Record() As IResultValue
    If Not m_Initialized Then Err.Raise(91)
    Set Record = m_Record
End Property

Public Property Get AthleteName() As String
    If Not m_Initialized Then Err.Raise(91)
    AthleteName = m_AthleteName
End Property

Public Property Get TeamName() As String
    If Not m_Initialized Then Err.Raise(91)
    TeamName = m_TeamName
End Property

Public Property Get TeamPlace() As String
    If Not m_Initialized Then Err.Raise(91)
    TeamPlace = m_TeamPlace
End Property

Public Property Get TeamCountry() As String
    If Not m_Initialized Then Err.Raise(91)
    TeamCountry = m_TeamCountry
End Property

Public Property Get RecordedYear() As String
    If Not m_Initialized Then Err.Raise(91)
    RecordedYear = m_RecordedYear
End Property

Public Property Get RecordedMonth() As String
    If Not m_Initialized Then Err.Raise(91)
    RecordedMonth = m_RecordedMonth
End Property

Public Property Get RecordedDay() As String
    If Not m_Initialized Then Err.Raise(91)
    RecordedDay = m_RecordedDay
End Property

Public Property Get RecordScale() As Long
    If Not m_Initialized Then Err.Raise(91)
    RecordScale = m_RecordScale
End Property

Public Property Get TargetGrade() As String
    If Not m_Initialized Then Err.Raise(91)
    TargetGrade = m_TargetGrade
End Property

Public Property Get Comment() As String
    If Not m_Initialized Then Err.Raise(91)
    Comment = m_Comment
End Property

Public Property Get RecordedYmd() As String
    If Not m_Initialized Then Err.Raise(91)
    RecordedYmd = m_RecordedYmd
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
    Call jsonDictionary.Add("serial", Me.Serial)
    Call jsonDictionary.Add("name", Me.Name)
    Call jsonDictionary.Add("abbrebiation", Me.Abbrebiation)
    Call jsonDictionary.Add("record", Me.Record.ToString())
    Call jsonDictionary.Add("athlete_name", Me.AthleteName)
    Call jsonDictionary.Add("team_name", Me.TeamName)
    Call jsonDictionary.Add("team_place", Me.TeamPlace)
    Call jsonDictionary.Add("team_country", Me.TeamCountry)
    Call jsonDictionary.Add("recorded_year", Me.RecordedYear)
    Call jsonDictionary.Add("recorded_month", Me.RecordedMonth)
    Call jsonDictionary.Add("recorded_day", Me.RecordedDay)
    Call jsonDictionary.Add("record_scale", Me.RecordScale)
    Call jsonDictionary.Add("target_grade", Me.TargetGrade)
    Call jsonDictionary.Add("comment", Me.Comment)
    Call jsonDictionary.Add("recorded_ymd", Me.RecordedYmd)
    Call jsonDictionary.Add("id", Me.Id)
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(jsonDictionary)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function
