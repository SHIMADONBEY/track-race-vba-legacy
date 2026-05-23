VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TrackEventSettingModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/module/event/model
Option Explicit
Implements IJsonConvertable

Private m_CompetitionEvent As CompetitionEventModel

Private m_ResultPrecision As Long
Private m_MinuteDelimiter As String
Private m_SecondDelimiter As String
Private m_SplitTimePrecision As Long
Private m_TrackRaceLength As Long
Private m_IntervalLength1 As Long
Private m_IntervalLength2 As Long
Private m_ScoringCombinedEvents As Boolean
Private m_ScoringIndexNumber1 As String
Private m_ScoringIndexNumber2 As String
Private m_ScoringIndexNumber3 As String
Private m_LaneSetting As LaneSettingModel
Private m_HtmlTemplate As HtmlTemplatePathsModel

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
End Sub

Private Sub Class_Terminate()
End Sub

Public Sub Initialize(Category As String, Sex As String, EventName As String, Supecification As String, PersonPerGroup As Long, RoundsCount As Long, OperationWindGauge As Boolean, MeasurementReactionTime As Boolean, ResultPrecision As Long, MinuteDelimiter As String, SecondDelimiter As String, TrackRaceLength As Long, SplitTimePrecision As Long, IntervalLength1 As Long, IntervalLength2 As Long, ScoringCombinedEvents As Boolean, ScoringIndexNumber1 As String, ScoringIndexNumber2 As String, ScoringIndexNumber3 As String, LaneSetting As LaneSettingModel, HtmlTemplate As HtmlTemplatePathsModel)
    If m_Initialized Then Err.Raise (91)

    Set m_CompetitionEvent = New CompetitionEventModel
    Call m_CompetitionEvent.Initialize( _
            -2147483648 _
            , Category _
            , Sex _
            , EventName _
            , Supecification _
            , EventTypeConstants.TRACK _
            , RoundsCount _
            , PersonPerGroup _
            , OperationWindGauge _
            , MeasurementReactionTime _
            , 0 _
            , 0 _
            , 0 _
            , "track-race-vba.xlsm" _
            , "TRK.csv" _
    )

    m_ResultPrecision = ResultPrecision
    m_MinuteDelimiter = MinuteDelimiter
    m_SecondDelimiter = SecondDelimiter
    m_TrackRaceLength = TrackRaceLength
    m_SplitTimePrecision = SplitTimePrecision
    m_IntervalLength1 = IntervalLength1
    m_IntervalLength2 = IntervalLength2
    m_ScoringCombinedEvents = ScoringCombinedEvents
    m_ScoringIndexNumber1 = ScoringIndexNumber1
    m_ScoringIndexNumber2 = ScoringIndexNumber2
    m_ScoringIndexNumber3 = ScoringIndexNumber3
    Set m_LaneSetting = LaneSetting
    Set m_HtmlTemplate = HtmlTemplate

    m_Initialized = True
End Sub

Public Property Get CompetitionEvent() As CompetitionEventModel
    If Not m_Initialized Then Err.Raise (91)
    Set CompetitionEvent = m_CompetitionEvent
End Property

Public Property Get Category() As String
    If Not m_Initialized Then Err.Raise (91)
    Category = m_CompetitionEvent.Category
End Property

Public Property Get Sex() As String
    If Not m_Initialized Then Err.Raise (91)
    Sex = m_CompetitionEvent.Sex
End Property

Public Property Get EventName() As String
    If Not m_Initialized Then Err.Raise (91)
    EventName = m_CompetitionEvent.EventName
End Property

Public Property Get Supecification() As String
    If Not m_Initialized Then Err.Raise (91)
    Supecification = m_CompetitionEvent.Supecification
End Property

Public Property Get PersonPerGroup() As Long
    If Not m_Initialized Then Err.Raise (91)
    PersonPerGroup = m_CompetitionEvent.PersonPerGroup
End Property

Public Property Get RoundsCount() As Long
    If Not m_Initialized Then Err.Raise (91)
    RoundsCount = m_CompetitionEvent.RoundsCount
End Property

Public Property Get OperationWindGauge() As Boolean
    If Not m_Initialized Then Err.Raise (91)
    OperationWindGauge = m_CompetitionEvent.OperationWindGauge
End Property

Public Property Get MeasurementReactionTime() As Boolean
    If Not m_Initialized Then Err.Raise (91)
    MeasurementReactionTime = m_CompetitionEvent.MeasurementReactionTime
End Property

Public Property Get ResultPrecision() As Long
    If Not m_Initialized Then Err.Raise (91)
    ResultPrecision = m_ResultPrecision
End Property

Public Property Get MinuteDelimiter() As String
    If Not m_Initialized Then Err.Raise (91)
    MinuteDelimiter = m_MinuteDelimiter
End Property

Public Property Get SecondDelimiter() As String
    If Not m_Initialized Then Err.Raise (91)
    SecondDelimiter = m_SecondDelimiter
End Property

Public Property Get SplitTimePrecision() As Long
    If Not m_Initialized Then Err.Raise (91)
    SplitTimePrecision = m_SplitTimePrecision
End Property

Public Property Get TrackRaceLength() As Long
    If Not m_Initialized Then Err.Raise (91)
    TrackRaceLength = m_TrackRaceLength
End Function

Public Property Get IntervalLength1() As Long
    If Not m_Initialized Then Err.Raise (91)
    IntervalLength1 = m_IntervalLength1
End Property

Public Property Get IntervalLength2() As Long
    If Not m_Initialized Then Err.Raise (91)
    IntervalLength2 = m_IntervalLength2
End Property

Public Property Get ScoringCombinedEvents() As Boolean
    If Not m_Initialized Then Err.Raise (91)
    ScoringCombinedEvents = m_ScoringCombinedEvents
End Property

Public Property Get ScoringIndexNumber1() As String
    If Not m_Initialized Then Err.Raise (91)
    ScoringIndexNumber1 = m_ScoringIndexNumber1
End Property

Public Property Get ScoringIndexNumber2() As String
    If Not m_Initialized Then Err.Raise (91)
    ScoringIndexNumber2 = m_ScoringIndexNumber2
End Property

Public Property Get ScoringIndexNumber3() As String
    If Not m_Initialized Then Err.Raise (91)
    ScoringIndexNumber3 = m_ScoringIndexNumber3
End Property

Public Property Get LaneSetting() As LaneSettingModel
    If Not m_Initialized Then Err.Raise (91)
    Set LaneSetting = m_LaneSetting
End Property

Public Property Get HtmlTemplate() As HtmlTemplatePathsModel
    If Not m_Initialized Then Err.Raise (91)
    Set HtmlTemplate = m_HtmlTemplate
End Property

Public Function SplitTimePoints() As Collection
    Dim checkPoints As Collection: Set checkPoints = New Collection
    Dim distance1 As Long: distance1 = Me.IntervalLength1
    Dim distance2 As Long: distance2 = Me.IntervalLength2

    ' ループを回す間隔の決定
    Dim distanceStep As Long
    If ((distance1 > 0) And (distance2 > 0)) Then
        ' ラップタイム間隔が両方あるため最小公倍数をとる。 
        distanceStep = Application.WorksheetFunction.Gcd(distance1, distance2)
    ElseIf (distance1 > 0) Then
        ' 片っぽだけ
        distanceStep = distance1
    ElseIf (distance2 > 0) Then
        ' 片っぽだけ
        distanceStep = distance2
    Else
        ' ラップタイム間隔が未指定なので、
        ' 空の配列を返す.
        Set SplitTimePoints = checkPoints
        Exit Function
    End If

    Dim i As Long
    For i = 0 To (m_TrackRaceLength - 1) Step distanceStep
        If (i > 0) Then
            Dim addToSplitTime As Boolean: addToSplitTime = False
            If (Not(addToSplitTime) And distance1 > 0) Then
                addToSplitTime = (i Mod distance1 = 0)
            End If

            If (Not(addToSplitTime) And distance2 > 0) Then
                addToSplitTime = (i Mod distance2 = 0)
            End If

            If (addToSplitTime) Then
                Call checkPoints.Add(i)
            End If
        End If
    Next i

    Set SplitTimePoints = checkPoints
End Function

Public Function ToJson() As String
    Dim jsonDict As Object: Set jsonDict = CreateObject("Scripting.Dictionary")

    Call jsonDict.Add("category", Me.Category)
    Call jsonDict.Add("sex", Me.Sex)
    Call jsonDict.Add("event_name", Me.EventName)
    Call jsonDict.Add("supecification", Me.Supecification)
    Call jsonDict.Add("person_per_group", Me.PersonPerGroup)
    Call jsonDict.Add("rounds_count", Me.RoundsCount)
    Call jsonDict.Add("operation_wind_gauge", Me.OperationWindGauge)
    Call jsonDict.Add("measurement_reaction_time", Me.MeasurementReactionTime)
    Call jsonDict.Add("minute_delimiter", Me.MinuteDelimiter)
    Call jsonDict.Add("second_delimiter", Me.SecondDelimiter)
    Call jsonDict.Add("split_time_precision", Me.SplitTimePrecision)
    Call jsonDict.Add("track_race_length", Me.TrackRaceLength)
    Call jsonDict.Add("interval_length_1", Me.IntervalLength1)
    Call jsonDict.Add("interval_length_2", Me.IntervalLength2)
    Call jsonDict.Add("scoring_combined_events", Me.ScoringCombinedEvents)
    Call jsonDict.Add("scoring_index_number_1", Me.ScoringIndexNumber1)
    Call jsonDict.Add("scoring_index_number_2", Me.ScoringIndexNumber2)
    Call jsonDict.Add("scoring_index_number_3", Me.ScoringIndexNumber3)
    Call jsonDict.Add("split_time_points", Me.SplitTimePoints)
    
    ToJson = JsonConverter.ConvertToJson(jsonDict)
End Function

Private Function IJsonConvertable_ToJson() As String
    IJsonConvertable_ToJson = Me.ToJson()
End Function
    