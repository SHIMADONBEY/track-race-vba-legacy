VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "CompetitionEventModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/common/model/event
Option Explicit
Implements IJsonConvertable
Implements IJsonParsable

Private m_Id As Long
Private m_Category As String
Private m_Sex As String
Private m_EventName As String
Private m_Supecification As String
Private m_EventType As String
Private m_RoundsCount As Long
Private m_PersonPerGroup As Long
Private m_OperationWindGauge As Boolean
Private m_MeasurementReactionTime As Boolean
Private m_MinutesPerGroup As Variant
Private m_CallStartTiming As Variant
Private m_CallEndTiming As Variant
Private m_TemplateFileName As String
Private m_OutputFileName As String

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
    ' TODO: Write Process To Initialize
End Sub

Private Sub Class_Terminate()
    ' TODO: Write Process To Terminate
End Sub

Public Sub Initialize( _
        Id As Long _
        , Category As String _
        , Sex As String _
        , EventName As String _
        , Supecification As String _
        , EventType As String _
        , RoundsCount As Long _
        , PersonPerGroup As Long _
        , OperationWindGauge As Boolean _
        , MeasurementReactionTime As Boolean _
        , MinutesPerGroup As Variant _
        , CallStartTiming As Variant _
        , CallEndTiming As Variant _
        , TemplateFileName As String _
        , OutputFileName As String)

    If m_Initialized Then Err.Raise(91)
    m_Id = Id
    m_Category = Category
    m_Sex = Sex
    m_EventName = EventName
    m_Supecification = Supecification
    m_EventType = EventType
    m_RoundsCount = RoundsCount
    m_PersonPerGroup = PersonPerGroup
    m_OperationWindGauge = OperationWindGauge
    m_MeasurementReactionTime = MeasurementReactionTime
    m_MinutesPerGroup = CDec(MinutesPerGroup)
    m_CallStartTiming = CDec(CallStartTiming)
    m_CallEndTiming = CDec(CallEndTiming)
    m_TemplateFileName = TemplateFileName
    m_OutputFileName = OutputFileName
    m_Initialized = True
End Sub

Public Property Get Id() As Long
    If Not m_Initialized Then Err.Raise(91)
    Id = m_Id
End Property

Public Property Get Category() As String
    If Not m_Initialized Then Err.Raise(91)
    Category = m_Category
End Property

Public Property Get Sex() As String
    If Not m_Initialized Then Err.Raise(91)
    Sex = m_Sex
End Property

Public Property Get EventName() As String
    If Not m_Initialized Then Err.Raise(91)
    EventName = m_EventName
End Property

Public Property Get Supecification() As String
    If Not m_Initialized Then Err.Raise(91)
    Supecification = m_Supecification
End Property

Public Property Get EventType() As String
    If Not m_Initialized Then Err.Raise(91)
    EventType = m_EventType
End Property

Public Property Get RoundsCount() As Long
    If Not m_Initialized Then Err.Raise(91)
    RoundsCount = m_RoundsCount
End Property

Public Property Get PersonPerGroup() As Long
    If Not m_Initialized Then Err.Raise(91)
    PersonPerGroup = m_PersonPerGroup
End Property

Public Property Get OperationWindGauge() As Boolean
    If Not m_Initialized Then Err.Raise(91)
    OperationWindGauge = m_OperationWindGauge
End Property

Public Property Get MeasurementReactionTime() As Boolean
    If Not m_Initialized Then Err.Raise(91)
    MeasurementReactionTime = m_MeasurementReactionTime
End Property

Public Property Get MinutesPerGroup() As Variant
    If Not m_Initialized Then Err.Raise(91)
    MinutesPerGroup = m_MinutesPerGroup
End Property

Public Property Get CallStartTiming() As Variant
    If Not m_Initialized Then Err.Raise(91)
    CallStartTiming = m_CallStartTiming
End Property

Public Property Get CallEndTiming() As Variant
    If Not m_Initialized Then Err.Raise(91)
    CallEndTiming = m_CallEndTiming
End Property
    
Public Property Get TemplateFileName() As String
    If Not m_Initialized Then Err.Raise(91)
    TemplateFileName = m_TemplateFileName
End Property

Public Property Get OutputFileName() As String
    If Not m_Initialized Then Err.Raise(91)
    OutputFileName = m_OutputFileName
End Property

Public Function Self() 
    If Not m_Initialized Then Err.Raise(91)
    Set Self = Me
End Function

Public Function EventKey() As String
    EventKey = Me.Category & Me.Sex & Me.EventName
    If (Me.Supecification = "") Then
        Exit Function
    End If

    EventKey = EventKey & " (" & Me.Supecification & ")"
End Function

Private Function IJsonConvertable_ToJson() As String
    Dim jsonDictionary As Object: Set jsonDictionary = CreateObject("Scripting.Dictionary")
    Call jsonDictionary.Add("id", Me.Id)
    Call jsonDictionary.Add("category", Me.Category)
    Call jsonDictionary.Add("sex", Me.Sex)
    Call jsonDictionary.Add("event_name", Me.EventName)
    Call jsonDictionary.Add("supecification", Me.Supecification)
    Call jsonDictionary.Add("event_type", Me.EventType)
    Call jsonDictionary.Add("rounds_count", Me.RoundsCount)
    Call jsonDictionary.Add("person_per_group", Me.PersonPerGroup)
    Call jsonDictionary.Add("operation_wind_gauge", Me.OperationWindGauge)
    Call jsonDictionary.Add("measurement_reaction_time", Me.MeasurementReactionTime)
    Call jsonDictionary.Add("minutes_per_group", Me.MinutesPerGroup)
    Call jsonDictionary.Add("call_start_timing", Me.CallStartTiming)
    Call jsonDictionary.Add("call_end_timing", Me.CallEndTiming)
    Call jsonDictionary.Add("template_file_name", Me.TemplateFileName)
    Call jsonDictionary.Add("output_file_name", Me.OutputFileName)
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(jsonDictionary)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function

Public Sub FromJson(JsonString As String)
    Call IJsonParsable_FromJson(JsonString)
End Sub

Private Sub IJsonParsable_FromJson(JsonString As String)
    With JsonConverter.ParseJson(JsonString)
        Call Me.Initialize( _
                .Item("id") _
                , .Item("category") _
                , .Item("sex") _
                , .Item("event_name") _
                , .Item("supecification") _
                , .Item("event_type") _
                , .Item("rounds_count") _
                , .Item("person_per_group") _
                , .Item("operation_wind_gauge") _
                , .Item("measurement_reaction_time") _
                , .Item("minutes_per_group") _
                , .Item("call_start_timing") _
                , .Item("call_end_timing") _
                , .Item("template_file_name") _
                , .Item("output_file_name") _
        )
    End With
End Sub
