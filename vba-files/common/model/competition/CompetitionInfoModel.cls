VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "CompetitionInfoModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/common/model/competition
Option Explicit
Implements IJsonConvertable
Implements IJsonParsable

Private m_CompetitionName As String
Private m_CompetitionLatin As String
Private m_CompetitionCode As String
Private m_FacilityName As String
Private m_FacilityLatin As String
Private m_FacilityCode As String
Private m_FacilityPlace As String
Private m_CompetitionYear As Long
Private m_CompetitionDateStart As Date
Private m_CompetitionDateEnd As Date

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
    ' TODO: Write Process To Initialize
End Sub

Private Sub Class_Terminate()
    ' TODO: Write Process To Terminate
End Sub

Public Sub Initialize( _
            CompetitionName As String _
            , CompetitionLatin As String _
            , CompetitionCode As String _
            , FacilityName As String _
            , FacilityLatin As String _
            , FacilityCode As String _
            , FacilityPlace As String _
            , CompetitionYear As Long _
            , CompetitionDateStart As Date _
            , CompetitionDateEnd As Date)
    If m_Initialized Then Err.Raise(91)
    m_CompetitionName = CompetitionName
    m_CompetitionLatin = CompetitionLatin
    m_CompetitionCode = CompetitionCode
    m_FacilityName = FacilityName
    m_FacilityLatin = FacilityLatin
    m_FacilityCode = FacilityCode
    m_FacilityPlace = FacilityPlace
    m_CompetitionYear = CompetitionYear
    m_CompetitionDateStart = CompetitionDateStart
    m_CompetitionDateEnd = CompetitionDateEnd
    m_Initialized = True
End Sub

Public Property Get CompetitionName() As String
    If Not m_Initialized Then Err.Raise(91)
    CompetitionName = m_CompetitionName
End Property

Public Property Get CompetitionLatin() As String
    If Not m_Initialized Then Err.Raise(91)
    CompetitionLatin = m_CompetitionLatin
End Property

Public Property Get CompetitionCode() As String
    If Not m_Initialized Then Err.Raise(91)
    CompetitionCode = m_CompetitionCode
End Property

Public Property Get FacilityName() As String
    If Not m_Initialized Then Err.Raise(91)
    FacilityName = m_FacilityName
End Property

Public Property Get FacilityLatin() As String
    If Not m_Initialized Then Err.Raise(91)
    FacilityLatin = m_FacilityLatin
End Property

Public Property Get FacilityCode() As String
    If Not m_Initialized Then Err.Raise(91)
    FacilityCode = m_FacilityCode
End Property

Public Property Get FacilityPlace() As String
    If Not m_Initialized Then Err.Raise(91)
    FacilityPlace = m_FacilityPlace
End Property

Public Property Get CompetitionYear() As Long
    If Not m_Initialized Then Err.Raise(91)
    CompetitionYear = m_CompetitionYear
End Property

Public Property Get CompetitionDateStart() As Date
    If Not m_Initialized Then Err.Raise(91)
    CompetitionDateStart = m_CompetitionDateStart
End Property

Public Property Get CompetitionDateEnd() As Date
    If Not m_Initialized Then Err.Raise(91)
    CompetitionDateEnd = m_CompetitionDateEnd
End Property

Public Function Self() 
    If Not m_Initialized Then Err.Raise(91)
    Set Self = Me
End Function

Private Function IJsonConvertable_ToJson() As String
    Dim jsonDictionary As Object: Set jsonDictionary = CreateObject("Scripting.Dictionary")
    Call jsonDictionary.Add("competition_name", Me.CompetitionName)
    Call jsonDictionary.Add("competition_latin", Me.CompetitionLatin)
    Call jsonDictionary.Add("competition_code", Me.CompetitionCode)
    Call jsonDictionary.Add("facility_name", Me.FacilityName)
    Call jsonDictionary.Add("facility_latin", Me.FacilityLatin)
    Call jsonDictionary.Add("facility_code", Me.FacilityCode)
    Call jsonDictionary.Add("facility_place", Me.FacilityPlace)
    Call jsonDictionary.Add("competition_year", Me.CompetitionYear)
    Call jsonDictionary.Add("competition_date_start", Me.CompetitionDateStart)
    Call jsonDictionary.Add("competition_date_end", Me.CompetitionDateEnd)
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
                .Item("competition_name") _
                , .Item("competition_latin") _
                , .Item("competition_code") _
                , .Item("facility_name") _
                , .Item("facility_latin") _
                , .Item("facility_code") _
                , .Item("facility_place") _
                , CLng(.Item("competition_year")) _
                , JsonConverter.ParseIso(.Item("competition_date_start")) _
                , JsonConverter.ParseIso(.Item("competition_date_end")) _
        )
    End With
End Sub
