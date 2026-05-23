VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "AthletePersonModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba_files/common/model/person
Option Explicit
Implements IJsonConvertable
Implements IJsonParsable

Private m_Initialized As Boolean
Private m_Id As String
Private m_Sex As String
Private m_Bib As String
Private m_PersonalName As String
Private m_Age As String
Private m_PersonalPhonetic As String
Private m_PersonalLatin As String
Private m_TeamName As String
Private m_TeamPlace As String
Private m_PersonalCountry As String
Private m_PersonalGrade As String
Private m_PersonalCode As String
Private m_BirthDate As Date
Private m_TeamId As String

Private Sub Class_Initialize()
    m_Initialized = False
End Sub

Public Sub Initialize( _
        Sex As String _
        , Bib As String _
        , PersonalName As String _
        , Age As String _
        , TeamName As String _
        , PersonalPhonetic As String _
        , PersonalLatin As String _
        , TeamPlace As String _
        , PersonalCountry As String _
        , PersonalGrade As String _
        , BirthDate As Date _
        , PersonalCode As String _
        , TeamId As String _
        , Id As String _
    )

    If m_Initialized Then Err.Raise(91)
    m_Sex = Sex
    m_Bib = Bib
    m_PersonalName = PersonalName
    m_Age = Age
    m_TeamName = TeamName
    m_PersonalPhonetic = PersonalPhonetic
    m_PersonalLatin = PersonalLatin
    m_TeamPlace = TeamPlace
    m_PersonalCountry = PersonalCountry
    m_PersonalGrade = PersonalGrade
    m_PersonalCode = PersonalCode
    m_BirthDate = BirthDate
    m_TeamId = TeamId

    If (Id = "") Then 
        m_Id = UuidFactory.GenerateUuid()
    Else
        m_Id = Id
    End If

    m_Initialized = True
End Sub

Public Property Get Id() As String
    If (Not m_Initialized) Then Err.Raise(91)
    Id = m_Id
End Property

Public Property Get Sex() As String
    If (Not m_Initialized) Then Err.Raise(91)
    Sex = m_Sex
End Property

Public Property Get Bib() As String
    If (Not m_Initialized) Then Err.Raise(91)
    Bib = m_Bib
End Property

Public Property Get PersonalName() As String
    If (Not m_Initialized) Then Err.Raise(91)
    PersonalName = m_PersonalName
End Property

Public Property Get Age() As String
    If (Not m_Initialized) Then Err.Raise(91)
    Age = m_Age
End Property

Public Property Get PersonalPhonetic() As String
    If (Not m_Initialized) Then Err.Raise(91)
    PersonalPhonetic = m_PersonalPhonetic
End Property

Public Property Get PersonalLatin() As String
    If (Not m_Initialized) Then Err.Raise(91)
    PersonalLatin = m_PersonalLatin
End Property

Public Property Get TeamName() As String
    If (Not m_Initialized) Then Err.Raise(91)
    TeamName = m_TeamName
End Property

Public Property Get TeamPlace() As String
    If (Not m_Initialized) Then Err.Raise(91)
    TeamPlace = m_TeamPlace
End Property

Public Property Get PersonalCountry() As String
    If (Not m_Initialized) Then Err.Raise(91)
    PersonalCountry = m_PersonalCountry
End Property

Public Property Get PersonalGrade() As String
    If (Not m_Initialized) Then Err.Raise(91)
    PersonalGrade = m_PersonalGrade
End Property

Public Property Get BirthDate() As Date
    If (Not m_Initialized) Then Err.Raise(91)
    BirthDate = m_BirthDate
End Function

Public Property Get PersonalCode() As String
    If Not m_Initialized Then Err.Raise(91)
    PersonalCode = m_PersonalCode
End Property

Public Property Get TeamId() As String
    If Not m_Initialized Then Err.Raise(91)
    TeamId = m_TeamId
End Property

Public Function Self() As AthletePersonModel
    If Not m_Initialized Then Err.Raise(91)
    Set Self = Me
End Function

Public Function ToJson() As String
    Dim jsonDict As Object: Set jsonDict = CreateObject("Scripting.Dictionary")

    Call jsonDict.Add("id", Me.Id)
    Call jsonDict.Add("personal_code", Me.PersonalCode)
    Call jsonDict.Add("sex", Me.Sex)
    Call jsonDict.Add("bib", Me.Bib)
    Call jsonDict.Add("personal_name", Me.PersonalName)
    Call jsonDict.Add("age", Me.Age)
    Call jsonDict.Add("personal_phonetic", Me.PersonalPhonetic)
    Call jsonDict.Add("personal_latin", Me.PersonalLatin)
    Call jsonDict.Add("personal_country", Me.PersonalCountry)
    Call jsonDict.Add("personal_grade", Me.PersonalGrade)
    Call jsonDict.Add("birth_date", Me.BirthDate)
    Call jsonDict.Add("team_name", Me.TeamName)
    Call jsonDict.Add("team_place", Me.TeamPlace)
    Call jsonDict.Add("team_id", Me.TeamId)

    ToJson = JsonConverter.ConvertToJson(jsonDict)
End Function

Private Function IJsonConvertable_ToJson() As String
    IJsonConvertable_ToJson = Me.ToJson()
End Function

Public Sub FromJson(JsonString As String)
    Call IJsonParsable_FromJson(JsonString)
End Sub

Private Sub IJsonParsable_FromJson(JsonString As String)
    With JsonConverter.ParseJson(JsonString)
        Call Me.Initialize( _
                .Item("sex") _
                , .Item("bib") _
                , .Item("personal_name") _
                , .Item("age") _
                , .Item("team_name") _
                , .Item("personal_phonetic") _
                , .Item("personal_latin") _
                , .Item("team_place") _
                , .Item("personal_country") _
                , .Item("personal_grade") _
                , JsonConverter.ParseIso(.Item("birth_date")) _
                , .Item("personal_code") _
                , .Item("team_id") _
                , .Item("id") _
        )
    End With
End Sub
    