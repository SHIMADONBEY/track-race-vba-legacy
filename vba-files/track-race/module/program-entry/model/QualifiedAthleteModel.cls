VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QualifiedAthleteModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/modules/program-entry/model
Option Explicit
Implements IJsonConvertable

Private Const NOT_ASSIGNED As Long = -2147483648

Private m_Initialized As Boolean
Private m_Id As String
Private m_Person As AthletePersonModel
Private m_Entry As PersonalEntryModel
Private m_QualifiedRank As Long
Private m_Group As Long
Private m_Order As Long
Private m_Weight As Single

Private Sub Class_Initialize()
    m_Initialized = False
End Sub

Public Sub Initialize( _
        Id As String _
        , Person As AthletePersonModel _
        , Entry As PersonalEntryModel _
        , QualifiedRank As Long _
    )

    If (Person Is Nothing) Then Err.Raise 327, "QualifiedAthleteModel.Initialize"
    If (Entry Is Nothing) Then Err.Raise 327, "QualifiedAthleteModel.Initialize"

    m_Id = Id
    Set m_Person = Person
    Set m_Entry = Entry
    m_QualifiedRank = QualifiedRank
    m_Group = NOT_ASSIGNED
    m_Order = NOT_ASSIGNED
    m_Weight = Rnd()
    m_Initialized = True
End Sub

Public Property Get Id() As String
    If Not m_Initialized Then Err.Raise (91)
    Id = m_Id
End Property

Public Property Get Person() As AthletePersonModel
    If Not m_Initialized Then Err.Raise (91)
    Set Person = m_Person
End Property

Public Property Get Entry() As PersonalEntryModel
    If Not m_Initialized Then Err.Raise (91)
    Set Entry = m_Entry
End Property

Public Property Get QualifiedRank() As Long
    If Not m_Initialized Then Err.Raise (91)
    QualifiedRank = m_QualifiedRank
End Property

Public Property Get Qualified1() As TimeResultValue
    If Not m_Initialized Then Err.Raise (91)
    Set Qualified1 = m_Entry.QualifiedValues(1)
End Property

Public Property Get Qualified2() As TimeResultValue
    If Not m_Initialized Then Err.Raise (91)
    Set Qualified2 = m_Entry.QualifiedValues(2)
End Property

Public Property Get Qualified3() As TimeResultValue
    If Not m_Initialized Then Err.Raise (91)
    Set Qualified3 = m_Entry.QualifiedValues(3)
End Property

Public Property Get DemoEntry() As Boolean
    If Not m_Initialized Then Err.Raise (91)
    DemoEntry = m_Entry.DemoEntry
End Property

Public Property Let Group(ByVal Value As Long)
    If Not m_Initialized Then Err.Raise (91)
    m_Group = Value
End Property

Public Property Get Group() As Long
    If Not m_Initialized Then Err.Raise (91)
    Group = m_Group
End Property

Public Property Let Order(ByVal Value As Long)
    If Not m_Initialized Then Err.Raise (91)
    m_Order = Value
End Property

Public Property Get Order() As Long
    If Not m_Initialized Then Err.Raise (91)
    Order = m_Order
End Property

Public Sub UpdateWeight()
    If Not m_Initialized Then Err.Raise (91)
    m_Weight = Rnd()
End Sub

Public Property Get Weight() As Single
    If Not m_Initialized Then Err.Raise (91)
    Weight = m_Weight
End Property

Public Sub AddWeight(Addend As Single)
    If Not m_Initialized Then Err.Raise (91)
    m_Weight = m_Weight + Addend
End Sub

Public Function Self() As QualifiedAthleteModel
    If Not m_Initialized Then Err.Raise (91)
    Set Self = Me
End Function

Private  Function IJsonConvertable_ToJson() As String
    Dim jsonDictionary As Object: Set jsonDictionary = CreateObject("Scripting.Dictionary")

    Call jsonDictionary.Add("id", Me.Id)
    Call jsonDictionary.Add("person", Me.Person)
    Call jsonDictionary.Add("qualified_rank", Me.QualifiedRank)
    Call jsonDictionary.Add("demo_entry", Me.DemoEntry)

    Dim vGroup As Long: vGroup = Me.Group
    If Not (vGroup = NOT_ASSIGNED) Then
        Call jsonDictionary.Add("group", vGroup)
    End If

    Dim vOrder As Long: vOrder = Me.Order
    If Not (vOrder = NOT_ASSIGNED) Then
        Call jsonDictionary.Add("order", vOrder)
    End If

    If (Me.Qualified1 Is Nothing) Then 
        Call jsonDictionary.Add("qualified_time_1", "")    
    Else
        Call jsonDictionary.Add("qualified_time_1", Me.Qualified1.ToString())
    End If

    If (Me.Qualified2 Is Nothing) Then 
        Call jsonDictionary.Add("qualified_time_2", "")    
    Else
        Call jsonDictionary.Add("qualified_time_2", Me.Qualified2.ToString())
    End If

    If (Me.Qualified3 Is Nothing) Then 
        Call jsonDictionary.Add("qualified_time_3", "")    
    Else
        Call jsonDictionary.Add("qualified_time_3", Me.Qualified3.ToString())
    End If

    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(jsonDictionary)
End Function