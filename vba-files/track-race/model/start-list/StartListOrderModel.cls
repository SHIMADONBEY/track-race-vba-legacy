VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "StartListOrderModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/model/start-list
Option Explicit
Implements IJsonConvertable

Private Const NOT_ORDERED As Long = -2147483648
Private Const NOT_RANKED As Long = 2147483647

Private m_Id As String
Private m_Person As AthletePersonModel
Private m_Entry As PersonalEntryModel
Private m_QualifiedRank As Long
Private m_RoundId As Long
Private m_Group As Long
Private m_Order As Long
Private m_Accumaltive As Collection
Private m_NotStarted As Boolean

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
    ' TODO: Write Process To Initialize
    Set m_Accumaltive = New Collection
End Sub

Private Sub Class_Terminate()
    ' TODO: Write Process To Terminate
End Sub

Public Sub Initialize( _
            Person As AthletePersonModel _
            , Entry As PersonalEntryModel _
            , QualifiedRank As Long _
            , RoundId As Long _
            , AccumaltiveExpressions() As String _
            , NotStarted As Boolean _
            , Optional Group As Long = NOT_ORDERED _
            , Optional Order As Long = NOT_ORDERED _
            , Optional Id As String = "" _
)
    If m_Initialized Then Err.Raise(91)
    
    If (Id = "") Then 
        m_Id = UuidFactory.GenerateUuid()
    Else
        m_Id = Id
    End If

    Set m_Person = Person
    Set m_Entry = Entry

    If (QualifiedRank <= 0) Then
        m_QualifiedRank = NOT_RANKED
    Else
        m_QualifiedRank = QualifiedRank
    End If

    m_RoundId = RoundId
    If (Group > 0) Then
        m_Group = Group
    Else
        m_Group = NOT_ORDERED
    End If

    If (Order > 0) Then 
        m_Order = Order
    Else
        m_Order = NOT_ORDERED
    End If

    m_NotStarted = NotStarted

    Dim i As Long
    For i = LBound(AccumaltiveExpressions) To UBound(AccumaltiveExpressions)
        Call m_Accumaltive.Add(AccumaltiveExpressions(i))
    Next i

    m_Initialized = True
End Sub

Public Property Get Id() As String
    If Not m_Initialized Then Err.Raise(91)
    Id = m_Id
End Property

Public Property Get Person() As AthletePersonModel
    If Not m_Initialized Then Err.Raise(91)
    Set Person = m_Person
End Property

Public Property Get Entry() As PersonalEntryModel
    If Not m_Initialized Then Err.Raise(91)
    Set Entry = m_Entry
End Property

Public Property Get DemoEntry() As Boolean
    If Not m_Initialized Then Err.Raise(91)
    DemoEntry = m_Entry.DemoEntry
End Property

Public Property Get QualifiedRank() As Long
    If Not m_Initialized Then Err.Raise(91)
    QualifiedRank = m_QualifiedRank
End Property

Public Property Get Qualified1() As TimeResultValue
    If Not m_Initialized Then Err.Raise(91)
    Set Qualified1 = m_Entry.QualifiedValues(1)
End Property

Public Property Get Qualified2() As TimeResultValue
    If Not m_Initialized Then Err.Raise(91)
    Set Qualified2 = m_Entry.QualifiedValues(2)
End Property

Public Property Get Qualified3() As TimeResultValue
    If Not m_Initialized Then Err.Raise(91)
    Set Qualified3 = m_Entry.QualifiedValues(3)
End Property

Public Property Get RoundId() As Long
    If Not m_Initialized Then Err.Raise(91)
    RoundId = m_RoundId
End Property

Public Property Get Group() As Long
    If Not m_Initialized Then Err.Raise(91)
    Group = m_Group
End Property

Public Property Let Group(Value As Long)
    If Not m_Initialized Then Err.Raise(91)
    m_Group = Value
End Property

Public Property Get Order() As Long
    If Not m_Initialized Then Err.Raise(91)
    Order = m_Order
End Property

Public Property Let Order(Value As Long)
    If Not m_Initialized Then Err.Raise(91)
    m_Order = Value
End Property

Public Property Get Accumaltive() As Collection
    If Not m_Initialized Then Err.Raise(91)
    Set Accumaltive = m_Accumaltive
End Property

Public Property Get NotStarted() As Boolean
    If Not m_Initialized Then Err.Raise(91)
    NotStarted = m_NotStarted
End Property

Public Function Self() As StartListOrderModel
    If Not m_Initialized Then Err.Raise(91)
    Set Self = Me
End Function

Private Function IJsonConvertable_ToJson() As String
    Dim jsonDictionary As Object: Set jsonDictionary = CreateObject("Scripting.Dictionary")

    Call jsonDictionary.Add("id", Me.Id)
    Call jsonDictionary.Add("person", Me.Person)
    Call jsonDictionary.Add("round_id", Me.RoundId)
    Call jsonDictionary.Add("group", Me.Group)
    Call jsonDictionary.Add("order", Me.Order)
    Call jsonDictionary.Add("demo_entry", Me.DemoEntry)
    Call jsonDictionary.Add("accumaltive", Me.Accumaltive)
    Call jsonDictionary.Add("not_started", Me.NotStarted)
    Call jsonDictionary.Add("qualified_rank", Me.QualifiedRank)

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

Public Function ToJson() As String
   ToJson = IJsonConvertable_ToJson()    
End Function

Public Function IsOrdered() As Boolean
    IsOrdered = (Me.Group > 0 And Me.Order > 0)
End Function
