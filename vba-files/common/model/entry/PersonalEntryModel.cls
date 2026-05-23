VERSION 1.0 CLASS
BEGIN
MultiUse = -1 ' True
END
Attribute VB_Name = "PersonalEntryModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/common/model/entry
Option Explicit
Implements IJsonConvertable

Private Const QUALIFIED_VALUES_MAX_COUNT As Long = 3

Private m_PersonalId As String
Private m_EntryId As Long
Private m_EventId As Long
Private m_DemoEntry As Boolean
Private m_QualifiedValues As Collection

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
    Set m_QualifiedValues = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_QualifiedValues = Nothing
End Sub

Public Sub Initialize( _
        PersonalId As String _
        , EntryId As Long _
        , EventId As Long _
        , DemoEntry As Boolean _
        , ParamArray QualifiedValues() As Variant _
    )
    If m_Initialized Then Err.Raise(91)
    m_PersonalId = PersonalId
    m_EntryId = EntryId
    m_EventId = EventId
    m_DemoEntry = DemoEntry

    Dim i As Long
    Dim vStartedIndex As Long: vStartedIndex = LBound(QualifiedValues)
    For i = vStartedIndex To UBound(QualifiedValues)
        If (i - vStartedIndex >= QUALIFIED_VALUES_MAX_COUNT) Then
            Exit For
        End If

        Dim vQualifiedValue As IResultValue: Set vQualifiedValue = QualifiedValues(i)
        Call m_QualifiedValues.Add(vQualifiedValue)
    Next i

    m_Initialized = True
End Sub

Public Property Get PersonalId() As String
    If Not m_Initialized Then Err.Raise(91)
    PersonalId = m_PersonalId
End Property

Public Property Get EntryId() As Long
    If Not m_Initialized Then Err.Raise(91)
    EntryId = m_EntryId
End Property

Public Property Get EventId() As Long
    If Not m_Initialized Then Err.Raise(91)
    EventId = m_EventId
End Property

Public Property Get DemoEntry() As Boolean
    If Not m_Initialized Then Err.Raise(91)
    DemoEntry = m_DemoEntry
End Property

Public Property Get QualifiedValues() As Collection
    If Not m_Initialized Then Err.Raise(91)
    Set QualifiedValues = m_QualifiedValues
End Property

Public Function Self() 
    If Not m_Initialized Then Err.Raise(91)
    Set Self = Me
End Function

Private Function IJsonConvertable_ToJson() As String
    Dim jsonDictionary As Object: Set jsonDictionary = CreateObject("Scripting.Dictionary")
    Call jsonDictionary.Add("personal_id", Me.PersonalId)
    Call jsonDictionary.Add("entry_id", Me.EntryId)
    Call jsonDictionary.Add("event_id", Me.EventId)
    Call jsonDictionary.Add("demo_entry", Me.DemoEntry)

    Dim vQualifiedStringValues As Collection: Set vQualifiedStringValues = New Collection
    Dim vQualifiedValue As IResultValue
    For Each vQualifiedValue In Me.QualifiedValues
        Dim vQualifiedStringValue As String: vQualifiedStringValue = ""
        If Not (vQualifiedValue Is Nothing) Then
            vQualifiedStringValue = vQualifiedValue.ToString()
        End If

        Call vQualifiedStringValues.Add(vQualifiedStringValue)
    Next vQualifiedValue 
    
    Call jsonDictionary.Add("qualified_values", vQualifiedStringValues)
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(jsonDictionary)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function
