VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "TrackSubResultGroupModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/model/sub-result/track
Option Explicit
Implements IJsonConvertable

Private m_Id As String
Private m_RoundId As Long
Private m_Group As Long
Private m_Wind As WindValue
Private m_UpdatedAt As Date

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
End Sub

Private Sub Class_Terminate()
End Sub

Public Sub Initialize(RoundId As Long, Group As Long, Wind As WindValue, UpdatedAt As Date, Optional Id As String = "")
    If m_Initialized Then Err.Raise(91)
    m_RoundId = RoundId
    m_Group = Group
    Set m_Wind = Wind
    m_UpdatedAt = UpdatedAt

    If (Id = "") Then 
        m_Id = UuidFactory.GenerateUuid()
    Else
        m_Id = Id
    End If

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

Public Property Get Wind() As WindValue
    If Not m_Initialized Then Err.Raise(91)
    Set Wind = m_Wind
End Property

Public Property Get UpdatedAt() As Date
    If Not m_Initialized Then Err.Raise(91)
    UpdatedAt = m_UpdatedAt
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

    If Not (Me.Wind Is Nothing) Then
        Call jsonDictionary.Add("wind", Me.Wind.Value)
    End If

    Call jsonDictionary.Add("updated_at", Me.UpdatedAt)
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(jsonDictionary)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function
