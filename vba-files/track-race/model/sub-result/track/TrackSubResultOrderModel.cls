VERSION 1.0 CLASS
  BEGIN
  MultiUse = -1  'True
  END
  Attribute VB_Name = "TrackSubResultOrderModel"
  Attribute VB_GlobalNameSpace = False
  Attribute VB_Creatable = False
  Attribute VB_PredeclaredId = False
  Attribute VB_Exposed = False
'namespace=vba-files/track-race/model/sub-result/track
Option Explicit
Implements IJsonConvertable

Private m_Id As String
Private m_RoundId As Long
Private m_GroupNumber As Long
Private m_Order As Long
Private m_SubNumber As Long
Private m_Comment As String
Private m_UpdatedAt As Date
Private m_ResultId As String
Private m_SubResult As TimeResultValue

' îCà”Ç≈ñÑÇﬂçûÇ›
Private m_Person As AthletePersonModel

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
    ' TODO: Write Process To Initialize
End Sub

Private Sub Class_Terminate()
    ' TODO: Write Process To Terminate
End Sub

Public Sub Initialize( _
        Id As String _
        , RoundId As Long _
        , GroupNumber As Long _
        , Order As Long _
        , SubNumber As Long _
        , Comment As String _
        , SubResult As TimeResultValue _
        , ResultId As String _
        , UpdatedAt As Date _
)
    If m_Initialized Then Err.Raise(91)
    m_Id = Id
    m_RoundId = RoundId
    m_GroupNumber = GroupNumber
    m_Order = Order
    m_SubNumber = SubNumber
    m_Comment = Comment
    m_UpdatedAt = UpdatedAt
    m_ResultId = ResultId
    Set m_SubResult = SubResult
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

Public Property Get GroupNumber() As Long
    If Not m_Initialized Then Err.Raise(91)
    GroupNumber = m_GroupNumber
End Property

Public Property Get Order() As Long
    If Not m_Initialized Then Err.Raise(91)
    Order = m_Order
End Property

Public Property Get SubNumber() As Long
    If Not m_Initialized Then Err.Raise(91)
    SubNumber = m_SubNumber
End Property

Public Property Get Comment() As String
    If Not m_Initialized Then Err.Raise(91)
    Comment = m_Comment
End Property

Public Property Get UpdatedAt() As Date
    If Not m_Initialized Then Err.Raise(91)
    UpdatedAt = m_UpdatedAt
End Property

Public Property Get ResultId() As String
    If Not m_Initialized Then Err.Raise(91)
    ResultId = m_ResultId
End Property
Public Property Let ResultId(Value As String)
    If Not m_Initialized Then Err.Raise(91)
    m_ResultId = Value
End Property
Public Property Get SubResult() As TimeResultValue
    If Not m_Initialized Then Err.Raise(91)
    Set SubResult = m_SubResult
End Property
Public Property Set SubResult(Value As TimeResultValue)
    If Not m_Initialized Then Err.Raise(91)
    Set m_SubResult = Value
End Property

Public Property Get Person() As AthletePersonModel
    If Not m_Initialized Then Err.Raise(91)
    Set Person = m_Person
End Property

Public Property Set Person(Value As AthletePersonModel)
    If Not m_Initialized Then Err.Raise(91)
    Set m_Person = Value
End Property
    
Public Function Self() 
    If Not m_Initialized Then Err.Raise(91)
    Set Self = Me
End Function

Private Function IJsonConvertable_ToJson() As String
    Dim jsonDictionary As Object: Set jsonDictionary = CreateObject("Scripting.Dictionary")
    Call jsonDictionary.Add("id", Me.Id)
    Call jsonDictionary.Add("round_id", Me.RoundId)
    Call jsonDictionary.Add("group_number", Me.GroupNumber)
    Call jsonDictionary.Add("order", Me.Order)
    Call jsonDictionary.Add("sub_number", Me.SubNumber)
    Call jsonDictionary.Add("comment", Me.Comment)
    Call jsonDictionary.Add("updated_at", Me.UpdatedAt)
    Call jsonDictionary.Add("result_id", Me.ResultId)

    If (Me.SubResult Is Nothing) Then
        Call jsonDictionary.Add("sub_result", "")
    Else
        Call jsonDictionary.Add("sub_result", Me.SubResult.ToString())
    End If

    If Not (Me.Person Is Nothing) Then
        Call jsonDictionary.Add("person", Me.Person)    
    End If

    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(jsonDictionary)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function
