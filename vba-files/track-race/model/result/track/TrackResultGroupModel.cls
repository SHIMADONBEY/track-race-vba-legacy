VERSION 1.0 CLASS
  BEGIN
  MultiUse = -1  'True
  END
  Attribute VB_Name = "TrackResultGroupModel"
  Attribute VB_GlobalNameSpace = False
  Attribute VB_Creatable = False
  Attribute VB_PredeclaredId = False
  Attribute VB_Exposed = False
'namespace=vba-files/track-race/model/result/track
Option Explicit
Implements IJsonConvertable

Private m_Id As String
Private m_RoundId As Long
Private m_GroupNumber As Long
Private m_Wind As WindValue
Private m_UpdatedAt As Date
Private m_Results As TrackResultOrderModels
Private m_SplitTimes As TrackSubResultOrderModels

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
    Set m_Results = New TrackResultOrderModels
    Set m_SplitTimes = New TrackSubResultOrderModels
End Sub

Private Sub Class_Terminate()
    ' TODO: Write Process To Terminate
    Set m_Results = Nothing
    Set m_SplitTimes = Nothing
End Sub

Public Sub Initialize( _
        RoundId As Long _
        , GroupNumber As Long _
        , Wind As WindValue _
        , UpdatedAt As Date _
        , Optional Id As String = "" _
        , Optional Results As TrackResultOrderModels = Nothing _
        , Optional SplitTimes As TrackSubResultOrderModels = Nothing _
)
    If m_Initialized Then Err.Raise(91)
    m_RoundId = RoundId
    m_GroupNumber = GroupNumber
    Set m_Wind = Wind
    m_UpdatedAt = UpdatedAt

    If (Id = "") Then 
        m_Id = UuidFactory.GenerateUuid()
    Else
        m_Id = Id
    End If

    If (Not Results Is Nothing) Then 
        Call m_Results.AddRange(Results.All())
    End If

    If (Not SplitTimes Is Nothing) Then
        Call m_SplitTimes.AddRange(SplitTimes.All())
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

Public Property Get GroupNumber() As Long
    If Not m_Initialized Then Err.Raise(91)
    GroupNumber = m_GroupNumber
End Property

Public Property Get Wind() As WindValue
    If Not m_Initialized Then Err.Raise(91)
    Set Wind = m_Wind
End Property

Public Property Get UpdatedAt() As Date
    If Not m_Initialized Then Err.Raise(91)
    UpdatedAt = m_UpdatedAt
End Property

Public Property Get Results() As TrackResultOrderModels
    If Not m_Initialized Then Err.Raise(91)

    With New TrackResultOrderModels
        Call .AddRange(m_Results.All())
        Set Results = .Self()
    End With
End Property

Public Property Get SplitTimes() As TrackSubResultOrderModels
    If Not m_Initialized Then Err.Raise(91)
    Set SplitTimes = New TrackSubResultOrderModels
    Call SplitTimes.AddRange(CollectionUtil.SortObjectCollection(m_SplitTimes.All(), New TrackSubResultNumberComparator))
End Property

Public  Sub AddResults(Items As Collection)
    If Not m_Initialized Then Err.Raise(91)
    Call m_Results.AddRange(Items)
End Sub

Public Sub AddSplitTimes(Items As Collection)
    If Not m_Initialized Then Err.Raise(91)
    Call m_SplitTimes.AddRange(Items)    
End Sub

Public Function Self() 
    If Not m_Initialized Then Err.Raise(91)
    Set Self = Me
End Function

Private Function IJsonConvertable_ToJson() As String
    Dim jsonDictionary As Object: Set jsonDictionary = CreateObject("Scripting.Dictionary")
    Call jsonDictionary.Add("id", Me.Id)
    Call jsonDictionary.Add("round_id", Me.RoundId)
    Call jsonDictionary.Add("group_number", Me.GroupNumber)

    If Not (Me.Wind Is Nothing) Then
        Call jsonDictionary.Add("wind", Me.Wind.Value)
    End If

    Call jsonDictionary.Add("updated_at", Me.UpdatedAt)
    Call jsonDictionary.Add("results", Me.Results)
    Call jsonDictionary.Add("split_times", Me.SplitTimes)
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(jsonDictionary)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function
