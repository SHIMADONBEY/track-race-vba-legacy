VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "TrackResultRoundModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/model/result/track
Option Explicit
Implements IJsonConvertable

Private m_Groups As TrackResultGroupModels
Private m_RoundId As Long

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
    Set m_Groups = New TrackResultGroupModels
End Sub

Private Sub Class_Terminate()
    Set m_Groups = Nothing
End Sub

Public Sub Initialize(GroupCount As Long, RoundId As Long, InputGroupList As TrackSubResultGroupModels)
    If m_Initialized Then Err.Raise(91)
    m_RoundId = RoundId

    Dim i As Long
    For i = 1 To GroupCount
        With New TrackResultGroupModel
            Dim inputGroup As TrackSubResultGroupModel: Set inputGroup = InputGroupList.FindByGroup(RoundId, i)
            If (inputGroup Is Nothing) Then 
                Call .Initialize(RoundId, i, Nothing, Empty)
            Else
                Call .Initialize(RoundId, i, inputGroup.Wind, inputGroup.UpdatedAt, inputGroup.Id)
            End If

            Call m_Groups.Add(.Self())
        End With
    Next i

    m_Initialized = True
End Sub

Public Property Get Groups() As TrackResultGroupModels
    If Not m_Initialized Then Err.Raise(91)
    Set Groups = New TrackResultGroupModels
    Call Groups.AddRange(m_Groups.All())
End Property

Public Function Group(GroupNumber As Long) As TrackResultGroupModel
    Set Group = m_Groups.FindByRoundGroup(Me.RoundId, GroupNumber)
End Function

Public Property Get RoundId() As Long
    If Not m_Initialized Then Err.Raise(91)
    RoundId = m_RoundId
End Property

Public Sub AddResults(Results As TrackResultOrderModels)
    Dim vClassifier As TrackResultGroupClassifier: Set vClassifier = New TrackResultGroupClassifier
    Dim groupOrders As Object: Set groupOrders = CollectionUtil.GroupingBy(Results.All(), vClassifier)

    Dim groupKey As Variant
    For Each groupKey In groupOrders
        Call m_Groups.AddResults(vClassifier.RoundId(groupKey), vClassifier.Group(groupKey), groupOrders.Item(groupKey))
    Next groupKey 
End Sub

Public Sub AddSplitTimes(SplitTimes As TrackSubResultOrderModels)
    Dim vClassifier As TrackSubResultGroupClassifier: Set vClassifier = New TrackSubResultGroupClassifier
    Dim groupedSplitTimes As Object: Set groupedSplitTimes = CollectionUtil.GroupingBy(SplitTimes.All(), vClassifier)

    Dim groupKey As Variant
    For Each groupKey In groupedSplitTimes
        Call m_Groups.AddSplitTimes(vClassifier.RoundId(groupKey), vClassifier.Group(groupKey), groupedSplitTimes.Item(groupKey))
    Next groupKey 
End Sub

Public Function Self() 
    If Not m_Initialized Then Err.Raise(91)
    Set Self = Me
End Function

Private Function IJsonConvertable_ToJson() As String
    Dim jsonDictionary As Object: Set jsonDictionary = CreateObject("Scripting.Dictionary")
    Call jsonDictionary.Add("items", Me.Groups)
    Call jsonDictionary.Add("round_id", Me.RoundId)
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(jsonDictionary)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function
