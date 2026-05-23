VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "ThreeLaneAssigner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/modules/program-entry/policy/order
Option Explicit
Implements IOrderAssignable

Private m_AssignLanes() As Long
Private m_Initialized As Boolean
Private m_TopLaneCount As Long
Private m_MiddleLaneCount As Long
Private m_BottomLaneCount As Long
Private m_PersonsPerGroup As Long
Private m_OrderSetting As LaneSettingModel
Private m_GroupIndexes(0 To 4) As Long

Private Sub Class_Initialize()
    Dim orderSetting As LaneSettingModel: Set orderSetting = CompetitionRepository.ReadSettinng().laneSetting
    Set m_OrderSetting = orderSetting

    m_TopLaneCount = UBound(orderSetting.TopLanes) - LBound(orderSetting.TopLanes) + 1
    m_MiddleLaneCount = UBound(orderSetting.MiddleLanes) - LBound(orderSetting.MiddleLanes) + 1
    m_BottomLaneCount = UBound(orderSetting.BottomLanes) - LBound(orderSetting.BottomLanes) + 1


    m_PersonsPerGroup = m_TopLaneCount + m_MiddleLaneCount + m_BottomLaneCount _
            + (UBound(orderSetting.ExtraLanes) - LBound(orderSetting.ExtraLanes) + 1)

    Dim i As Long, j As Long
    Dim lowerBound As Long
    ReDim m_AssignLanes(0 To m_PersonsPerGroup - 1)
    Dim laneList As Variant

    j = 0
    m_GroupIndexes(0) = j + 1
    laneList = orderSetting.TopLanes
    For i = LBound(laneList) To UBound(laneList)
        m_AssignLanes(j) = CLng(laneList(i))
        j = j + 1
    Next i

    m_GroupIndexes(1) = j + 1
    laneList = orderSetting.MiddleLanes
    For i = LBound(laneList) To UBound(laneList)
        m_AssignLanes(j) = CLng(laneList(i))
        j = j + 1
    Next i

    m_GroupIndexes(2) = j + 1
    laneList = orderSetting.BottomLanes
    For i = LBound(laneList) To UBound(laneList)
        m_AssignLanes(j) = CLng(laneList(i))
        j = j + 1
    Next i

    m_GroupIndexes(3) = j + 1
    laneList = orderSetting.ExtraLanes
    For i = LBound(laneList) To UBound(laneList)
        m_AssignLanes(j) = CLng(laneList(i))
        j = j + 1
    Next i

    m_GroupIndexes(4) = j + 1

    m_Initialized = True
End Sub

Private Function IOrderAssignable_AssignOrder(ByRef Qualifiers As QualifiedAthleteModels) As QualifiedAthleteModels
    Set IOrderAssignable_AssignOrder = AssignOrder(Qualifiers)
End Function

Private Property Get IOrderAssignable_PersonsPerGroup() As Long
    IOrderAssignable_PersonsPerGroup = PersonsPerGroup
End Property

Public Function AssignOrder(ByRef Qualifiers As QualifiedAthleteModels) As QualifiedAthleteModels
    Dim i As Long
    Dim currentLevel As Long
    Dim row As QualifiedAthleteModel
    Dim rows As Collection

    currentLevel = 0
    Set rows = Qualifiers.Sort().All()
    For i = 1 To rows.Count
        If i >= m_GroupIndexes(currentLevel + 1) Then
            currentLevel = currentLevel + 1
        End If

        Set row = rows.Item(i)
        Call row.AddWeight(CSng(currentLevel))
    Next i

    Dim weightSortedRows As Collection: Set weightSortedRows = CollectionUtil.SortObjectCollection(rows, New QualifiedWeightComparator)

    Set AssignOrder = New QualifiedAthleteModels
    For i = 1 To weightSortedRows.Count
        Set row = weightSortedRows.Item(i)
        row.Order = m_AssignLanes(i - 1)
        Call AssignOrder.Add(row)
    Next i
End Function

Public Property Get PersonsPerGroup() As Long
  PersonsPerGroup = m_PersonPerGroup
End Property
