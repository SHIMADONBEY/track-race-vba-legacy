VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "AscendOrderAssigner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/modules/program-entry/policy/order
Option Explicit
Implements IOrderAssignable

Private m_AssignLanes() As Long
Private m_Initialized As Boolean
Private m_PersonsPerGroup As Long

Private Sub Class_Initialize()
    m_PersonsPerGroup = CompetitionRepository.ReadSettinng().PersonPerGroup
    m_Initialized = True
End Sub

Private Function IOrderAssignable_AssignOrder(ByRef Qualifiers As QualifiedAthleteModels) As QualifiedAthleteModels
    Set IOrderAssignable_AssignOrder = AssignOrder(Qualifiers)
End Function

Private Property Get IOrderAssignable_PersonsPerGroup() As Long
    IOrderAssignable_PersonsPerGroup = Me.PersonsPerGroup
End Property

Public Function AssignOrder(ByRef Qualifiers As QualifiedAthleteModels) As QualifiedAthleteModels
    Dim i As Long
    Dim currentLevel As Long
    Dim row As QualifiedAthleteModel
    Dim rows As Collection

    currentLevel = 0
    Set rows = Qualifiers.Sort(False).All()

    Set AssignOrder = New QualifiedAthleteModels
    For i = 1 To rows.Count
        Set row = rows.Item(i)
        row.Order = i
        Call AssignOrder.Add(row)
    Next i
End Function

Public Property Get PersonsPerGroup() As Long
    PersonsPerGroup = m_PersonPerGroup
End Property
