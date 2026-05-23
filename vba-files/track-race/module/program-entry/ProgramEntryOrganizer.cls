VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgramEntryOrganizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/modules/program-entry
Option Explicit

Private m_GroupAssigner As IGroupAssignable
Private m_LaneAssigner As IOrderAssignable

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
End Sub

Private Sub Class_Terminate()
End Sub

Public Sub Initialize(GroupAssigner As IGroupAssignable, LaneAssigner As IOrderAssignable)
    If m_Initialized Then Err.Raise (91)
    Set m_GroupAssigner = GroupAssigner
    Set m_LaneAssigner = LaneAssigner
    m_Initialized = True
End Sub

Public Function Organize(ByRef Qualifiers As QualifiedAthleteModels, ByVal PersionsPerGroup As Long) As QualifiedAthleteModels
    If Not m_Initialized Then Call Err.Raise(91)
    Dim groupedQualifiers() As QualifiedAthleteModels: groupedQualifiers = m_GroupAssigner.AssignGroup(Qualifiers, PersionsPerGroup)

    Dim organizedQualifiers As QualifiedAthleteModels: Set organizedQualifiers = New QualifiedAthleteModels
    Dim i As Long
    For i = LBound(groupedQualifiers) To UBound(groupedQualifiers)
        Dim assignedOrders As QualifiedAthleteModels: Set assignedOrders = m_LaneAssigner.AssignOrder(groupedQualifiers(i))
        Call organizedQualifiers.AddRange(assignedOrders.All())
    Next i

    Set Organize = organizedQualifiers
End Function
