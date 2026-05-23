VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "DescendGroupAssigner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/module/program-entry/policy/group
Option Explicit
Implements IGroupAssignable

Private Function IGroupAssignable_AssignGroup(ByRef Qualifiers As QualifiedAthleteModels, ByVal PersonsPerGroup As Long) As QualifiedAthleteModels()
    Dim pNumOfQualifiers As Long: pNumOfQualifiers = Qualifiers.Count()
    Dim vNumOfGroups As Long: vNumOfGroups = Fix((pNumOfQualifiers + PersonsPerGroup - 1) / PersonsPerGroup)

    Dim aGroups() As QualifiedAthleteModels
    ReDim aGroups(0 To vNumOfGroups - 1)

    Dim i As Long
    Dim j As Long
    For i = 0 To vNumOfGroups - 1
        Set aGroups(i) = New QualifiedAthleteModels
    Next i

    Dim vSortedQualifiers As QualifiedAthleteModels: Set vSortedQualifiers = Qualifiers.Sort(True)

    i = 0
    j = 0

    Dim vComparator As IObjectComparator: Set vComparator = New QualifiedRankComparator
    Dim vRecord As QualifiedAthleteModel
    Dim vPrevious As QualifiedAthleteModel
    For Each vRecord In vSortedQualifiers.All()
        If (j < PersonsPerGroup) Then
            j = j + 1
        ElseIf (vPrevious Is Nothing) Then
            ' 前のレコードがない場合は次の組へ移動
            i = i + 1
            j = 0
        ElseIf (vComparator.Compare(vRecord, vPrevious) = 0) Then
            ' 資格成績同順位は同じ組で振り分け
            j = j + 1
        Else
            i = i + 1
            j = 0
        End If

        vRecord.Group = i + 1
        Call aGroups(i).Add(vRecord)
        Set vPrevious = vRecord
    Next vRecord

    IGroupAssignable_AssignGroup = aGroups
End Function
