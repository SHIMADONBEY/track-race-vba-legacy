VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EquallyGroupAssigner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/modules/program-entry/policy/group
Option Explicit
Implements IGroupAssignable

Private Function IGroupAssignable_AssignGroup(ByRef Qualifiers As QualifiedAthleteModels, ByVal PersonsPerGroup As Long) As QualifiedAthleteModels()
    Dim returnGroups() As QualifiedAthleteModels
    Dim allQualifiersCount As Long
    Dim numGroupCount As Long
    Dim i As Long
    
    allQualifiersCount = Qualifiers.Count
    numGroupCount = Fix((allQualifiersCount + PersonsPerGroup - 1) / PersonsPerGroup)
    
    ReDim returnGroups(0 To numGroupCount - 1)
    
    For i = 0 To numGroupCount - 1
        Set returnGroups(i) = New QualifiedAthleteModels
    Next i
    
    Dim sortedAllList As QualifiedAthleteModels: Set sortedAllList = Qualifiers.Sort()
    
    Dim loBound As Long
    Dim upBound As Long
    Dim direction As Long
    Dim cursor As Long
    Dim record As QualifiedAthleteModel
    
    loBound = LBound(returnGroups)
    upBound = UBound(returnGroups)
    
    cursor = loBound
    direction = 0
    For Each record In sortedAllList.All()
        Call returnGroups(cursor).Add(record)
        If cursor = loBound Then
            direction = direction + 1
        End If
        If cursor = upBound Then
            direction = direction - 1
        End If
        
        cursor = cursor + direction
    Next record

    Dim groupList() As Long
    groupList = GenerateGroupList(numGroupCount, returnGroups)
    
    For i = 0 To numGroupCount - 1
        For Each record In returnGroups(i).All
            record.Group = groupList(i)
        Next record
    Next i
       
    IGroupAssignable_AssignGroup = returnGroups
End Function

Private Function GenerateGroupList(ByVal GroupCount As Long, ByRef GroupedQualifers() As QualifiedAthleteModels) As Long()
    Dim groupIndexes() As Long
    Dim groupWeight() As Single
    Dim returnList() As Long

    Dim qualifiersGroup As QualifiedAthleteModels
    
    ReDim groupIndexes(1 To GroupCount)
    ReDim groupWeight(1 To GroupCount)
    Dim i As Long
    For i = 1 To GroupCount
        groupIndexes(i) = i - 1
        groupWeight(i) = GroupedQualifers(i - 1).Count + Rnd
    Next
    
    Dim topIndex As Long, bottomIndex As Long
    Dim direction As Long
    
    topIndex = 1
    bottomIndex = GroupCount
    direction = 1
    
    Do While topIndex < bottomIndex
        Dim startIndex As Long
        Dim endIndex As Long
        Dim current As Long
        Dim neighbor As Long
        
        If direction = 1 Then
            startIndex = topIndex
            endIndex = bottomIndex
        Else
            startIndex = bottomIndex
            endIndex = topIndex
        End If
    
        For current = startIndex To endIndex - direction Step direction
            neighbor = current + direction
            
            If Sgn(groupWeight(current) - groupWeight(neighbor)) = direction Then
                Dim tempWeight As Single
                Dim tempIndex As Long
                
                tempWeight = groupWeight(current)
                tempIndex = groupIndexes(current)
                
                groupWeight(current) = groupWeight(neighbor)
                groupIndexes(current) = groupIndexes(neighbor)
                
                groupWeight(neighbor) = tempWeight
                groupIndexes(neighbor) = tempIndex
            End If
        Next

        If direction = 1 Then
            bottomIndex = bottomIndex - direction
        Else
            topIndex = topIndex - direction
        End If
        direction = -direction
    Loop
    
    ReDim returnList(0 To GroupCount - 1)
    
    For i = 1 To GroupCount
        returnList(groupIndexes(i)) = i
    Next i
    
    GenerateGroupList = returnList
End Function
