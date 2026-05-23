VERSION 1.0 CLASS
BEGIN
MultiUse = -1 ' True
END
Attribute VB_Name = "TrackResultRankingModels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/modules/ranking/model
Option Explicit
Implements IJsonConvertable

Private m_Records As Object

Private Sub Class_Initialize()
    Set m_Records = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set m_Records = Nothing
End Sub

Public Sub Add(Item As TrackResultRankingModel)
    Call m_Records.Add(Item.Id, Item)
End Sub

Public Sub AddRange(Items As Collection)
    Dim recordItem As TrackResultRankingModel
    For Each recordItem In Items
        Call Me.Add(recordItem)
    Next recordItem
End Sub

Public Function All() As Collection
    Dim recordItems As Collection: Set recordItems = New Collection
    Dim recordId As Variant
    For Each recordId In m_Records
        Dim recordItem As TrackResultRankingModel: Set recordItem = m_Records.Item(recordId)
        Call recordItems.Add(recordItem)
    Next recordId

    Set All = recordItems
End Function

Public Function Count() As Long
    Count = m_Records.Count()
End Function

Public Function Item(Id As Variant) As TrackResultRankingModel
    If Not (m_Records.Exists(Id)) Then
        Set Item = Nothing
        Exit Function
    End If

    Set Item = m_Records.Item(Id)
End Function

Public Function SetRank() As TrackResultRankingModels
    Dim recordItems As Collection: Set recordItems = New Collection
    Dim recordId As Variant
    For Each recordId In m_Records
        Call recordItems.Add(Me.Item(recordId).Clone(), recordId)
    Next recordId 

    Dim vRanked As Collection: Set vRanked = Ranking(recordItems, New RecordRankingComparator)
    Set SetRank = New TrackResultRankingModels
    Call SetRank.AddRange(vRanked)
End Function

Public Function SetQualification(GroupCount As Long, QualifyByRank As Long, QualifyByResult As Long) As TrackResultRankingModels
    ' 着順通過者の確認
    Dim recordItems As Collection: Set recordItems = New Collection
    Dim recordId As Variant
    For Each recordId In m_Records
        With Me.Item(recordId).Clone()
            .QualifiedByRank = (.Group <= GroupCount) And (.Rank <= QualifyByRank)
            Call recordItems.Add(.Self())
        End With
    Next recordId 

    Dim vSorted As Collection: Set vSorted = Ranking(recordItems, New QualificationComparator)

    Dim vTargetNumber As Long: vTargetNumber = GroupCount * QualifyByRank + QualifyByResult
    Dim vRotterry As Long: vRotterry = -2147483648

    Dim i As Long
    For i = recordItems.Count To 1 Step -1
        Dim recordItem As TrackResultRankingModel: Set recordItem = vSorted.Item(i)
        If (recordItem.TotalRanking > vTargetNumber) Then
            ' Continues
        ElseIf (vRotterry <= 0 And i > vTargetNumber) Then
            vRotterry = recordItem.TotalRanking
            recordItem.DrawToQualify = True
        ElseIf (vRotterry > 0 And recordItem.TotalRanking = vRotterry) Then
            recordItem.DrawToQualify = True
        ElseIf Not(recordItem.QualifiedByRank) Then
            recordItem.QualifiedByResult = True
        End If
    Next i

    Set SetQualification = New TrackResultRankingModels
    Call SetQualification.AddRange(vSorted)
End Function

Private Function IJsonConvertable_ToJson() As String
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(Me.All)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function

Private Function Ranking(RecordCollection As Collection, Comparator As IObjectComparator) As Collection
    Set Ranking = CollectionUtil.SortObjectCollection(RecordCollection, Comparator)

    Dim i As Long
    Dim sameCount As Long: sameCount = 0
    
    Dim prevOfficial As TimeResultValue
    Dim prevRecord As TrackResultRankingModel: Set prevRecord = Nothing
    Dim currRecord As TrackResultRankingModel
    For i = 1 To Ranking.Count
        Set currRecord = Ranking.Item(i)
        If (i = 1) Then 
            currRecord.TotalRanking = i
        Else
            If (prevRecord.OfficialResult.Compare(currRecord.OfficialResult) = 0) Then 
                If (prevOfficial Is Nothing) Then
                    prevRecord.SameResult = 2
                    Set prevOfficial = prevRecord.OfficialResult
                End If
                currRecord.SameResult = 2
            Else
                Set prevOfficial = Nothing
                currRecord.SameResult = 0
            End If

            If (comparator.Compare(prevRecord, currRecord) = 0) Then
                If (sameCount = 0) Then
                    prevRecord.SameResult = prevRecord.SameResult + 1
                End If

                currRecord.SameResult = currRecord.SameResult + 1
                sameCount = sameCount + 1
            Else
                ' 昇順に並び替え済みのため、Compare すると同順以外は-1が返される.
                sameCount = 0
            End If

            currRecord.TotalRanking = i - sameCount
        End If
        Set prevRecord = currRecord
    Next i
End Function
