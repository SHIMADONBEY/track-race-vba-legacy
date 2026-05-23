Attribute VB_Name = "TrackResultRankingService"
'namespace=vba-files/track-race/modules/ranking
Option Explicit
Option Private Module

Public Sub UpdateTotalRanking(TargetRoundId As Long)
    Dim resultRecords As TrackResultOrderModels: Set resultRecords = TrackOrderResultsReposotory.ReadAllTrackResult().FilterByRound(TargetRoundId)

    Dim rankingRecords As TrackResultRankingModels: Set rankingRecords = New TrackResultRankingModels
    Dim resultRecord As TrackResultOrderModel
    For Each resultRecord In resultRecords.All()
        If (resultRecord.IsRankable()) Then
            With New TrackResultRankingModel
                Call .Initialize( _
                        resultRecord.Id _
                        , resultRecord.RoundId _
                        , resultRecord.Group _
                        , resultRecord.Rank _
                        , resultRecord.RealResult _
                        , resultRecord.Result _
                )

                Call rankingRecords.Add(.Self())
            End With
        End If
    Next resultRecord

    Dim rankedRecords As TrackResultRankingModels: Set rankedRecords = rankingRecords.SetRank()

    Dim roundInfo As EventRoundModel: Set roundInfo = EventRoundRepository.ReadAllRounds().Item(TargetRoundId)
    Dim qualifiedRecords As TrackResultRankingModels: Set qualifiedRecords = rankingRecords.SetQualification(roundInfo.GroupCount, roundInfo.PlaceEachGroup, roundInfo.AdditionCount)
    
    Dim updateRecords As Collection: Set updateRecords = New Collection
    Dim transactionDateTime As Date: transactionDateTime = Now

    Dim finishedGroupCount As Long: finishedGroupCount = TrackSubResultGroupRepository.ReadAllRecords().FilterByRound(TargetRoundId).Count
    For Each resultRecord In resultRecords.All()
        Dim rankedRecord As TrackResultRankingModel: Set rankedRecord = rankedRecords.Item(resultRecord.Id)
        Dim qualifiedRecord As TrackResultRankingModel: Set qualifiedRecord = qualifiedRecords.Item(resultRecord.Id)

        If Not((rankedRecord Is Nothing) Or (qualifiedRecord Is Nothing)) Then
            With resultRecord
                .TotalRank = rankedRecord.TotalRanking
                .SameResult = rankedRecord.SameResult
                .ResultRank = qualifiedRecord.TotalRanking

                If (qualifiedRecord.QualifiedByRank) Then
                    .NextQualified = "Q"
                ElseIf (finishedGroupCount < roundInfo.GroupCount) Then
                    ' 全組終了するまでは、記録による通過者は表示させない.
                    .NextQualified = ""
                ElseIf (qualifiedRecord.QualifiedByResult) Then
                    .NextQualified = "q"
                ElseIf (qualifiedRecord.DrawToQualify) Then
                    .NextQualified = "?"
                Else
                    .NextQualified = ""
                End If
                ' .UpdatedAt = transactionDateTime
                Call updateRecords.Add(.Self(), .Id)
            End With
        End If
    Next resultRecord

    Call TrackOrderResultsReposotory.Upsert(updateRecords)
End Sub
