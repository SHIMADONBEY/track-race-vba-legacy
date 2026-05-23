Attribute VB_Name = "ResultRankingWritingService"
'namespace=vba-files/track-race/modules/html/output/result-ranking
Option Explicit
Option Private Module

Public Function WriteRanking(TargetRound As EventRoundModel)
    Dim pRoundId As Long: pRoundId = TargetRound.Id
    Dim athletes As ResultRankingModels: Set athletes = New ResultRankingModels

    ' åãâ ÉfÅ[É^ÇÃéÊìæ
    Dim resultRecords As TrackResultOrderModels: Set resultRecords = TrackOrderResultsReposotory.ReadAllTrackResult().FilterByRound(pRoundId)
    Dim resultRecord As TrackResultOrderModel
    Dim startList As StartListOrderModels: Set startList = ProgramEntryRepository.ReadAllStartList(True).FilterByRound(pRoundId)
    Dim groupList As TrackSubResultGroupModels: Set groupList = TrackSubResultGroupRepository.ReadAllRecords().FilterByRound(pRoundId)

    For Each resultRecord In resultRecords.All()
        Dim startData As StartListOrderModel: Set startData = startList.FindOrder(pRoundId, resultRecord.Group, resultRecord.Order)
        Dim groupData As TrackSubResultGroupModel: Set groupData = groupList.FindByGroup(pRoundId, resultRecord.Group)

        With New ResultRankingModel
            Call .Initialize( _
                    resultRecord.Id _
                    , resultRecord.RoundId _
                    , resultRecord.Group _
                    , resultRecord.Order _
                    , resultRecord.Rank _
                    , resultRecord.Result _
                    , resultRecord.RealResult _
                    , resultRecord.Remark _
                    , resultRecord.ReactionTime _
                    , resultRecord.NotStarted _
                    , resultRecord.NotFinished _
                    , resultRecord.DisqualifiedReason _
                    , resultRecord.AdvancedNextRound _
                    , resultRecord.DemoEntry _
                    , resultRecord.TotalRank _
                    , resultRecord.SameResult _
                    , resultRecord.ResultRank _
                    , resultRecord.NextQualified _
                    , resultRecord.Score _
                    , startData.Accumaltive  _
                    , startData.Person _
                    , groupData.Wind _
            )
            Call athletes.Add(.Self())
        End With
    Next resultRecord

    Dim pRoundName As String: pRoundName = TargetRound.Name
    Dim resultRankingData As ResultRankingOutputModel: Set resultRankingData = New ResultRankingOutputModel
    Call resultRankingData.Initialize( _
            pRoundName _
            , CompetitionRepository.ReadSettinng() _
            , EventRoundRepository.ReadAllRounds() _
            , GameRecordRepository.ReadAllGameRecords(true) _
            , athletes.SortByTotalRanking() _
    )

    Dim htmlReader As HtmlTemplateWrapper: Set htmlReader = New HtmlTemplateWrapper
    htmlReader.Initialize (CompetitionRepository.ReadSettinng.HtmlTemplate.TrackRankingListPath)

    Dim writer As HtmlResultRankingWriter: Set writer = New HtmlResultRankingWriter
    Call writer.Initialize(htmlReader.HtmlCollection, resultRankingData)

    Dim htmlWriter As TrackHtmlOutputer: Set htmlWriter = New TrackHtmlOutputer
    Call htmlWriter.Initialize(writer)

    Dim vFileName As String: vFileName = "Ranking_" & CompetitionRepository.ReadSettinng.Category & CompetitionRepository.ReadSettinng.Sex & CompetitionRepository.ReadSettinng.EventName & pRoundName & ".html"
    WriteRanking = htmlWriter.WriteFile(vFileName)
End Function
