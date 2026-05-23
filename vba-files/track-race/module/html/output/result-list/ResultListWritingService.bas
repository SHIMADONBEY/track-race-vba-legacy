Attribute VB_Name = "ResultListWritingService"
'namespace=vba-files/track-race/modules/html/output/result-list
Option Explicit
Option Private Module

Public Function WriteResultList(TargetRound As EventRoundModel) As String
    Dim pRoundId As Long: pRoundId = TargetRound.Id
    Dim resultData As TrackResultRoundModel: Set resultData = New TrackResultRoundModel
    Dim mergedResults As TrackResultOrderModels: Set mergedResults = TrackOrderResultsReposotory.ReadAllTrackResult().FilterByRound(pRoundId).SortByRank()
    Dim splitTimeList As TrackSubResultOrderModels: Set splitTimeList = TrackSubResultOrderRepository.ReadAllRecords().FilterLapTimeByRound(pRoundId)
    Dim startList As StartListOrderModels: Set startList = ProgramEntryRepository.ReadAllStartList(true).FilterByRound(pRoundId)
    Dim groupList As TrackSubResultGroupModels: Set groupList = TrackSubResultGroupRepository.ReadAllRecords().FilterByRound(pRoundId)

    ' スタートリストデータの読み込み.
    Call mergedResults.MergeStartList(startList)
    Call splitTimeList.MergeStartList(startList)

    Dim vGroupCount As Long: vGroupCount = TargetRound.GroupCount
    With resultData
        Call .Initialize(vGroupCount, pRoundId, groupList)
        Call .AddResults(mergedResults)
        If (CompetitionRepository.ReadSettinng().SplitTimePoints().Count > 0) Then
            Call .AddSplitTimes(splitTimeList)
        End If
    End With

    Dim pRoundName As String: pRoundName = TargetRound.Name
    Dim resultListData As ResultListOutputModel: Set resultListData = New ResultListOutputModel
    Call resultListData.Initialize( _
            pRoundName _
            , CompetitionRepository.ReadSettinng() _
            , EventRoundRepository.ReadAllRounds() _
            , GameRecordRepository.ReadAllGameRecords(true) _
            , resultData _
    )

    Dim htmlReader As HtmlTemplateWrapper: Set htmlReader = New HtmlTemplateWrapper
    htmlReader.Initialize (CompetitionRepository.ReadSettinng.HtmlTemplate.ResultTemlatePath)

    Dim writer As HtmlResultListWriter: Set writer = New HtmlResultListWriter
    Call writer.Initialize(htmlReader.HtmlCollection, resultListData)

    Dim htmlWriter As TrackHtmlOutputer : Set htmlWriter = New TrackHtmlOutputer
    Call htmlWriter.Initialize(writer)

    Dim vFileName As String: vFileName = "Result_" & CompetitionRepository.ReadSettinng.Category & CompetitionRepository.ReadSettinng.Sex & CompetitionRepository.ReadSettinng.EventName & pRoundName & ".html"
    WriteResultList = htmlWriter.WriteFile(vFileName)
End Function
