Attribute VB_Name = "GameRecordsImporter"
'namespace=vba-files/track-race/import
Option Explicit

Public Sub ImportGameRecords(CompetitionEvent As String, GameRecords As String)
    Dim vCompetitionEvent As CompetitionEventModel: Set vCompetitionEvent = JsonConverter.ParseJsonTo(CompetitionEvent, New CompetitionEventModel)
    Dim vGameRecords As GameRecordModels: Set vGameRecords = JsonConverter.ParseJsonTo(GameRecords, New GameRecordModels, New TrackGameRecordsJsonParser)
    Dim vAllGameRecords As GameRecordModels: Set vAllGameRecords = GameRecordRepository.ReadAllGameRecords()

    Call GameRecordRepository.Upsert(vGameRecords.All())
End Sub
