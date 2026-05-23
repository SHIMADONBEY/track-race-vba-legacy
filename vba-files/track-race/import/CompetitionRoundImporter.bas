Attribute VB_Name = "CompetitionRoundImporter"
'namespace=vba-files/track-race/import
Option Explicit

Public Sub ImportCompetitionRounds(CompetitionEvent As String, CompetitionRounds As String)
    Dim vCompetionEvent As CompetitionEventModel: Set vCompetionEvent = JsonConverter.ParseJsonTo(CompetitionEvent, New CompetitionEventModel)
    Dim vCompetitionRounds As CompetitionRoundModels: Set vCompetitionRounds = JsonConverter.ParseJsonTo(CompetitionRounds, New CompetitionRoundModels)

    Call EventRoundRepository.ImportCompetitionRounds(vCompetionEvent, vCompetitionRounds)
End Sub