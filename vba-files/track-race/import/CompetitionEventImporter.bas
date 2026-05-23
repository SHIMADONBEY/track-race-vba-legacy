Attribute VB_Name = "CompetitionEventImporter"
'namespace=vba-files/track-race/import
Option Explicit

Public Sub ImportCompetitionEvent(CompetitionInfo, CompetitionEvent, CompetitionConfigurations)
    Dim vCompetitionInfo As CompetitionInfoModel    : Set vCompetitionInfo = JsonConverter.ParseJsonTo(CompetitionInfo, New CompetitionInfoModel)
    Dim vCompetitionEvent As CompetitionEventModel  : Set vCompetitionEvent = JsonConverter.ParseJsonTo(CompetitionEvent, New CompetitionEventModel)

    Dim vTrackEvent As TrackEventSettingModel: Set vTrackEvent = New TrackEventSettingModel
    With JsonConverter.ParseJson(CompetitionConfigurations)
        Call vTrackEvent.Initialize( _
                vCompetitionEvent.Category _
                , vCompetitionEvent.Sex _
                , vCompetitionEvent.EventName _
                , vCompetitionEvent.Supecification _
                , vCompetitionEvent.PersonPerGroup _
                , vCompetitionEvent.RoundsCount _
                , vCompetitionEvent.OperationWindGauge _
                , vCompetitionEvent.MeasurementReactionTime _
                , 2 _
                , .Item("minute_delimiter") _
                , .Item("second_delimiter") _
                , 400 _
                , 0 _
                , 400 _
                , 1000 _
                , False _
                , "0.00" _
                , "0.00" _
                , "0.00" _
                , LaneSettingFactory.GenerateByLanesCount(.Item("lanes_count")) _
                , Nothing _
        )
    End With

    Call CompetitionRepository.LoadFromCentral(vCompetitionInfo, vTrackEvent)
    Call EventSettingController.UpdateRoundsCount(vTrackEvent.RoundsCount)
End Sub
