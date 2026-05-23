Attribute VB_Name = "CompetitionRepository"
'namespace=vba-files/track-race/repository
Option Explicit
Option Private Module

Private Enum EventSettingRowEnum
    IDX_CATEGORY = 2
    IDX_SEX
    IDX_EVENT_NAME
    IDX_SUPECIFICATION
    IDX_PERSON_PER_GROUP
    IDX_ROUNDS_COUNT
    IDX_OPERATION_WIND_GAUGE
    IDX_MEASUREMENT_REACTION_TIME
    IDX_RESULT_PRECISION
    IDX_MINUTE_DELIMITER = 12
    IDX_SECOND_DELIMITER
    IDX_TRACK_TIME_LENGTH
    IDX_SPLIT_TIME_PRECISION
    IDX_INTERVAL_LENGTH_1
    IDX_INTERVAL_LENGTH_2
    IDX_SCORING_COMBINED_EVENTS
    IDX_SCORING_INDEX_NUMBER_1
    IDX_SCORING_INDEX_NUMBER_2
    IDX_SCORING_INDEX_NUMBER_3
    IDX_LANE_COUNT = 24
    IDX_TOP_LANES
    IDX_MIDDLE_LANES
    IDX_BOTTOM_LANES
    IDX_EXTRA_LANES
    IDX_COMPETITION = 32
    IDX_FACILITY
    IDX_FACILITY_PLACE
    IDX_COMPETITION_YEAR
    IDX_COMPETITION_DATES
    IDX_TEMPLATES = 39
End Enum

Private Const TEMPLATES_COUNT As Long = 4

Private Const INPUT_COLUMN As Long = 2
Private Const CODE_COLUMN As Long = 3

Private m_Instance As TrackEventSettingModel
Private m_Competition As CompetitionInfoModel

Private Function ThisSheet() As Worksheet
    Set ThisSheet = EventSettingSheet
End Function

Public Function ReadSettinng(Optional Reload As Boolean = False) As TrackEventSettingModel
    If (m_Instance Is Nothing Or Reload) Then
        Dim sh As Worksheet: Set sh = ThisSheet
        Dim laneSetting As LaneSettingModel: Set laneSetting = New LaneSettingModel
        Call laneSetting.Initialize( _
            sh.Cells(EventSettingRowEnum.IDX_LANE_COUNT, INPUT_COLUMN).Value _
            , sh.Cells(EventSettingRowEnum.IDX_TOP_LANES, INPUT_COLUMN).Value _
            , sh.Cells(EventSettingRowEnum.IDX_MIDDLE_LANES, INPUT_COLUMN).Value _
            , sh.Cells(EventSettingRowEnum.IDX_BOTTOM_LANES, INPUT_COLUMN).Value _
            , sh.Cells(EventSettingRowEnum.IDX_EXTRA_LANES, INPUT_COLUMN).Value _
        )

        Dim templateSetting As HtmlTemplatePathsModel: Set templateSetting = New HtmlTemplatePathsModel
        Call templateSetting.Initialize( _
            ThisWorkbook.Path & "\" & sh.Cells(EventSettingRowEnum.IDX_TEMPLATES, INPUT_COLUMN).Value _
            , ThisWorkbook.Path & "\" & sh.Cells(EventSettingRowEnum.IDX_TEMPLATES + 1, INPUT_COLUMN).Value _
            , ThisWorkbook.Path & "\" & sh.Cells(EventSettingRowEnum.IDX_TEMPLATES + 2, INPUT_COLUMN).Value _
            , ThisWorkbook.Path & "\" & sh.Cells(EventSettingRowEnum.IDX_TEMPLATES + 3, INPUT_COLUMN).Value _
            , ThisWorkbook.Path & "\" & sh.Cells(EventSettingRowEnum.IDX_TEMPLATES + 4, INPUT_COLUMN).Value _
        )

        Set m_Instance = New TrackEventSettingModel

        Call m_Instance.Initialize( _
            sh.Cells(EventSettingRowEnum.IDX_CATEGORY, INPUT_COLUMN).Value _
            , sh.Cells(EventSettingRowEnum.IDX_SEX, INPUT_COLUMN).Value _
            , sh.Cells(EventSettingRowEnum.IDX_EVENT_NAME, INPUT_COLUMN).Value _
            , sh.Cells(EventSettingRowEnum.IDX_SUPECIFICATION, INPUT_COLUMN).Value _
            , sh.Cells(EventSettingRowEnum.IDX_PERSON_PER_GROUP, INPUT_COLUMN).Value _
            , sh.Cells(EventSettingRowEnum.IDX_ROUNDS_COUNT, INPUT_COLUMN).Value _
            , sh.Cells(EventSettingRowEnum.IDX_OPERATION_WIND_GAUGE, CODE_COLUMN).Value _
            , sh.Cells(EventSettingRowEnum.IDX_MEASUREMENT_REACTION_TIME, CODE_COLUMN).Value _
            , sh.Cells(EventSettingRowEnum.IDX_RESULT_PRECISION, CODE_COLUMN).Value _
            , sh.Cells(EventSettingRowEnum.IDX_MINUTE_DELIMITER, INPUT_COLUMN).Value _
            , sh.Cells(EventSettingRowEnum.IDX_SECOND_DELIMITER, INPUT_COLUMN).Value _
            , sh.Cells(EventSettingRowEnum.IDX_TRACK_TIME_LENGTH, CODE_COLUMN).Value _
            , sh.Cells(EventSettingRowEnum.IDX_SPLIT_TIME_PRECISION, CODE_COLUMN).Value _
            , sh.Cells(EventSettingRowEnum.IDX_INTERVAL_LENGTH_1, INPUT_COLUMN).Value _
            , sh.Cells(EventSettingRowEnum.IDX_INTERVAL_LENGTH_2, INPUT_COLUMN).Value _
            , sh.Cells(EventSettingRowEnum.IDX_SCORING_COMBINED_EVENTS, CODE_COLUMN).Value _
            , sh.Cells(EventSettingRowEnum.IDX_SCORING_INDEX_NUMBER_1, INPUT_COLUMN).Value _
            , sh.Cells(EventSettingRowEnum.IDX_SCORING_INDEX_NUMBER_2, INPUT_COLUMN).Value _
            , sh.Cells(EventSettingRowEnum.IDX_SCORING_INDEX_NUMBER_3, INPUT_COLUMN).Value _
            , laneSetting _
            , templateSetting _
        )
    End If

    Set ReadSettinng = m_Instance
End Function

Public Function ReadCompetition(Optional Reload As Boolean = False) As CompetitionInfoModel
    If (Reload Or m_Competition Is Nothing) Then
        With New CompetitionInfoModel
            Call .Initialize( _
                    ThisSheet.Cells(EventSettingRowEnum.IDX_COMPETITION, INPUT_COLUMN).Value _
                    , "" _
                    , ThisSheet.Cells(EventSettingRowEnum.IDX_COMPETITION, CODE_COLUMN).Value _
                    , ThisSheet.Cells(EventSettingRowEnum.IDX_FACILITY, INPUT_COLUMN).Value _
                    , "" _
                    , ThisSheet.Cells(EventSettingRowEnum.IDX_FACILITY, CODE_COLUMN).Value _
                    , ThisSheet.Cells(EventSettingRowEnum.IDX_FACILITY_PLACE, INPUT_COLUMN).Value _
                    , ThisSheet.Cells(EventSettingRowEnum.IDX_COMPETITION_YEAR, INPUT_COLUMN).Value _
                    , ThisSheet.Cells(EventSettingRowEnum.IDX_COMPETITION_DATES, INPUT_COLUMN).Value _
                    , ThisSheet.Cells(EventSettingRowEnum.IDX_COMPETITION_DATES, CODE_COLUMN).Value _
            )
            Set m_Competition = .Self()
        End With
    End If

    Set ReadCompetition = m_Competition
End Function

Public Sub LoadFromCentral(CompetitionInfo As CompetitionInfoModel, TrackEvent As TrackEventSettingModel)
    Set m_Competition = Nothing
    Set m_Instance = Nothing

    ThisSheet.Cells(EventSettingRowEnum.IDX_COMPETITION, INPUT_COLUMN).Value                = CompetitionInfo.CompetitionName
    ThisSheet.Cells(EventSettingRowEnum.IDX_COMPETITION, CODE_COLUMN).Value                 = "'" & CompetitionInfo.CompetitionCode
    ThisSheet.Cells(EventSettingRowEnum.IDX_FACILITY, INPUT_COLUMN).Value                   = CompetitionInfo.FacilityName
    ThisSheet.Cells(EventSettingRowEnum.IDX_FACILITY, CODE_COLUMN).Value                    = "'" & CompetitionInfo.FacilityCode
    ThisSheet.Cells(EventSettingRowEnum.IDX_FACILITY_PLACE, INPUT_COLUMN).Value             = CompetitionInfo.FacilityPlace
    ThisSheet.Cells(EventSettingRowEnum.IDX_COMPETITION_YEAR, INPUT_COLUMN).Value           = CompetitionInfo.CompetitionYear
    ThisSheet.Cells(EventSettingRowEnum.IDX_COMPETITION_DATES, INPUT_COLUMN).Value          = CompetitionInfo.CompetitionDateStart
    ThisSheet.Cells(EventSettingRowEnum.IDX_COMPETITION_DATES, CODE_COLUMN).Value           = CompetitionInfo.CompetitionDateEnd

    ThisSheet.Cells(EventSettingRowEnum.IDX_CATEGORY, INPUT_COLUMN).Value                   = TrackEvent.Category
    ThisSheet.Cells(EventSettingRowEnum.IDX_SEX, INPUT_COLUMN).Value                        = TrackEvent.Sex
    ThisSheet.Cells(EventSettingRowEnum.IDX_EVENT_NAME, INPUT_COLUMN).Value                 = TrackEvent.EventName
    ThisSheet.Cells(EventSettingRowEnum.IDX_SUPECIFICATION, INPUT_COLUMN).Value             = TrackEvent.Supecification
    ThisSheet.Cells(EventSettingRowEnum.IDX_PERSON_PER_GROUP, INPUT_COLUMN).Value           = TrackEvent.PersonPerGroup
    ThisSheet.Cells(EventSettingRowEnum.IDX_ROUNDS_COUNT, INPUT_COLUMN).Value               = TrackEvent.RoundsCount
    ThisSheet.Cells(EventSettingRowEnum.IDX_OPERATION_WIND_GAUGE, INPUT_COLUMN).Value       = Range("Code_Wind").Cells(IIf(TrackEvent.OperationWindGauge, 2, 1), 2).Value
    ThisSheet.Cells(EventSettingRowEnum.IDX_MEASUREMENT_REACTION_TIME, INPUT_COLUMN).Value  = Range("Code_Reaction").Cells(IIf(TrackEvent.MeasurementReactionTime, 2, 1), 2).Value
    ThisSheet.Cells(EventSettingRowEnum.IDX_MINUTE_DELIMITER, INPUT_COLUMN).Value           = "'" & TrackEvent.MinuteDelimiter
    ThisSheet.Cells(EventSettingRowEnum.IDX_SECOND_DELIMITER, INPUT_COLUMN).Value           = "'" & TrackEvent.SecondDelimiter

    ' TODO: Combined Event Settings
    ThisSheet.Cells(EventSettingRowEnum.IDX_SCORING_COMBINED_EVENTS, INPUT_COLUMN).Value    = Range("Code_Point_1").Cells(IIf(TrackEvent.ScoringCombinedEvents, 2, 1), 2).Value

    ' Lane Settings
    ThisSheet.Cells(EventSettingRowEnum.IDX_LANE_COUNT, INPUT_COLUMN).Value                 = TrackEvent.LaneSetting.LaneCount
    ThisSheet.Cells(EventSettingRowEnum.IDX_TOP_LANES, INPUT_COLUMN).Value                  = StringUtil.JoinArrayToString(TrackEvent.LaneSetting.TopLanes, ",")
    ThisSheet.Cells(EventSettingRowEnum.IDX_MIDDLE_LANES, INPUT_COLUMN).Value               = StringUtil.JoinArrayToString(TrackEvent.LaneSetting.MiddleLanes, ",")
    ThisSheet.Cells(EventSettingRowEnum.IDX_BOTTOM_LANES, INPUT_COLUMN).Value               = StringUtil.JoinArrayToString(TrackEvent.LaneSetting.BottomLanes, ",")
    ThisSheet.Cells(EventSettingRowEnum.IDX_EXTRA_LANES, INPUT_COLUMN).Value                = StringUtil.JoinArrayToString(TrackEvent.LaneSetting.ExtraLanes, ",")
End Sub
