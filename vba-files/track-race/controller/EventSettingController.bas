Attribute VB_Name = "EventSettingController"
'namespace=vba-files/track-race/controller
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

Private Enum CompetionRoundsEnum
    IDX_ID = 1
    IDX_ROUND_NAME
    IDX_START_DATE_TIME
    IDX_ORDER_COUNT
    IDX_GROUP_COUNT
    IDX_GROUP_STRATEGY
    IDX_LANE_STRATEGY
    IDX_PLACE_EACH_GROUP
    IDX_ADDITION_COUNT
    IDX_NEXT_ROUNDS
    IDX_GROUP_STRATEGY_CODE
    IDX_LANE_STRATEGY_CODE
End Enum

Private Const COMPETITION_ROUNDS_COLUMN_OFFSET As Long = 10
Private Const COMPETITION_ROUNDS_ROW_HEADER As Long = 3

Private Function ThisSheet() As Worksheet
    Set ThisSheet = EventSettingSheet
End Function

Public Sub UpdateEventData()
    Call CompetitionRepository.ReadSettinng(True)
End Sub

Public Sub UpdateEventRounds()
    Call EventRoundRepository.ReadAllRounds(True)
End Sub

Public Sub OnRoundCountUpdated(Spot As Range)
    If (Application.Intersect(Spot, Range_RoundCounts) Is Nothing) Then
        Exit Sub
    End If

    Call UpdateRoundsCount(Spot.Value)
End Sub

Public Sub OnLaneCountUpdated(Spot As Range)
    If (Application.Intersect(Spot, Range_LaneCount) Is Nothing) Then
        Exit Sub
    End If

    Call UpdateLaneCount(Spot.Value)
End Sub

Public Sub UpdateRoundsCount(RoundsCount As Long)
    If RoundsCount <= 0 Then
        Err.Raise CustomErrorCodeEnum.InvalidRoundCount, "EventSettingController", MessageFactory.Generate("SE002").Prompt
        Exit Sub
    End If

    Dim i As Long

    Call Range_RoundList(Application.WorksheetFunction.CountA(ThisSheet.Columns(COMPETITION_ROUNDS_COLUMN_OFFSET + CompetionRoundsEnum.IDX_ID))).Clear
    With Range_RoundList(RoundsCount)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Weight = xlHairline
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlHairline

        .Columns(CompetionRoundsEnum.IDX_START_DATE_TIME).NumberFormatLocal = "yyyy/MM/dd hh:mm"
        Call .Columns(CompetionRoundsEnum.IDX_GROUP_STRATEGY).Validation.Add(xlValidateList, xlValidAlertStop, xlBetween, "=OFFSET(Code_Group_Strategy, 0, 1, , 1)")
        Call .Columns(CompetionRoundsEnum.IDX_LANE_STRATEGY).Validation.Add(xlValidateList, xlValidAlertStop, xlBetween, "=OFFSET(Code_Lane_Strategy, 0, 1, , 1)")

        .Columns(CompetionRoundsEnum.IDX_NEXT_ROUNDS).FormulaR1C1 = "=RC[-5]*RC[-2]+RC[-1]"
        .Columns(CompetionRoundsEnum.IDX_NEXT_ROUNDS).Borders(xlEdgeRight).Weight = xlThin
        For i = 1 To RoundsCount
            .Cells(i, CompetionRoundsEnum.IDX_ID).Value = i
        Next i
        
        .Columns(CompetionRoundsEnum.IDX_GROUP_STRATEGY_CODE).FormulaR1C1 = "=VLookUp(RC[-5], Offset(Code_Group_Strategy, 0, 1), 2, False)"
        .Columns(CompetionRoundsEnum.IDX_LANE_STRATEGY_CODE).FormulaR1C1 = "=VLookUp(RC[-5], Offset(Code_Lane_Strategy, 0, 1), 2, False)"
    End With

End Sub

Private Sub UpdateLaneCount(LanesCount As Long)
    Dim vLaneSetting As LaneSettingModel: Set vLaneSetting = LaneSettingFactory.GenerateByLanesCount(LanesCount)
    Call Range_LaneSetting().ClearContents

    If (vLaneSetting Is Nothing) Then
        Exit Sub
    End If

    With Range_LaneSetting()
        .Cells(1, 1).Value = StringUtil.JoinArrayToString(vLaneSetting.TopLanes, ",")
        .Cells(2, 1).Value = StringUtil.JoinArrayToString(vLaneSetting.MiddleLanes, ",")
        .Cells(3, 1).Value = StringUtil.JoinArrayToString(vLaneSetting.BottomLanes, ",")
        .Cells(4, 1).Value = StringUtil.JoinArrayToString(vLaneSetting.ExtraLanes, ",")
    End With
End Sub

Private Function Range_RoundList(Optional RowsCount As Long = 0) As Range
    Set Range_RoundList = ThisSheet.Cells(COMPETITION_ROUNDS_ROW_HEADER + 1, COMPETITION_ROUNDS_COLUMN_OFFSET + 1).Resize(RowsCount, 20)
End Function

Private Function Range_LaneSetting()
    Set Range_LaneSetting = ThisSheet.Range(Cells(EventSettingRowEnum.IDX_TOP_LANES, INPUT_COLUMN), Cells(EventSettingRowEnum.IDX_EXTRA_LANES, INPUT_COLUMN))
End Function

Private Function Range_RoundCounts() As Range
    Set Range_RoundCounts = ThisSheet.Cells(EventSettingRowEnum.IDX_ROUNDS_COUNT, INPUT_COLUMN)
End Function

Private Function Range_LaneCount() As Range
    Set Range_LaneCount = ThisSheet.Cells(EventSettingRowEnum.IDX_LANE_COUNT, INPUT_COLUMN)
End Function
