Attribute VB_Name = "EventRoundRepository"
'namespace=vba-files/track-race/repository
Option Explicit
Option Private Module

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

Private m_Repository As EventRoundModels

Private Function ThisSheet() As Worksheet
    Set ThisSheet = EventSettingSheet
End Function

Public Function ReadAllRounds(Optional Reload As Boolean) As EventRoundModels
    If (Reload Or m_Repository Is Nothing) Then
        Dim vRecords As EventRoundModels: Set vRecords = New EventRoundModels

        Dim rowRange As Range
        For Each rowRange In RangeRoundList().rows()
            Dim vRecord As EventRoundModel: Set vRecord = New EventRoundModel
            Call vRecord.Initialize( _
                    rowRange.Cells(1, CompetionRoundsEnum.IDX_ID).Value _
                    , rowRange.Cells(1, CompetionRoundsEnum.IDX_ROUND_NAME).Value _
                    , rowRange.Cells(1, CompetionRoundsEnum.IDX_START_DATE_TIME).Value _
                    , rowRange.Cells(1, CompetionRoundsEnum.IDX_ORDER_COUNT).Value _
                    , rowRange.Cells(1, CompetionRoundsEnum.IDX_GROUP_COUNT).Value _
                    , rowRange.Cells(1, CompetionRoundsEnum.IDX_GROUP_STRATEGY_CODE).Value _
                    , rowRange.Cells(1, CompetionRoundsEnum.IDX_LANE_STRATEGY_CODE).Value _
                    , rowRange.Cells(1, CompetionRoundsEnum.IDX_PLACE_EACH_GROUP).Value _
                    , rowRange.Cells(1, CompetionRoundsEnum.IDX_ADDITION_COUNT).Value _
            )

            If Not (vRecords.FindByName(vRecord.Name) Is Nothing) Then
                ' ラウンド名が重複しているので、入力エラー
                Err.Raise CustomErrorCodeEnum.DuplicateRoundName, "EventRoundRepository", MessageFactory.Generate("SW017").Prompt(vRecord.Name)
            End If

            Call vRecords.Add(vRecord)
        Next rowRange

        Set m_Repository = vRecords
    End If

    Set ReadAllRounds = m_Repository
End Function

Public Sub ImportCompetitionRounds(CompetitionEvent As CompetitionEventModel, CompetitionRounds As CompetitionRoundModels)
    Dim vRecordRow As Range
    For Each vRecordRow In RangeRoundList().Rows
        Dim pRoundId As Long: pRoundId = vRecordRow.Cells(1, CompetionRoundsEnum.IDX_ID)
        Dim vRecordItem As CompetitionRoundModel: Set vRecordItem = CompetitionRounds.FindByEventRound(CompetitionEvent.Id, pRoundId)
        If Not (vRecordItem Is Nothing) Then
            With vRecordRow
                .Cells(1, CompetionRoundsEnum.IDX_ROUND_NAME).Value = vRecordItem.Name
                .Cells(1, CompetionRoundsEnum.IDX_START_DATE_TIME).Value = vRecordItem.StartDateTime
                .Cells(1, CompetionRoundsEnum.IDX_ORDER_COUNT).Value = vRecordItem.OrderCount
                .Cells(1, CompetionRoundsEnum.IDX_GROUP_COUNT).Value = vRecordItem.GroupCount
                .Cells(1, CompetionRoundsEnum.IDX_GROUP_STRATEGY).Value = CodeMasterRepository.GroupStrategyCode.Item(vRecordItem.GroupStrategyType)
                .Cells(1, CompetionRoundsEnum.IDX_LANE_STRATEGY).Value = CodeMasterRepository.OrderStrategyCode.Item(vRecordItem.OrderStrategyType)
                .Cells(1, CompetionRoundsEnum.IDX_PLACE_EACH_GROUP).Value = vRecordItem.PlaceEachGroup
                .Cells(1, CompetionRoundsEnum.IDX_ADDITION_COUNT).Value = vRecordItem.AdditionCount
            End With
        End If
    Next vRecordRow 
End Sub

Private Function RangeRoundList() As Range
    Set RangeRoundList = ThisSheet _
            .Cells(COMPETITION_ROUNDS_ROW_HEADER, COMPETITION_ROUNDS_COLUMN_OFFSET) _
            .Offset(1, 1) _
            .Resize(CompetitionRepository.ReadSettinng().RoundsCount, 12)
End Function

Private Function ReadFromRowArray(RowArray() As Variant) As EventRoundModel
    With New EventRoundModel
        Call .Initialize( _
                RowArray(CompetionRoundsEnum.IDX_ID) _
                , RowArray(CompetionRoundsEnum.IDX_ROUND_NAME) _
                , RowArray(CompetionRoundsEnum.IDX_START_DATE_TIME) _
                , RowArray(CompetionRoundsEnum.IDX_ORDER_COUNT) _
                , RowArray(CompetionRoundsEnum.IDX_GROUP_COUNT) _
                , RowArray(CompetionRoundsEnum.IDX_GROUP_STRATEGY_CODE) _
                , RowArray(CompetionRoundsEnum.IDX_LANE_STRATEGY_CODE) _
                , RowArray(CompetionRoundsEnum.IDX_PLACE_EACH_GROUP) _
                , RowArray(CompetionRoundsEnum.IDX_ADDITION_COUNT) _
        )

        Set ReadFromRowArray = .Self()
    End With
End Function
