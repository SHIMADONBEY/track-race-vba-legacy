Attribute VB_Name = "GameRecordController"
'namespace=vba-files/track-race/controller
Option Explicit
Option Private Module

Private Enum GameRecordColumnEnum
    IDX_NAME = 1
    IDX_ABBREBIATION
    IDX_RECORD
    IDX_TIME_SCALE
    IDX_ATHLETE_NAME
    IDX_TEAM_NAME
    IDX_TEAM_PLACE
    IDX_RECORD_YEAR
    IDX_RECORD_MONTH
    IDX_RECORD_DAY
    IDX_TIME_UNIT
    IDX_ID
End Enum

Private Const HEADER_ROW As Long = 2

Public Function OnUpdatedRecord(Spot As Range) As Boolean
    If (Application.Intersect(Spot, Range("GameRecordList").Columns(GameRecordColumnEnum.IDX_RECORD)) Is Nothing) Then
        OnUpdatedRecord = False
        Exit Function
    End If

    Dim vCurrentRow As Range: Set vCurrentRow = Range("GameRecordList").Rows(Spot.Row - HEADER_ROW)
    Dim vPrecision As Long: vPrecision = CompetitionRepository.ReadSettinng().ResultPrecision
    If (vCurrentRow.Cells(1, GameRecordColumnEnum.IDX_TIME_SCALE).Value = "") Then
        vCurrentRow.Cells(1, GameRecordColumnEnum.IDX_TIME_SCALE).Value = CodeMasterRepository.ResultUnitCode.Item(vPrecision)
    Else
        vPrecision = vCurrentRow.Cells(1, GameRecordColumnEnum.IDX_TIME_UNIT).Value
    End If

    Dim vRecordResult As TimeResultValue: Set vRecordResult = TimeProcessFunction.ParseFromNumberString(Spot.Value, vPrecision)
    If Not (vRecordResult Is Nothing) Then
        Spot.Value = "'" & vRecordResult.ToString()
    End If

    OnUpdatedRecord = True
End Function

Public Sub UpdateGameRecords()
    Call GameRecordRepository.ReadAllGameRecords(True)
End Sub
