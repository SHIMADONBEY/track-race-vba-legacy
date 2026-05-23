Attribute VB_Name = "TrackOrderResultsReposotory"
'namespace=vba-files/track-race/repository
Option Explicit
Option Private Module

Private Enum TrackResultOrdersColumnEnum
    IDX_ID = 1
    IDX_ROUND_ID
    IDX_GROUP
    IDX_ORDER
    IDX_RANK
    IDX_REAL_RESULT
    IDX_RESULT
    IDX_REMARK
    IDX_REACTION_TIME
    IDX_NOT_STARTED
    IDX_NOT_FINISHED
    IDX_DISQUALIFIED_REASON
    IDX_ADVANCED_NEXT_ROUND
    IDX_DEMO_ENTRY
    IDX_TOTAL_RANK
    IDX_SAME_TIME
    IDX_RESULT_RANK
    IDX_NEXT_QUALIFIED
    IDX_SCORE
    IDX_UPDATED_AT
End Enum

Private Const BLANK_COUNT As Long = 3

Private m_Records As TrackResultOrderModels
Private m_RecordRows As Object
Private m_EndRow As Long

Public  Function ReadAllTrackResult() As TrackResultOrderModels
    Dim blankCount As Long: blankCount = BLANK_COUNT
    Dim resultRow As Range: Set resultRow = Range_TrackResultOrders_Header()
    Dim results As TrackResultOrderModels: Set results = New TrackResultOrderModels

    m_EndRow = 2
    Set m_RecordRows = CreateObject("Scripting.Dictionary")
    Do While blankCount > 0
        Set resultRow = resultRow.Offset(1, 0)
        If (resultRow.Cells(1, TrackResultOrdersColumnEnum.IDX_ID).Value = "") Then 
            blankCount = blankCount - 1
        Else
            blankCount = BLANK_COUNT
            m_EndRow = resultRow.Cells(1, 1).Row
            With New TrackResultOrderModel
                Call .Initialize( _
                    resultRow.Cells(1, TrackResultOrdersColumnEnum.IDX_ROUND_ID).value _
                    , resultRow.Cells(1, TrackResultOrdersColumnEnum.IDX_GROUP).value _
                    , resultRow.Cells(1, TrackResultOrdersColumnEnum.IDX_ORDER).value _
                    , resultRow.Cells(1, TrackResultOrdersColumnEnum.IDX_RANK).value _
                    , TimeProcessFunction.ParseTimeResultValue(resultRow.Cells(1, TrackResultOrdersColumnEnum.IDX_REAL_RESULT).value) _
                    , TimeProcessFunction.ParseTimeResultValue(resultRow.Cells(1, TrackResultOrdersColumnEnum.IDX_RESULT).value) _
                    , resultRow.Cells(1, TrackResultOrdersColumnEnum.IDX_REMARK).value _
                    , resultRow.Cells(1, TrackResultOrdersColumnEnum.IDX_REACTION_TIME).value _
                    , resultRow.Cells(1, TrackResultOrdersColumnEnum.IDX_NOT_STARTED).value _
                    , resultRow.Cells(1, TrackResultOrdersColumnEnum.IDX_NOT_FINISHED).value _
                    , resultRow.Cells(1, TrackResultOrdersColumnEnum.IDX_DISQUALIFIED_REASON).value _
                    , resultRow.Cells(1, TrackResultOrdersColumnEnum.IDX_ADVANCED_NEXT_ROUND).value _
                    , resultRow.Cells(1, TrackResultOrdersColumnEnum.IDX_DEMO_ENTRY).value _
                    , resultRow.Cells(1, TrackResultOrdersColumnEnum.IDX_SCORE).value _
                    , resultRow.Cells(1, TrackResultOrdersColumnEnum.IDX_UPDATED_AT).value _
                    , resultRow.Cells(1, TrackResultOrdersColumnEnum.IDX_ID).value _
                    , resultRow.Cells(1, TrackResultOrdersColumnEnum.IDX_TOTAL_RANK).value _
                    , resultRow.Cells(1, TrackResultOrdersColumnEnum.IDX_SAME_TIME).value _
                    , resultRow.Cells(1, TrackResultOrdersColumnEnum.IDX_RESULT_RANK).value _
                    , resultRow.Cells(1, TrackResultOrdersColumnEnum.IDX_NEXT_QUALIFIED).value _
                )
                Call results.Add(.Self())
                Call m_RecordRows.Add(.Id, resultRow)
            End With
        End If
    Loop

    Set m_Records = results
    Set ReadAllTrackResult = results
End Function

Public Function Upsert(Results As Collection)
    Dim result As TrackResultOrderModel
    Dim updatedCount As Long: updatedCount = 0
    Dim headerRow As Long: headerRow = Range_TrackResultOrders_Header().Cells(1, 1).Row

    For Each result In Results
        If (Not m_RecordRows.Exists(result.Id)) Then 
            m_EndRow = m_EndRow + 1
            Call m_RecordRows.Add(result.Id, Range_TrackResultOrders_Header().Offset(m_EndRow - headerRow, 0))
        End If

        updatedCount = updatedCount + WriteRow(result, m_RecordRows.Item(result.Id))
    Next result 

End Function

Private  Function WriteRow(Record As TrackResultOrderModel, TargetRow As Range)
    With TargetRow
        .Cells(1, TrackResultOrdersColumnEnum.IDX_ID).value = Record.Id
        .Cells(1, TrackResultOrdersColumnEnum.IDX_ROUND_ID).value = Record.RoundId
        .Cells(1, TrackResultOrdersColumnEnum.IDX_GROUP).value = Record.Group
        .Cells(1, TrackResultOrdersColumnEnum.IDX_ORDER).value = Record.Order
        .Cells(1, TrackResultOrdersColumnEnum.IDX_RANK).value = Record.Rank
        .Cells(1, TrackResultOrdersColumnEnum.IDX_REAL_RESULT).value = "'" & Record.RealResult.ToString()
        .Cells(1, TrackResultOrdersColumnEnum.IDX_RESULT).value = "'" & Record.Result.ToString()
        .Cells(1, TrackResultOrdersColumnEnum.IDX_REMARK).value = Record.Remark
        .Cells(1, TrackResultOrdersColumnEnum.IDX_REACTION_TIME).value = Record.ReactionTime
        .Cells(1, TrackResultOrdersColumnEnum.IDX_NOT_STARTED).value = Record.NotStarted
        .Cells(1, TrackResultOrdersColumnEnum.IDX_NOT_FINISHED).value = Record.NotFinished
        .Cells(1, TrackResultOrdersColumnEnum.IDX_DISQUALIFIED_REASON).value = Record.DisqualifiedReason
        .Cells(1, TrackResultOrdersColumnEnum.IDX_ADVANCED_NEXT_ROUND).value = Record.AdvancedNextRound
        .Cells(1, TrackResultOrdersColumnEnum.IDX_DEMO_ENTRY).value = Record.DemoEntry
        .Cells(1, TrackResultOrdersColumnEnum.IDX_SCORE).value = Record.Score
        .Cells(1, TrackResultOrdersColumnEnum.IDX_UPDATED_AT).value = Record.UpdatedAt
        .Cells(1, TrackResultOrdersColumnEnum.IDX_TOTAL_RANK).value = Record.TotalRank
        .Cells(1, TrackResultOrdersColumnEnum.IDX_SAME_TIME).value = Record.SameResult
        .Cells(1, TrackResultOrdersColumnEnum.IDX_RESULT_RANK).value = Record.ResultRank
        .Cells(1, TrackResultOrdersColumnEnum.IDX_NEXT_QUALIFIED).value = Record.NextQualified
    End With

    WriteRow = 1
End Function

Private  Function Get_ThisSheet() As Worksheet
    Set Get_ThisSheet = ResultListSheet
End Function

Private  Function Range_TrackResultOrders_Header() As Range
    Set Range_TrackResultOrders_Header = Get_ThisSheet().Cells(2, 1).Resize(1, 30)
End Function