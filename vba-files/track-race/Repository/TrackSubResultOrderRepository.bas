Attribute VB_Name = "TrackSubResultOrderRepository"
'namespace=vba-files/track-race/repository
Option Explicit
Option Private Module

Private Enum TrackSubResultOrderColumnEnum
    IDX_ID = 1
    IDX_ROUND
    IDX_GROUP
    IDX_ORDER
    IDX_RESULT_ID
    IDX_SUB_NUMBER
    IDX_RANK
    IDX_RESULT
    IDX_WIND
    IDX_REACTION_TIME
    IDX_COMMENT
    IDX_UPDATED_AT
End Enum

Private Const BLANK_COUNT As Long = 3

Private m_Records As TrackSubResultOrderModels
Private m_RecordRows As Object
Private m_EndRow As Long

Public  Function ReadAllRecords() As TrackSubResultOrderModels
    Dim blankCount As Long: blankCount = BLANK_COUNT
    Dim resultRow As Range: Set resultRow = Range_TrackSubResultOrders_Header()
    Dim results As TrackSubResultOrderModels: Set results = New TrackSubResultOrderModels

    m_EndRow = 2
    Set m_RecordRows = CreateObject("Scripting.Dictionary")
    Do While blankCount > 0
        Set resultRow = resultRow.Offset(1, 0)
        If (resultRow.Cells(1, TrackSubResultOrderColumnEnum.IDX_ID).value = "") Then
            blankCount = blankCount - 1
        Else
            blankCount = BLANK_COUNT
            m_EndRow = resultRow.Cells(1, 1).Row
            With New TrackSubResultOrderModel
                Call .Initialize( _
                        resultRow.Cells(1, TrackSubResultOrderColumnEnum.IDX_ID).Value _
                        , resultRow.Cells(1, TrackSubResultOrderColumnEnum.IDX_ROUND).Value _
                        , resultRow.Cells(1, TrackSubResultOrderColumnEnum.IDX_GROUP).Value _
                        , resultRow.Cells(1, TrackSubResultOrderColumnEnum.IDX_ORDER).Value _
                        , resultRow.Cells(1, TrackSubResultOrderColumnEnum.IDX_SUB_NUMBER).Value _
                        , resultRow.Cells(1, TrackSubResultOrderColumnEnum.IDX_COMMENT).Value _
                        , TimeProcessFunction.ParseTimeResultValue(resultRow.Cells(1, TrackSubResultOrderColumnEnum.IDX_RESULT).Value) _
                        , resultRow.Cells(1, TrackSubResultOrderColumnEnum.IDX_RESULT_ID).Value _
                        , resultRow.Cells(1, TrackSubResultOrderColumnEnum.IDX_UPDATED_AT).Value _
                )
                Call results.Add(.Self())
                Call m_RecordRows.Add(.Id, resultRow)
            End With
        End If
    Loop

    Set m_Records = results
    Set ReadAllRecords = results
End Function

Public Function Upsert(Results As Collection)
    Dim result As TrackSubResultOrderModel
    Dim updatedCount As Long: updatedCount = 0
    Dim headerRow As Long: headerRow = Range_TrackSubResultOrders_Header().Cells(1, 1).Row

    For Each result In Results
        If (Not m_RecordRows.Exists(result.Id)) Then
            m_EndRow = m_EndRow + 1
            Call m_RecordRows.Add(result.Id, Range_TrackSubResultOrders_Header().Offset(m_EndRow - headerRow, 0))
        End If
        updatedCount = updatedCount + WriteRow(result, m_RecordRows.Item(result.Id))
    Next result

End Function

Private  Function WriteRow(Record As TrackSubResultOrderModel, TargetRow As Range) As Long
    With TargetRow
        .Cells(1, TrackSubResultOrderColumnEnum.IDX_ID).Value = Record.Id
        .Cells(1, TrackSubResultOrderColumnEnum.IDX_ROUND).Value = Record.RoundId
        .Cells(1, TrackSubResultOrderColumnEnum.IDX_GROUP).Value = Record.GroupNumber
        .Cells(1, TrackSubResultOrderColumnEnum.IDX_ORDER).Value = Record.Order
        .Cells(1, TrackSubResultOrderColumnEnum.IDX_SUB_NUMBER).Value = Record.SubNumber
        .Cells(1, TrackSubResultOrderColumnEnum.IDX_COMMENT).Value = Record.Comment
        .Cells(1, TrackSubResultOrderColumnEnum.IDX_RESULT).Value = "'" & Record.SubResult.ToString()
        .Cells(1, TrackSubResultOrderColumnEnum.IDX_RESULT_ID).Value = Record.ResultId
        .Cells(1, TrackSubResultOrderColumnEnum.IDX_UPDATED_AT).Value = Record.UpdatedAt
    End With

    WriteRow = 1
End Function

Private Function Get_ThisSheet() As Worksheet
    Set Get_ThisSheet = ResultSubListSheet
End Function

Private  Function Range_TrackSubResultOrders_Header() As Range
    Set Range_TrackSubResultOrders_Header = Get_ThisSheet().Cells(2, 21).Resize(1, 20)
End Function

