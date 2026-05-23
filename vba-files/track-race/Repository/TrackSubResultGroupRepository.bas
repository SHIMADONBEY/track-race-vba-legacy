Attribute VB_Name = "TrackSubResultGroupRepository"
'namespace=vba-files/track-race/repository
Option Explicit
Option Private Module

Private Enum TrackSubResultGroupColumnEnum
    IDX_ID = 1
    IDX_ROUND
    IDX_GROUP
    IDX_WIND
    IDX_UPDATED_AT
End Enum

Private Const BLANK_COUNT As Long = 3

Private m_Records As TrackSubResultGroupModels
Private m_RecordRows As Object
Private m_EndRow As Long

Public Function ReadAllRecords() As TrackSubResultGroupModels
    Dim blankCount As Long: blankCount = BLANK_COUNT
    Dim resultRow As Range: Set resultRow = Range_TrackSubResultGroups_Header()
    Dim results As TrackSubResultGroupModels: Set results = New TrackSubResultGroupModels
    Dim windText As String

    m_EndRow = resultRow.Cells(1, 1).Row
    Set m_RecordRows = CreateObject("Scripting.Dictionary")
    Do While blankCount > 0
        Set resultRow = resultRow.Offset(1, 0)
        If (resultRow.Cells(1, TrackSubResultGroupColumnEnum.IDX_ID) = "") Then 
            blankCount = blankCount - 1
        Else
            blankCount = BLANK_COUNT
            m_EndRow = resultRow.Cells(1, 1).Row
            windText = IIf(CompetitionRepository.ReadSettinng().OperationWindGauge, resultRow.Cells(1, TrackSubResultGroupColumnEnum.IDX_WIND).Value, "")
            With New TrackSubResultGroupModel
                Call .Initialize( _
                        resultRow.Cells(1, TrackSubResultGroupColumnEnum.IDX_ROUND).Value _
                        , resultRow.Cells(1, TrackSubResultGroupColumnEnum.IDX_GROUP).Value _
                        , IIf(windText = "", Nothing, WindProcessFunction.ConvertWind(windText)) _
                        , resultRow.Cells(1, TrackSubResultGroupColumnEnum.IDX_UPDATED_AT).Value _
                        , resultRow.Cells(1, TrackSubResultGroupColumnEnum.IDX_ID).Value _
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
    Dim result As TrackSubResultGroupModel
    Dim updateCount As Long: updateCount = 0
    Dim headerRow As Long: headerRow = Range_TrackSubResultGroups_Header().Cells(1, 1).Row

    For Each result In Results
        If Not (m_RecordRows.Exists(result.Id)) Then
            m_EndRow = m_EndRow + 1
            Call m_RecordRows.Add(result.Id, Range_TrackSubResultGroups_Header().Offset(m_EndRow - headerRow, 0))
        End If
        updateCount = updateCount + WriteRow(result,m_RecordRows.Item(result.Id))
    Next result 

    Upsert = updateCount
End Function

Private Function WriteRow(Record As TrackSubResultGroupModel, TargetRow As Range) As Long
    With TargetRow
        .Cells(1, TrackSubResultGroupColumnEnum.IDX_ID).Value = Record.Id
        .Cells(1, TrackSubResultGroupColumnEnum.IDX_ROUND).Value = Record.RoundId
        .Cells(1, TrackSubResultGroupColumnEnum.IDX_GROUP).Value = Record.Group

        If (Record.Wind Is Nothing) Then
            .Cells(1, TrackSubResultGroupColumnEnum.IDX_WIND).Value = Empty
        Else
            .Cells(1, TrackSubResultGroupColumnEnum.IDX_WIND).Value = Record.Wind.Value
        End If

        .Cells(1, TrackSubResultGroupColumnEnum.IDX_UPDATED_AT).Value = Record.UpdatedAt
    End With

    WriteRow = 1
End Function

Private Function Get_ThisSheet() As Worksheet
    Set Get_ThisSheet = ResultSubListSheet
End Function

Private  Function Range_TrackSubResultGroups_Header() As Range
    Set Range_TrackSubResultGroups_Header = Get_ThisSheet().Cells(2, 1).Resize(1, 20)
End Function

