Attribute VB_Name = "GameRecordRepository"
'namespace=vba-files/track-race/repository
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
    IDX_TEAM_COUNTRY
    IDX_RECORD_YEAR
    IDX_RECORD_MONTH
    IDX_RECORD_DAY
    IDX_TARGET_GRADE
    IDX_COMMENT
    IDX_TIME_UNIT
    IDX_EVENT_ID
    IDX_ID
End Enum

Private Const HEADER_ROW As Long = 2

Private m_Repository As GameRecordModels
Private m_RowRange As Object
Private m_EmptyRecordRows As Collection

Public Function ReadAllGameRecords(Optional Reload As Boolean = False) As GameRecordModels
    If (Reload Or m_Repository Is Nothing Or m_RowRange Is Nothing) Then
        Dim vRecords As GameRecordModels: Set vRecords = New GameRecordModels
        Set m_RowRange = CreateObject("Scripting.Dictionary")
        Set m_EmptyRecordRows = New Collection

        Dim vRowRange As Range
        For Each vRowRange In RangeGameRecords().Rows()
            Dim vRecord As GameRecordModel: Set vRecord = ReadFromRow(vRowRange)
            If (vRecord Is Nothing) Then
                If (vRowRange.Cells(1, GameRecordColumnEnum.IDX_ID).Value = "") Then
                    Call m_EmptyRecordRows.Add(vRowRange)
                End If
            Else
                Call vRecords.Add(vRecord)
                vRowRange.Cells(1, GameRecordColumnEnum.IDX_ID).Value = vRecord.Id
                Call m_RowRange.Add(vRecord.Id, vRowRange)
            End If
        Next vRowRange
        Set m_Repository = vRecords
    End If

    Set ReadAllGameRecords = m_Repository
End Function

Public Function Upsert(Records As Collection) As Long
    Dim vRecord As GameRecordModel
    Dim vUpdatedCount As Long: vUpdatedCount = 0

    For Each vRecord In Records
        Dim vRowRange As Range
        If (m_RowRange.Exists(vRecord.Id)) Then 
            Set vRowRange = m_RowRange.Item(vRecord.Id)
        Else
            Set vRowRange = RangeGameRecords().Offset(RangeGameRecords().Rows.Count, 0).Rows(1)
            If (m_EmptyRecordRows Is Nothing) Then
                ' DO NOTHING
            ElseIf (m_EmptyRecordRows.Count() > 0) Then
                ' ÉLÉÖÅ[Ç…ÇΩÇﬂÇƒÇ¢ÇÈãÛçsÇóDêÊÇ∑ÇÈ.
                Set vRowRange = m_EmptyRecordRows.Item(1)
                Call m_EmptyRecordRows.Remove(1)
            End If

            Call m_RowRange.Add(vRecord.Id, vRowRange)
        End If

        vUpdatedCount = vUpdatedCount + WriteRow(vRecord, vRowRange)
    Next vRecord 
End Function

Private Function ThisSheet() As Worksheet
    Set ThisSheet = GameRecordSheet
End Function

Private Function RangeGameRecords() As Range
    Set RangeGameRecords = GameRecordSheet.Range("GameRecordList")
End Function

Private Function ReadFromRow(RowRange As Range) As GameRecordModel
    Dim vResult As TimeResultValue  : Set vResult = TimeProcessFunction.ParseTimeResultValue(RowRange.Cells(1, GameRecordColumnEnum.IDX_RECORD).value)
    Dim vAthleteName As String      : vAthleteName = RowRange.Cells(1, GameRecordColumnEnum.IDX_ATHLETE_NAME).Value
    Dim vTeamName As String         : vTeamName = RowRange.Cells(1, GameRecordColumnEnum.IDX_TEAM_NAME).Value

    If (vAthleteName = "" Or vTeamName = "") Then
        Set ReadFromRow = Nothing
        Exit Function
    End If

    If (vResult.Valid = False) Then
        Set ReadFromRow = Nothing
        Exit Function
    End If

    Dim vRecordedYear As String     : vRecordedYear = RowRange.Cells(1, GameRecordColumnEnum.IDX_RECORD_YEAR).Value
    Dim vRecordedMonth As String    : vRecordedMonth = RowRange.Cells(1, GameRecordColumnEnum.IDX_RECORD_MONTH).Value
    Dim vRecordedDay As String      : vRecordedDay = RowRange.Cells(1, GameRecordColumnEnum.IDX_RECORD_DAY).Value

    Dim vRecordYmd As String
    vRecordYmd = IIf(vRecordedDay = "", "", "." & vRecordedDay)
    vRecordYmd = IIf(vRecordedMonth = "", "", "." & vRecordedMonth) & vRecordYmd
    vRecordYmd = vRecordedYear & vRecordYmd

    With New GameRecordModel
        Call .Initialize( _
                RowRange.Cells(1, GameRecordColumnEnum.IDX_EVENT_ID).Value _
                , RowRange.Row _
                , RowRange.Cells(1, GameRecordColumnEnum.IDX_NAME).Value _
                , RowRange.Cells(1, GameRecordColumnEnum.IDX_ABBREBIATION).Value _
                , vResult _
                , vAthleteName _
                , vTeamName _
                , RowRange.Cells(1, GameRecordColumnEnum.IDX_TEAM_PLACE).Value _
                , RowRange.Cells(1, GameRecordColumnEnum.IDX_TEAM_COUNTRY).Value _
                , vRecordedYear _
                , vRecordedMonth _
                , vRecordedDay _
                , RowRange.Cells(1, GameRecordColumnEnum.IDX_TIME_UNIT).Value _
                , RowRange.Cells(1, GameRecordColumnEnum.IDX_TARGET_GRADE).Value _
                , RowRange.Cells(1, GameRecordColumnEnum.IDX_COMMENT).Value _
                , vRecordYmd _
                , RowRange.Cells(1, GameRecordColumnEnum.IDX_ID).Value _
        )

        Set ReadFromRow = .Self()
    End With
End Function

Private Function WriteRow(Record As GameRecordModel, TargetRow As Range) As Long
    With TargetRow
        .Cells(1, GameRecordColumnEnum.IDX_NAME).Value          = Record.Name
        .Cells(1, GameRecordColumnEnum.IDX_ABBREBIATION).Value  = Record.Abbrebiation
        .Cells(1, GameRecordColumnEnum.IDX_RECORD).Value        = "'" & Record.Record.ToString()
        .Cells(1, GameRecordColumnEnum.IDX_TIME_SCALE).Value    = CodeMasterRepository.ResultUnitCode.Item(Record.RecordScale)
        .Cells(1, GameRecordColumnEnum.IDX_ATHLETE_NAME).Value  = Record.AthleteName
        .Cells(1, GameRecordColumnEnum.IDX_TEAM_NAME).Value     = Record.TeamName
        .Cells(1, GameRecordColumnEnum.IDX_TEAM_PLACE).Value    = Record.TeamPlace
        .Cells(1, GameRecordColumnEnum.IDX_TEAM_COUNTRY).Value  = Record.TeamCountry
        .Cells(1, GameRecordColumnEnum.IDX_RECORD_YEAR).Value   = Record.RecordedYear
        .Cells(1, GameRecordColumnEnum.IDX_RECORD_MONTH).Value  = Record.RecordedMonth
        .Cells(1, GameRecordColumnEnum.IDX_RECORD_DAY).Value    = Record.RecordedDay
        .Cells(1, GameRecordColumnEnum.IDX_TARGET_GRADE).Value  = Record.TargetGrade
        .Cells(1, GameRecordColumnEnum.IDX_COMMENT).Value       = Record.Comment
        .Cells(1, GameRecordColumnEnum.IDX_EVENT_ID).Value      = Record.EventId
        .Cells(1, GameRecordColumnEnum.IDX_ID).Value            = Record.Id
    End With

    WriteRow = 1
End Function
