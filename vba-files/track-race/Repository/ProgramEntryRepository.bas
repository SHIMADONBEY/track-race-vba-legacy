Attribute VB_Name = "ProgramEntryRepository"
'namespace=vba-files/track-race/repository
Option Explicit
Option Private Module

Public Enum EntryListColumnIndex
    IDX_ROUND = 1
    IDX_PERSONAL_ID
    IDX_PERSONAL_CODE
    IDX_PERSON_SEX
    IDX_PERSON_BIB
    IDX_PERSON_PERSONAL_NAME
    IDX_PERSON_AGE
    IDX_PERSON_PERSONAL_PHONETIC
    IDX_PERSON_PERSONAL_LATIN
    IDX_PERSON_TEAM_NAME
    IDX_PERSON_TEAM_PLACE
    IDX_PERSON_PERSONAL_COUNTRY
    IDX_PERSON_DATE_OF_BIRTH
    IDX_PERSON_PERSONAL_GRADE
    IDX_PERSON_TEAM_ID
    IDX_QUALIFIED_1
    IDX_QUALIFIED_2
    IDX_QUALIFIED_3
    IDX_QUALIFIED_RANK
    IDX_DEMO_ENTRY
    IDX_GROUP
    IDX_ORDER
    IDX_ACCUMALTIVE
    IDX_NOT_STARTED
    IDX_START_ID
End Enum

Private m_RowRange As Object
Private m_Records As StartListOrderModels

Public Function ReadAllStartList(Optional Reload As Boolean) As StartListOrderModels
    If ((m_Records Is Nothing) Or (m_RowRange Is Nothing) Or (Reload = True)) Then
        Call ReadStartList()
    End If

    Set ReadAllStartList = New StartListOrderModels
    Call ReadAllStartList.AddRange(m_Records.All())
End Function

Public Function UpdateOrder(Records As Collection) As Long
    Dim vRecord As StartListOrderModel
    Dim updatedCount As Long: updatedCount = 0
    
    For Each vRecord In Records
        If (m_RowRange.Exists(vRecord.Id)) Then
            With m_RowRange.Item(vRecord.Id)
                .Cells(1, EntryListColumnIndex.IDX_GROUP).Value = vRecord.Group
                .Cells(1, EntryListColumnIndex.IDX_ORDER).Value = vRecord.Order
            End With

            updatedCount = updatedCount + 1
        End If
    Next vRecord
    UpdateOrder = updatedCount
End Function

Public Function Upsert(Records As Collection)
    Dim vRecord As StartListOrderModel
    Dim updatedCount As Long: updatedCount = 0
    Dim vRowRange As Range
    For Each vRecord In Records
        If (m_RowRange.Exists(vRecord.Id)) Then 
            Set vRowRange = m_RowRange.Item(vRecord.Id)
        Else
            Set vRowRange = Range_EntryList().Offset(Range_EntryList().Rows.Count, 0).Rows(1)
            Call m_RowRange.Add(vRecord.Id, vRowRange)
        End If

        updatedCount = updatedCount + WriteRow(vRecord, vRowRange)
    Next vRecord
    Upsert = updatedCount
End Function

Private Function WriteRow(Record As StartListOrderModel, TargetRow As Range) As Long
    With TargetRow
        .Cells(1, EntryListColumnIndex.IDX_ROUND).Value = EventRoundRepository.ReadAllRounds().Item(Record.RoundId).Name
        .Cells(1, EntryListColumnIndex.IDX_PERSONAL_ID).Value = Record.Person.Id
        .Cells(1, EntryListColumnIndex.IDX_PERSONAL_CODE).Value = Record.Person.PersonalCode
        .Cells(1, EntryListColumnIndex.IDX_PERSON_SEX).Value = Record.Person.Sex
        .Cells(1, EntryListColumnIndex.IDX_PERSON_BIB).Value = Record.Person.Bib
        .Cells(1, EntryListColumnIndex.IDX_PERSON_PERSONAL_NAME).Value = Record.Person.PersonalName
        .Cells(1, EntryListColumnIndex.IDX_PERSON_AGE).Value = Record.Person.Age
        .Cells(1, EntryListColumnIndex.IDX_PERSON_PERSONAL_PHONETIC).Value = Record.Person.PersonalPhonetic
        .Cells(1, EntryListColumnIndex.IDX_PERSON_PERSONAL_LATIN).Value = Record.Person.PersonalLatin
        .Cells(1, EntryListColumnIndex.IDX_PERSON_TEAM_NAME).Value = Record.Person.TeamName
        .Cells(1, EntryListColumnIndex.IDX_PERSON_TEAM_PLACE).Value = Record.Person.TeamPlace
        .Cells(1, EntryListColumnIndex.IDX_PERSON_PERSONAL_COUNTRY).Value = Record.Person.PersonalCountry

        Dim vBirthDate As Date: vBirthDate = Record.Person.BirthDate
        If (IsEmpty(vBirthDate) = False And vBirthDate >= 1) Then 
            .Cells(1, EntryListColumnIndex.IDX_PERSON_DATE_OF_BIRTH).Value = Record.Person.BirthDate
        Else
            .Cells(1, EntryListColumnIndex.IDX_PERSON_DATE_OF_BIRTH).Value = ""
        End If

        .Cells(1, EntryListColumnIndex.IDX_PERSON_PERSONAL_GRADE).Value = Record.Person.PersonalGrade
        .Cells(1, EntryListColumnIndex.IDX_PERSON_TEAM_ID).Value = Record.Person.TeamId
        .Cells(1, EntryListColumnIndex.IDX_QUALIFIED_1).Value = IIf(Record.Qualified1.Valid, Record.Qualified1.ToString(), "")
        .Cells(1, EntryListColumnIndex.IDX_QUALIFIED_2).Value = IIf(Record.Qualified2.Valid, Record.Qualified2.ToString(), "")
        .Cells(1, EntryListColumnIndex.IDX_QUALIFIED_3).Value = IIf(Record.Qualified3.Valid, Record.Qualified3.ToString(), "")
        .Cells(1, EntryListColumnIndex.IDX_QUALIFIED_RANK).Value = Record.QualifiedRank
        .Cells(1, EntryListColumnIndex.IDX_DEMO_ENTRY).Value = IIf(Record.DemoEntry, "o", "")
        .Cells(1, EntryListColumnIndex.IDX_GROUP).Value = Record.Group
        .Cells(1, EntryListColumnIndex.IDX_ORDER).Value = Record.Order
        .Cells(1, EntryListColumnIndex.IDX_ACCUMALTIVE).Value = StringUtil.JoinCollectionToString(Record.Accumaltive, ", ")
        .Cells(1, EntryListColumnIndex.IDX_NOT_STARTED).Value = IIf(Record.NotStarted, "DNS", "")
        .Cells(1, EntryListColumnIndex.IDX_START_ID).Value = Record.Id
    End With
    WriteRow = 1
End Function

Private Function Range_EntryList() As Range
    Set Range_EntryList = EntryListSheet.Range("EntryList")
End Function

Private Sub ReadStartList()
    Set m_RowRange = CreateObject("Scripting.Dictionary")
    Set m_Records = New StartListOrderModels
    Dim vRowRange As Range
    For Each vRowRange In Range_EntryList().Rows
        Dim vRound As EventRoundModel: Set vRound = EventRoundRepository.ReadAllRounds().FindByName(vRowRange.Cells(1, EntryListColumnIndex.IDX_ROUND).Value)
        If (vRound Is Nothing) Then 
            ' ラウンドデータがないためスキップ
            ' DO NOTHING
        Else
            Dim vRecordId As String: vRecordId = vRowRange.Cells(1, EntryListColumnIndex.IDX_START_ID).Value
            Dim vPerson As AthletePersonModel: Set vPerson = ReadPersonFromRow(vRowRange)
            Dim vEntry As PersonalEntryModel: Set vEntry = ReadEntryFromRow(vRowRange)
            Dim vQRank As Long: vQRank = vRowRange.Cells(1, EntryListColumnIndex.IDX_QUALIFIED_RANK).Value
            Dim vAccumaltiveArray() As String: vAccumaltiveArray = Split(vRowRange.Cells(1, EntryListColumnIndex.IDX_ACCUMALTIVE).Value, ",")
    
            Dim vRoundId As Long: vRoundId = vRound.Id
            Dim vGroup As Long: vGroup = vRowRange.Cells(1, EntryListColumnIndex.IDX_GROUP).Value
            Dim vOrder As Long: vOrder = vRowRange.Cells(1, EntryListColumnIndex.IDX_ORDER).Value
    
            With New StartListOrderModel
                Call .Initialize( _
                        vPerson _
                        , vEntry _
                        , vQRank _
                        , vRoundId _
                        , vAccumaltiveArray _
                        , Not (vRowRange.Cells(1, EntryListColumnIndex.IDX_NOT_STARTED).Value = "") _
                        , vGroup _
                        , vOrder _
                        , vRecordId _
                )
    
                Call m_Records.Add(.Self())
    
                If (vRecordId = "") Then
                    vRecordId = .Id
                    vRowRange.Cells(1, EntryListColumnIndex.IDX_START_ID).Value = vRecordId
                End If
            End With
    
            Call m_RowRange.Add(vRecordId, vRowRange)
        End If
    Next vRowRange 
End Sub

Private Function ReadPersonFromRow(RowRange As Range) As AthletePersonModel
    With New AthletePersonModel
        Call .Initialize( _
                RowRange.Cells(1, EntryListColumnIndex.IDX_PERSON_SEX).Value _
                , RowRange.Cells(1, EntryListColumnIndex.IDX_PERSON_BIB).Value _
                , RowRange.Cells(1, EntryListColumnIndex.IDX_PERSON_PERSONAL_NAME).Value _
                , RowRange.Cells(1, EntryListColumnIndex.IDX_PERSON_AGE).Value _
                , RowRange.Cells(1, EntryListColumnIndex.IDX_PERSON_TEAM_NAME).Value _
                , RowRange.Cells(1, EntryListColumnIndex.IDX_PERSON_PERSONAL_PHONETIC).Value _
                , RowRange.Cells(1, EntryListColumnIndex.IDX_PERSON_PERSONAL_LATIN).Value _
                , RowRange.Cells(1, EntryListColumnIndex.IDX_PERSON_TEAM_PLACE).Value _
                , RowRange.Cells(1, EntryListColumnIndex.IDX_PERSON_PERSONAL_COUNTRY).Value _
                , RowRange.Cells(1, EntryListColumnIndex.IDX_PERSON_PERSONAL_GRADE).Value _
                , RowRange.Cells(1, EntryListColumnIndex.IDX_PERSONAL_CODE).Value _
                , RowRange.Cells(1, EntryListColumnIndex.IDX_PERSON_DATE_OF_BIRTH).Value _
                , RowRange.Cells(1, EntryListColumnIndex.IDX_PERSON_TEAM_ID).Value _
                , RowRange.Cells(1, EntryListColumnIndex.IDX_PERSONAL_ID).Value _
        )

        Set ReadPersonFromRow = .Self()
    End With
End Function

Private Function ReadEntryFromRow(RowRange As Range) As PersonalEntryModel
    With New PersonalEntryModel
        Call .Initialize( _
                RowRange.Cells(1, EntryListColumnIndex.IDX_PERSONAL_ID).Value _
                , 0 _
                , 0 _
                , Not (RowRange.Cells(1, EntryListColumnIndex.IDX_DEMO_ENTRY).Value = "") _
                , TimeProcessFunction.ParseTimeResultValue(RowRange.Cells(1, EntryListColumnIndex.IDX_QUALIFIED_1).Value) _
                , TimeProcessFunction.ParseTimeResultValue(RowRange.Cells(1, EntryListColumnIndex.IDX_QUALIFIED_2).Value) _
                , TimeProcessFunction.ParseTimeResultValue(RowRange.Cells(1, EntryListColumnIndex.IDX_QUALIFIED_3).Value) _
        )

        Set ReadEntryFromRow = .Self()
    End With
End Function
