Attribute VB_Name = "QualifierDrawerController"
'namespace=vba-files/track-race/controller
Option Explicit
Option Private Module

Private Enum QualifierDrawerColumnEnum
    IDX_QUALIFY = 1
    IDX_GROUP
    IDX_RANK
    IDX_RESULT
    IDX_ORD
    IDX_PERSON_SEX
    IDX_PERSON_BIB
    IDX_PERSON_PERSONAL_NAME
    IDX_PERSON_AGE
    IDX_PERSON_PERSONAL_PHONETIC
    IDX_PERSON_PERSONAL_LATIN
    IDX_PERSON_TEAM_NAME
    IDX_PERSON_TEAM_PLACE
    IDX_PERSON_PERSONAL_COUNTRY
    IDX_PERSON_PERSONAL_GRADE
    IDX_ACCUMALTIVE
    IDX_PERSONAL_ID
    IDX_RESULT_ID
End Enum

Private r_LotsList As Range

Public Sub LoadQualifierLots(TargetRoundId As Long)
    Call Range_LotsList(True).Clear()
    Dim rangeRow As Range: Set rangeRow = Range_LotsList(True)

    Dim roundInfo As EventRoundModel: Set roundInfo = EventRoundRepository.ReadAllRounds().Item(TargetRoundId)
    Dim startList As StartListOrderModels: Set startList = ProgramEntryRepository.ReadAllStartList(True).FilterByRound(TargetRoundId)
    Dim results As TrackResultOrderModels: Set results = TrackOrderResultsReposotory.ReadAllTrackResult().FilterNeedToDraw(TargetRoundId)
    Dim record As TrackResultOrderModel
    For Each record In results.All()
        Dim startOrder As StartListOrderModel: Set startOrder = startList.FindOrder(TargetRoundId, record.Group, record.Order)

        rangeRow.Cells(1, QualifierDrawerColumnEnum.IDX_GROUP).Value = record.Group
        rangeRow.Cells(1, QualifierDrawerColumnEnum.IDX_RANK).Value = record.Rank
        rangeRow.Cells(1, QualifierDrawerColumnEnum.IDX_RESULT).Value = "'" & record.Result.ToString()
        rangeRow.Cells(1, QualifierDrawerColumnEnum.IDX_ORD).Value = record.Order
        rangeRow.Cells(1, QualifierDrawerColumnEnum.IDX_RESULT_ID).Value = record.Id
        rangeRow.Cells(1, QualifierDrawerColumnEnum.IDX_PERSONAL_ID).Value = startOrder.Person.Id
        rangeRow.Cells(1, QualifierDrawerColumnEnum.IDX_PERSON_SEX).Value = startOrder.Person.Sex
        rangeRow.Cells(1, QualifierDrawerColumnEnum.IDX_PERSON_BIB).Value = startOrder.Person.Bib
        rangeRow.Cells(1, QualifierDrawerColumnEnum.IDX_PERSON_PERSONAL_NAME).Value = startOrder.Person.PersonalName
        rangeRow.Cells(1, QualifierDrawerColumnEnum.IDX_PERSON_AGE).Value = startOrder.Person.Age
        rangeRow.Cells(1, QualifierDrawerColumnEnum.IDX_PERSON_PERSONAL_PHONETIC).Value = startOrder.Person.PersonalPhonetic
        rangeRow.Cells(1, QualifierDrawerColumnEnum.IDX_PERSON_PERSONAL_LATIN).Value = startOrder.Person.PersonalLatin
        rangeRow.Cells(1, QualifierDrawerColumnEnum.IDX_PERSON_TEAM_NAME).Value = startOrder.Person.TeamName
        rangeRow.Cells(1, QualifierDrawerColumnEnum.IDX_PERSON_TEAM_PLACE).Value = startOrder.Person.TeamPlace
        rangeRow.Cells(1, QualifierDrawerColumnEnum.IDX_PERSON_PERSONAL_COUNTRY).Value = startOrder.Person.PersonalCountry
        rangeRow.Cells(1, QualifierDrawerColumnEnum.IDX_PERSON_PERSONAL_GRADE).Value = startOrder.Person.PersonalGrade
        rangeRow.Cells(1, QualifierDrawerColumnEnum.IDX_ACCUMALTIVE).Value = StringUtil.JoinCollectionToString(startOrder.Accumaltive, ",")

        Set rangeRow = rangeRow.Offset(1, 0)
    Next record 

    With Range_LotsList(true)
        .Columns(QualifierDrawerColumnEnum.IDX_RESULT).HorizontalAlignment = xlRight

        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With

        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlHairLine
            .ColorIndex = xlAutomatic
        End With

        Range_CurrentTargetNumber().FormulaR1C1 = "=CountA(" & .Columns(QualifierDrawerColumnEnum.IDX_QUALIFY).Address(ReferenceStyle:=xlR1C1) & ")"
    End With

    If (results.Count < 1) Then
        Get_ThisSheet().Visible = 2
        Exit Sub
    Else
        Get_ThisSheet().Visible = -1
        Get_ThisSheet().Activate
        Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
        Call MessageFactory.Generate("SI007").ToMsgBox()
        Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
    End If

    Range_LotsRoundName().Value = roundInfo.Name
    Range_LotsRoundId().Value = roundInfo.Id
    Range_TargetNumber().Value = roundInfo.GroupCount * roundInfo.PlaceEachGroup + roundInfo.AdditionCount - TrackOrderResultsReposotory.ReadAllTrackResult().FilterQualified(TargetRoundId).Count()
End Sub

Public Sub SelectQualifiers(TargetRoundId As Long, TargetNumber As Long)
    Dim currentRoundInfo As EventRoundModel: Set currentRoundInfo = EventRoundRepository.ReadAllRounds().Item(TargetRoundId)
    Dim lots As TrackResultOrderModels: Set lots = TrackOrderResultsReposotory.ReadAllTrackResult().FilterNeedToDraw(TargetRoundId)
    Dim updateLotList As Collection: Set updateLotList = New Collection
    Dim lot As TrackResultOrderModel
    Dim rowRange As Range

    For Each rowRange In Range_LotsList().Rows()
        Set lot = lots.FindByOrder(TargetRoundId, rowRange.Cells(1, IDX_GROUP).Value, rowRange.Cells(1, IDX_ORD).Value)
        If (IsEmpty(rowRange.Cells(1, IDX_QUALIFY).Value)) Then 
            lot.NextQualified = ""
            lot.ResultRank = lot.ResultRank + TargetNumber
        Else
            lot.NextQualified = IIf(lot.Rank <= currentRoundInfo.PlaceEachGroup, "Q", "q")
        End If
        Call updateLotList.Add(lot)
    Next rowRange

    Call TrackOrderResultsReposotory.Upsert(updateLotList)
    Call LoadQualifierLots(TargetRoundId)
    Call ResultRankingWritingService.WriteRanking(TargetRoundId)
End Sub

Public Sub QualifyAll(TargetRoundId As Long)
    Dim currentRoundInfo As EventRoundModel: Set currentRoundInfo = EventRoundRepository.ReadAllRounds().Item(TargetRoundId)
    Dim lots As TrackResultOrderModels: Set lots = TrackOrderResultsReposotory.ReadAllTrackResult().FilterNeedToDraw(TargetRoundId)
    Dim lot As TrackResultOrderModel
    Dim updateLotList As Collection: Set updateLotList = New Collection

    For Each lot In lots.All()
        lot.NextQualified = IIf(lot.Rank <= currentRoundInfo.PlaceEachGroup, "Q", "q")
        Call updateLotList.Add(lot)
    Next lot

    Call TrackOrderResultsReposotory.Upsert(updateLotList)
    Call LoadQualifierLots(TargetRoundId)
    Call ResultRankingWritingService.WriteRanking(TargetRoundId)
End Sub

Private Function Get_ThisSheet() As Worksheet
    Set Get_ThisSheet = QualifierRotterySheet
End Function

Private Function Range_LotsList(Optional Reload As Boolean = False) As Range
    If (Not(r_LotsList Is Nothing) And Not Reload) Then
        Set Range_LotsList = r_LotsList
    End If

    Dim rowCount As Long: rowCount = Application.WorksheetFunction.CountA(Get_ThisSheet().Columns(QualifierDrawerColumnEnum.IDX_RESULT_ID + 1))
    rowCount = IIf(rowCount > 1, rowCount, 1)
    Set r_LotsList = Get_ThisSheet().Cells(10, 2).Offset(1, 0).Resize(rowCount - IIf(rowCount > 1, 1, 0), 30)
    Set Range_LotsList = r_LotsList
End Function

Private Function Range_LotsRoundName() As Range
    Set Range_LotsRoundName = Get_ThisSheet().Cells(2, 5)
End Function

Private Function Range_LotsRoundId() As Range
    Set Range_LotsRoundId = Get_ThisSheet().Cells(2, 6)
End Function

Private Function Range_TargetNumber() As Range
    Set Range_TargetNumber = Get_ThisSheet().Cells(3, 5)
End Function

Private Function Range_CurrentTargetNumber() As Range
    Set Range_CurrentTargetNumber = Get_ThisSheet().Cells(9, 2)
End Function
