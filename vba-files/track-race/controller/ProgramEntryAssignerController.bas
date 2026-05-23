Attribute VB_Name = "ProgramEntryAssignerController"
'namespace=vba-files/track-race/controller
Option Explicit
Option Private Module

Private Enum EntryListColumnIndex
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
End Enum

Public Sub ProgramEntries(ByVal TargetRoundName As String)
    Randomize
    Dim pRound As EventRoundModel: Set pRound = EventRoundRepository.ReadAllRounds().FindByName(TargetRoundName)
    If (pRound Is Nothing) Then
        Err.Raise CustomErrorCodeEnum.TargetRoundNotSelected, "ProgramEntryAssignerController.ProgramEntries", MessageFactory.Generate("SE022").Prompt
    End If

    Dim vEntryRecords As StartListOrderModels: Set vEntryRecords = ProgramEntryRepository.ReadAllStartList(true).FilterByRound(pRound.Id)
    Dim vEntryRecord As StartListOrderModel
    Dim vAthletes As QualifiedAthleteModels: Set vAthletes = New QualifiedAthleteModels

    Dim vResults As TrackResultOrderModels: Set vResults = TrackOrderResultsReposotory.ReadAllTrackResult().FilterByRound(pRound.Id)
    If (vResults.Count() > 0) Then
        Err.Raise CustomErrorCodeEnum.ResultAlreadyRegistred, "ProgramEntryAssignerController.ProgramEntries", MessageFactory.Generate("SE018").Prompt
    End If

    For Each vEntryRecord In vEntryRecords.All()
        With New QualifiedAthleteModel
            Call .Initialize(vEntryRecord.Id, vEntryRecord.Person, vEntryRecord.Entry, vEntryRecord.QualifiedRank)
            Call vAthletes.Add(.Self())
        End With
    Next vEntryRecord

    If (vAthletes.IsEmpty()) Then
        Exit Sub
    End If

    Dim vOrganizer As ProgramEntryOrganizer: Set vOrganizer = New ProgramEntryOrganizer
    Call vOrganizer.Initialize(GroupAssignerFactory.Generate(pRound.GroupStrategyType), OrderAssignerFactory.Generate(pRound.OrderStrategyType))

    Dim vProgrammedAthletes As QualifiedAthleteModels: Set vProgrammedAthletes = vOrganizer.Organize(vAthletes, CompetitionRepository.ReadSettinng().PersonPerGroup)

    Dim vProgrammedRecords As StartListOrderModels: Set vProgrammedRecords = New StartListOrderModels
    Dim vAthlete As QualifiedAthleteModel
    For Each vAthlete In vProgrammedAthletes.All
        With vEntryRecords.Item(vAthlete.Id)
            .Group = vAthlete.Group
            .Order = vAthlete.Order
            Call vProgrammedRecords.Add(.Self())
        End With
    Next vAthlete 

    Call ProgramEntryRepository.UpdateOrder(vProgrammedRecords.All)
End Sub

Public Sub WriteStartList(ByVal TargetRoundName As String)
    Dim pRound As EventRoundModel: Set pRound = EventRoundRepository.ReadAllRounds().FindByName(TargetRoundName)
    Call HtmlLoader.OpenHtmlFile(StartListWritingService.WriteStartList(pRound))
End Sub

Public Sub PickUpQualifiers(ByVal TargetRoundName As String, Optional ByVal TopCount As Long = -1)
    Dim vTargetRound As EventRoundModel: Set vTargetRound = EventRoundRepository.ReadAllRounds().FindByName(TargetRoundName)

    If (vTargetRound.NextQualifiersCount() < 1) Then
        Err.Raise CustomErrorCodeEnum.NoQualifiers, "ProgramEntryAssignerController.PickUpQualifiers", MessageFactory.Generate("SI020").Prompt
    End If

    Dim vQualifiersStartList As StartListOrderModels: Set vQualifiersStartList = ProgramEntryAssignService.PickUpQualifiersByResults(vTargetRound.Id, TopCount)
    If (vQualifiersStartList.Count() < 1) Then
        Err.Raise CustomErrorCodeEnum.QualifiersNotAssigned, "ProgramEntryAssignerController.PickUpQualifiers", MessageFactory.Generate("SW012").Prompt
    End If

    Call ProgramEntryRepository.Upsert(vQualifiersStartList.All())
End Sub

Private Function Range_EntryList() As Range
    Set Range_EntryList = EntryListSheet.Range("EntryList")
End Function
