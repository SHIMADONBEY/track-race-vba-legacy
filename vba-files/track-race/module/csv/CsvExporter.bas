Attribute VB_Name = "CsvExporter"
'namespace=vba-files/track-race/module/csv
Option Explicit

Public Sub ExportToCsv(FileStream As Object, Competition As CompetitionInfoModel, CompetitionEvent As CompetitionEventModel)
    Dim vResults As TrackResultOrderModels: Set vResults = TrackOrderResultsReposotory.ReadAllTrackResult()
    Call vResults.MergeStartList(ProgramEntryRepository.ReadAllStartList(True))
    Dim vSplitsGroup As Object: Set vSplitsGroup = CollectionUtil.GroupingBy(TrackSubResultOrderRepository.ReadAllRecords().All(), New TrackSubResultOrderClassifier)

    Call FileStream.WriteText(Header, 1)
    Dim vResult As TrackResultOrderModel
    For Each vResult In vResults.All()
        Dim vRound As EventRoundModel: Set vRound = EventRoundRepository.ReadAllRounds().Item(vResult.RoundId)
        Dim vGroup As TrackSubResultGroupModel: Set vGroup = TrackSubResultGroupRepository.ReadAllRecords().FindByGroup(vResult.RoundId, vResult.Group)

        Dim vSplits As Object: Set vSplits = CreateObject("Scripting.Dictionary")

        If (vSplitsGroup.Exists(vResult.Id)) Then 
            Dim vSplitData As TrackSubResultOrderModel
            For Each vSplitData In vSplitsGroup.Item(vResult.Id)
                If Not (vSplitData.SubNumber = 0) Then
                    Call vSplits.Add(vSplitData.Comment, vSplitData.SubResult)
                End If
            Next vSplitData 
        End If
        Call FileStream.WriteText(CsvRecord(Competition, CompetitionEvent, vRound, vResult, vGroup, vSplits), 1)
    Next vResult 
End Sub

Private Function Header()
    Dim vRowRange As Range
    For Each vRowRange In Range("Code_Csv_Header")
        Header = Header & "," & vRowRange.Value
    Next vRowRange 
    Header = Replace(Header, ",", "", , 1)
End Function

Private Function CsvRecord(Competition As CompetitionInfoModel, CompetitionEvent As CompetitionEventModel, Round As EventRoundModel, Result As TrackResultOrderModel, Group As TrackSubResultGroupModel, SplitTimes As Object)
    CsvRecord = CsvRecord & ","  & """" & Competition.CompetitionName & """"
    CsvRecord = CsvRecord & ","  & """" & Competition.CompetitionCode & """"
    CsvRecord = CsvRecord & ","  & """" & Competition.FacilityName & """"
    CsvRecord = CsvRecord & ","  & """" & Competition.FacilityCode & """"
    CsvRecord = CsvRecord & ","  & """" & Competition.FacilityPlace & """"
    CsvRecord = CsvRecord & ","  & """" & Competition.CompetitionYear & """"
    CsvRecord = CsvRecord & ","  & """" & Format(Competition.CompetitionDateStart, "mm.dd") & "-" & Format(Competition.CompetitionDateEnd, "mm.dd") & """"
    CsvRecord = CsvRecord & ","  & """" & Format(Round.StartDateTime, "mm.dd") & "*"""
    CsvRecord = CsvRecord & ","  & """" & Format(Round.StartDateTime, "h:nn") & """"
    CsvRecord = CsvRecord & ","  & """" & CompetitionEvent.Category & """"
    CsvRecord = CsvRecord & ","  & """" & CompetitionEvent.Sex & """"
    CsvRecord = CsvRecord & ","  & """" & CompetitionEvent.EventName & """"
    CsvRecord = CsvRecord & ","  & """" & CompetitionEvent.Supecification & """"
    CsvRecord = CsvRecord & ","  & """" & Round.Name & """"

    If (Round.NextQualifiersCount() > 0) Then
        CsvRecord = CsvRecord & ","  & """" & Round.GroupCount & "‘g" & Round.PlaceEachGroup & "’…" & IIf(Round.AdditionCount > 0, "+" & Round.AdditionCount, "") & """"
    Else
        CsvRecord = CsvRecord & ","  & """" & "" & """"
    End If

    CsvRecord = CsvRecord & ","  & """" & Result.Group & """"
    CsvRecord = CsvRecord & ","  & """" & Result.Order & """"

    If (Result.IsRankable()) Then
        CsvRecord = CsvRecord & ","  & """" & Result.Rank & """"
    Else
        CsvRecord = CsvRecord & ","
    End If
 
    If Not (Result.Person Is Nothing) Then
        CsvRecord = CsvRecord & ","  & """" & Result.Person.PersonalGrade & """"
        CsvRecord = CsvRecord & ","  & """" & Result.Person.Bib & """"
        CsvRecord = CsvRecord & ","  & """" & Result.Person.PersonalName & """"
        CsvRecord = CsvRecord & ","  & """" & Result.Person.PersonalCountry & """"
        CsvRecord = CsvRecord & ","  & """" & Result.Person.PersonalPhonetic & """"
        CsvRecord = CsvRecord & ","  & """" & Result.Person.TeamName & """"
        CsvRecord = CsvRecord & ","  & """" & Result.Person.PersonalLatin & """"
        CsvRecord = CsvRecord & ","  & """" & Result.Person.Age & """"
        CsvRecord = CsvRecord & ","  & """" & Result.Person.TeamPlace & """"
    Else
        Err.Raise CustomErrorCodeEnum.NoEntryAssigned, "CsvExporter.CsvRecord", MessageFactory.Generate("SE019").Prompt(Round.Name, Result.Group, Result.Order)
    End If

    CsvRecord = CsvRecord & ","  & """" & Result.Result.ToString() & "*"""
    CsvRecord = CsvRecord & ","  & """" & Result.Remark & """"

    If (CompetitionEvent.OperationWindGauge) Then 
        Dim vWind As WindValue: Set vWind = Group.Wind
        If (vWind Is Nothing) Then
            ' Œv‘ª‚¹‚¸.
            CsvRecord = CsvRecord & ",""NI"""
        Else
            CsvRecord = CsvRecord & ","  & """" & Format(Group.Wind.Value, "+0.0;-0.0; 0.0") & "*"""
        End If
    Else
        CsvRecord = CsvRecord & ","
    End If
    CsvRecord = CsvRecord & ","  & """" & Result.NextQualified & Result.AdvancedNextRound & """"
    CsvRecord = CsvRecord & ","  & """" & IIf(Result.NotStarted, "DNS", "") & IIf(Result.NotFinished, "DNF", "") & IIf(Result.Disqualified, "DQ " & Result.DisqualifiedReason, "") & """"
    CsvRecord = CsvRecord & ","  & """" & "" & """"

    If (CompetitionEvent.MeasurementReactionTime) Then 
        CsvRecord = CsvRecord & ","  & """" & Format(Result.ReactionTime, "0.000;") & "*"""
    Else
        CsvRecord = CsvRecord & ","
    End If

    If (SplitTimes.Count() > 0) Then 
        Dim checkPoint As Long
        For checkPoint = 200 To 9999 Step 200
            If (checkPoint Mod 400 = 0 Or checkPoint Mod 1000 = 0) Then
                If (SplitTimes.Exists(checkPoint & "m")) Then 
                    CsvRecord = CsvRecord & ","  & """" & SplitTimes.Item(checkPoint & "m").ToString() & "*"""
                Else
                    CsvRecord = CsvRecord & ","
                End If
            End If
        Next checkPoint
    Else
        CsvRecord = CsvRecord & ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
    End If

    CsvRecord = CsvRecord & ","  & """" & Result.RealResult.ToString() & "*"""
    CsvRecord = Replace(CsvRecord, ",", "", , 1)
End Function
