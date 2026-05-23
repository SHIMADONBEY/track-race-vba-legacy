Attribute VB_Name = "StartListWritingService"
'namespace=vba-files/track-race/modules/html/output/start-list
Option Explicit
Option Private Module

Public Function WriteStartList(TargetRound As EventRoundModel)
    Dim pRoundId As Long: pRoundId = TargetRound.Id
    Dim vAthletes As StartListRoundModel: Set vAthletes = New StartListRoundModel
    Dim i As Long, j As Long

    Call vAthletes.Initialize(TargetRound.GroupCount, pRoundId, ProgramEntryRepository.ReadAllStartList(true).FilterByRound(pRoundId).SortByOrder())
    If (vAthletes.EmptyOrders()) Then
        ' ëgÇï‚è[Ç∑ÇÈ.
        Dim emptyAccmaltiveArray() As String: emptyAccmaltiveArray = Split("", ",")
        For i = 1 To TargetRound.GroupCount()
            Dim vOrderNumbers As Variant
            Dim vLaneSettings As LaneSettingModel: Set vLaneSettings = CompetitionRepository.ReadSettinng().LaneSetting

            If CompetitionRepository.ReadSettinng().PersonPerGroup <= vLaneSettings.LaneCount() Then
                vOrderNumbers = vLaneSettings.TopLanes
                For j = LBound(vLaneSettings.TopLanes) To UBound(vLaneSettings.TopLanes)
                    With New StartListOrderModel
                        Call .Initialize( _
                                GenerateEmptyAthlete() _
                                , GenerateEmptyEntry() _
                                , 2147483647 _
                                , pRoundId _
                                , emptyAccmaltiveArray _
                                , false _
                                , i _
                                , CLng(vOrderNumbers(j)) _
                        )

                        Call vAthletes.AddOrder(.Self())
                    End With
                Next j
    
                vOrderNumbers = vLaneSettings.MiddleLanes
                For j = LBound(vLaneSettings.MiddleLanes) To UBound(vLaneSettings.MiddleLanes)
                    With New StartListOrderModel
                        Call .Initialize( _
                                GenerateEmptyAthlete() _
                                , GenerateEmptyEntry() _
                                , 2147483647 _
                                , pRoundId _
                                , emptyAccmaltiveArray _
                                , false _
                                , i _
                                , CLng(vOrderNumbers(j)) _
                        )

                        Call vAthletes.AddOrder(.Self())
                    End With
                Next j

                vOrderNumbers = vLaneSettings.BottomLanes
                For j = LBound(vLaneSettings.BottomLanes) To UBound(vLaneSettings.BottomLanes)
                    With New StartListOrderModel
                        Call .Initialize( _
                                GenerateEmptyAthlete() _
                                , GenerateEmptyEntry() _
                                , 2147483647 _
                                , pRoundId _
                                , emptyAccmaltiveArray _
                                , false _
                                , i _
                                , CLng(vOrderNumbers(j)) _
                        )

                        Call vAthletes.AddOrder(.Self())
                    End With
                Next j
            Else
                Dim groupPersons As Long: groupPersons = Fix((TargetRound.OrderCount + (TargetRound.GroupCount - 1)) / TargetRound.GroupCount) 
                For j = 1 To groupPersons
                    With New StartListOrderModel
                        Call .Initialize( _
                                GenerateEmptyAthlete() _
                                , GenerateEmptyEntry() _
                                , 2147483647 _
                                , pRoundId _
                                , emptyAccmaltiveArray _
                                , false _
                                , i _
                                , j _
                        )

                        Call vAthletes.AddOrder(.Self())
                    End With
                Next j
            End If
        Next i
    End If

    Dim pRoundName As String: pRoundName = TargetRound.Name
    Dim vStartListData As StartListOutputModel: Set vStartListData = New StartListOutputModel
    Call vStartListData.Initialize( _
            pRoundName _
            , CompetitionRepository.ReadSettinng() _
            , EventRoundRepository.ReadAllRounds() _
            , GameRecordRepository.ReadAllGameRecords(true) _
            , vAthletes _
    )

    Dim htmlReader As HtmlTemplateWrapper: Set htmlReader = New HtmlTemplateWrapper
    htmlReader.Initialize (CompetitionRepository.ReadSettinng.HtmlTemplate.StartListTemplatePath)

    Dim writer As HtmlStartListWriter: Set writer = New HtmlStartListWriter
    Call writer.Initialize(htmlReader.HtmlCollection, vStartListData)

    Dim htmlWriter As TrackHtmlOutputer: Set htmlWriter = New TrackHtmlOutputer
    Call htmlWriter.Initialize(writer)

    Dim vFileName As String: vFileName = "Start_" & CompetitionRepository.ReadSettinng.Category & CompetitionRepository.ReadSettinng.Sex & CompetitionRepository.ReadSettinng.EventName & pRoundName & ".html"
    WriteStartList = htmlWriter.WriteFile(vFileName)
End Function

Private Function GenerateEmptyAthlete() As AthletePersonModel
    With New AthletePersonModel
        Call .Initialize("", "", "", "", "", "", "", "", "", "", Now, "", "", "")
        Set GenerateEmptyAthlete = .Self()
    End With
End Function

Private Function GenerateEmptyEntry() As PersonalEntryModel
    With New PersonalEntryModel
        Call .Initialize("", 0, 0, False, Nothing, Nothing, Nothing)
        Set GenerateEmptyEntry = .Self()
    End With
End Function
