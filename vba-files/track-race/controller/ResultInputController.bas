Attribute VB_Name = "ResultInputController"
'namespace=vba-files/track-race/controller
Option Explicit
Option Private Module

Private Enum ResultSheetColumnEnum
    IDX_START_ORD = 1
    IDX_RANK
    IDX_REAL_RESULT
    IDX_REMARK
    IDX_REACTION_TIME
    IDX_NOT_STARTED
    IDX_NOT_FINISHED
    IDX_DISQUALIFIED_REASON
    IDX_ADVANCED_TO_NEXT_ROUND
    IDX_DEMO_ENTRY
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
    IDX_OFFICIAL_RESULT
    IDX_OFFICIAL_SCORE
    IDX_PERSONAL_ID
    IDX_RESULT_ID
    IDX_UPDATED_AT
    IDX_SUB_RESULT_ID
    IDX_SUB_NUMBER
End Enum

Private Enum OrderListColumnEnum
    IDX_ORD = 1
    IDX_PERSONAL_ID
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
    IDX_NOT_STARTED
    IDX_RESULT_ID
End Enum

Private Const TABLE_ROW_OFFSET As Long = 10
Private Const RESULT_LIST_TABLE_COLUMN_OFFSET As Long = 1
Private Const ORDER_LIST_TABLE_COLUMN_OFFSET As Long = 30

Private m_StartOrderIdList As Object
Private m_StartOrderList As StartListGroupModel
Private m_ResultOrderList As TrackResultOrderModels
Private m_ResultSubOrderList As TrackSubResultOrderModels

Private m_ResultList As Range
Private m_StartList As Range

Public Function ChangingGroup(TargetRoundName As String, TargetGroup As Long) As Boolean
    Dim vRounds As EventRoundModels : Set vRounds = EventRoundRepository.ReadAllRounds()
    Dim pRound As EventRoundModel   : Set pRound = vRounds.FindByName(TargetRoundName)
    Dim pGroupNumber As Long        : pGroupNumber = TargetGroup

    If (Range_TargetRoundId().Value <> pRound.Id Or Range_TargetGroup().Value <> pGroupNumber) Then 
        Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
        Dim vConfirmed As Boolean: vConfirmed =(MessageFactory.Generate("SQ003").ToMsgBox() = vbYes)
        Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()

        If Not (vConfirmed) Then
            Call RollBackGroupInput()
        End If

        ChangingGroup = vConfirmed
    Else
        ChangingGroup = True
    End If
End Function

Public Sub InitializeResultList(TargetRoundName As String, TargetGroup As Long)
    Dim pRoundId As Long: pRoundId = EventRoundRepository.ReadAllRounds().FindByName(TargetRoundName).Id
    ' 入力欄のリセット
    Call Range_OldStartList(True).Clear()
    Call Range_OldResultList(True).Clear()
    Set m_StartOrderIdList = Nothing
    Set m_StartOrderList = Nothing
    Set m_StartOrderIdList = CreateObject("Scripting.Dictionary")

    Call WriteToStartList(pRoundId, LoadStartOrderList(pRoundId, TargetGroup), LoadResultOrderList(pRoundId, TargetGroup, true))
    Call LoadResultSubOrderList(pRoundId, TargetGroup, true)

    Range_TargetRoundId().Value = pRoundId
    Range_TargetGroup().Value = TargetGroup
End Sub

Public Sub ReadResults(TargetRoundName As String, TargetGroup As Long, Optional Output As Boolean = False)
    Dim pRoundId As Long    : pRoundId = EventRoundRepository.ReadAllRounds().FindByName(TargetRoundName).Id
    Dim pGroup As Long      : pGroup = TargetGroup

    Dim oldResultList As TrackResultOrderModels
    Dim oldResultSubList As TrackSubResultOrderModels
    If (Output) Then
        ' リポジトリから取得
        Set oldResultList = LoadResultOrderList(pRoundId, pGroup)
        Set oldResultSubList = LoadResultSubOrderList(pRoundId, pGroup)
    End If

    Dim oldStartList As StartListGroupModel: Set oldStartList = LoadStartOrderList(pRoundId, pGroup)
    If (ProgramEntryRepository.ReadAllStartList().FilterByGroup(pGroup, pRoundId).Count() < 1) Then
        Err.Raise CustomErrorCodeEnum.NoEntryAssigned, "ResultInputController.ReadResults", MessageFactory.Generate("SE021").Prompt
    End If

    Call WriteToResultSheetFromGroup(pRoundId, pGroup)

    ' スタートリストから取得
    Call WriteToResultSheetFromStartList(pRoundId, oldStartList, oldResultList, oldResultSubList)

    ' ラップタイム表の設定
    Call WriteToSplitTimeList(pRoundId, pGroup, oldStartList, oldResultList, oldResultSubList)
End Sub

Public Sub RegisterResults()
    Call TrackSubResultGroupRepository.ReadAllRecords()
    Call TrackSubResultGroupRepository.Upsert(CreateGroupDataInput())

    Call TrackOrderResultsReposotory.ReadAllTrackResult()
    Call TrackOrderResultsReposotory.Upsert(CreateResultInput().All())

    Call TrackSubResultOrderRepository.ReadAllRecords()
    Call TrackSubResultOrderRepository.Upsert(CreateSubResultInput())
End Sub

Public Sub CalculateRanking(TargetRoundName As String)
    Dim pRound As EventRoundModel: Set pRound = EventRoundRepository.ReadAllRounds().FindByName(TargetRoundName)
    Call TrackResultRankingService.UpdateTotalRanking(pRound.Id)
    Call QualifierDrawerController.LoadQualifierLots(pRound.Id)
    Call ResultRankingWritingService.WriteRanking(pRound)
End Sub

Public Sub WriteRanking(TargetRoundName As String)
    Dim pRound As EventRoundModel: Set pRound = EventRoundRepository.ReadAllRounds().FindByName(TargetRoundName)
    Call HtmlLoader.OpenHtmlFile(ResultRankingWritingService.WriteRanking(pRound))
End Sub

Public Sub WriteResultList(TargetRoundName As String)
    Dim pRound As EventRoundModel: Set pRound = EventRoundRepository.ReadAllRounds().FindByName(TargetRoundName)
    Call HtmlLoader.OpenHtmlFile(ResultListWritingService.WriteResultList(pRound))
End Sub

Public Function ResultHasRegistered() As Boolean
    ResultHasRegistered = m_ResultOrderList.Count > 0
End Function

Private Function CreateResultInput() As TrackResultOrderModels
    Dim vRowRange As Range
    Dim vRecords As TrackResultOrderModels: Set vRecords = New TrackResultOrderModels
    Dim vCurrentDate As Date: vCurrentDate = Now

    For Each vRowRange In Range_OldResultList().Resize(Range_OldStartList().Rows.Count).Rows()
        Dim vResultId As String: vResultId = vRowRange.Cells(1, ResultSheetColumnEnum.IDX_RESULT_ID).value
        If Not (vResultId = "" Or vResultId = "?") Then
            Dim vRemarks As String          : vRemarks = vRowRange.Cells(1, ResultSheetColumnEnum.IDX_REMARK).value
            Dim vNotStarted As Boolean      : vNotStarted = Not(vRowRange.Cells(1, ResultSheetColumnEnum.IDX_NOT_STARTED).value = "") Or (InStr(vRemarks, "DNS") > 0)
            Dim vNotFinished As Boolean     : vNotFinished = Not(vRowRange.Cells(1, ResultSheetColumnEnum.IDX_NOT_FINISHED).value = "") Or (InStr(vRemarks, "DNF") > 0)
            Dim vDisqualified As Boolean    : vDisqualified = Not(vRowRange.Cells(1, ResultSheetColumnEnum.IDX_DISQUALIFIED_REASON).value = "") Or (InStr(vRemarks, "DQ") > 0)
            Dim vDemoEntry As Boolean       : vDemoEntry = Not(vRowRange.Cells(1, ResultSheetColumnEnum.IDX_DEMO_ENTRY).value = "") Or (InStr(vRemarks, "OPN") > 0)

            If Not (vRemarks = "") Then
                Dim vComments As Collection: Set vComments = New Collection
                Dim i As Long
                Dim vCommentArray() As String: vCommentArray = Split(vRemarks, ",")
                For i = LBound(vCommentArray) To UBound(vCommentArray)
                    Select Case Trim(vCommentArray(i)) 
                        Case "DNS", "DNF", "DQ", "OPN", " " 
                            ' DO NOTHING
                        Case Else 
                            Call vComments.Add(Trim(vCommentArray(i)))
                    End Select
                Next i

                vRemarks = StringUtil.JoinCollectionToString(vComments, ", ")
            End If

            Dim vRound As EventRoundModel: Set vRound = EventRoundRepository.ReadAllRounds().Item(Range_TargetRoundId().value)
            Dim vGroup As Long: vGroup = Range_TargetGroup().Value
            Dim vOrder As Long: vOrder = vRowRange.Cells(1, ResultSheetColumnEnum.IDX_START_ORD).value

            If Not (vRecords.FindByOrder(vRound.Id, vGroup, vOrder) Is Nothing) Then
                Err.Raise CustomErrorCodeEnum.ResultOrderDuplicated, "ResultInputController.CreateResultInput", MessageFactory.Generate("SE020").Prompt(vRound.Name, vGroup, vOrder)
            End If

            With New TrackResultOrderModel
                Call .Initialize( _
                        Range_TargetRoundId().value _
                        , vGroup _
                        , vOrder _
                        , vRowRange.Cells(1, ResultSheetColumnEnum.IDX_RANK).value _
                        , TimeProcessFunction.ParseTimeResultValue(vRowRange.Cells(1, ResultSheetColumnEnum.IDX_REAL_RESULT).value) _
                        , TimeProcessFunction.ParseTimeResultValue(vRowRange.Cells(1, ResultSheetColumnEnum.IDX_OFFICIAL_RESULT).value) _
                        , vRemarks _
                        , vRowRange.Cells(1, ResultSheetColumnEnum.IDX_REACTION_TIME).value _
                        , vNotStarted _
                        , vNotFinished _
                        , vRowRange.Cells(1, ResultSheetColumnEnum.IDX_DISQUALIFIED_REASON).value _
                        , vRowRange.Cells(1, ResultSheetColumnEnum.IDX_ADVANCED_TO_NEXT_ROUND).value _
                        , vDemoEntry _
                        , vRowRange.Cells(1, ResultSheetColumnEnum.IDX_OFFICIAL_SCORE).value _
                        , IIf(vRowRange.Cells(1, ResultSheetColumnEnum.IDX_UPDATED_AT).value = "", vCurrentDate, vRowRange.Cells(1, ResultSheetColumnEnum.IDX_UPDATED_AT).value) _
                        , vResultId _
                )

                Call vRecords.Add(.Self())
            End With
        End If
    Next vRowRange
    Set CreateResultInput = vRecords
End Function

Private Function CreateSubResultInput() As Collection
    Dim vRowRange As Range
    Dim vRecords As Collection: Set vRecords = New Collection
    Dim vCurrentDate As Date: vCurrentDate = Now

    For Each vRowRange In Range_OldResultList().Rows()
        Dim vSubResultId As String: vSubResultId = vRowRange.Cells(1, ResultSheetColumnEnum.IDX_SUB_RESULT_ID).Value
        Dim vResultId As String: vResultId = vRowRange.Cells(1, ResultSheetColumnEnum.IDX_RESULT_ID).Value

        If Not (vSubResultId = "" Or vSubResultId = "?" Or vResultId = "" Or vResultId = "?") Then
            Dim vPrecision As Long
            If (vRowRange.Cells(1, ResultSheetColumnEnum.IDX_SUB_NUMBER).Value = 0) Then
                vPrecision = CompetitionRepository.ReadSettinng().ResultPrecision
            Else
                vPrecision = CompetitionRepository.ReadSettinng().SplitTimePrecision
            End If

            Dim vSecondDelimiter As String: vSecondDelimiter = CompetitionRepository.ReadSettinng().SecondDelimiter
            Dim vSubResultValue As String: vSubResultValue = vRowRange.Cells(1, ResultSheetColumnEnum.IDX_REAL_RESULT).value
            If (vPrecision = 0 And InStr(vSecondDelimiter, vSubResultValue) = 0) Then
                If (Right(vSubResultValue, Len(vSecondDelimiter)) <> vSecondDelimiter) Then
                    vSubResultValue = vSubResultValue & vSecondDelimiter
                End If
            End If

            With New TrackSubResultOrderModel
                Call .Initialize( _
                        vSubResultId _
                        , Range_TargetRoundId().value _
                        , Range_TargetGroup().Value _
                        , vRowRange.Cells(1, ResultSheetColumnEnum.IDX_START_ORD).value _
                        , vRowRange.Cells(1, ResultSheetColumnEnum.IDX_SUB_NUMBER).Value _
                        , vRowRange.Cells(1, ResultSheetColumnEnum.IDX_REMARK).value _
                        , TimeProcessFunction.ParseTimeResultValue(vSubResultValue) _
                        , vResultId _
                        , IIf(vRowRange.Cells(1, ResultSheetColumnEnum.IDX_UPDATED_AT) = "", vCurrentDate, vRowRange.Cells(1, ResultSheetColumnEnum.IDX_UPDATED_AT)) _
                )

                Call vRecords.Add(.Self())
            End With
        End If
    Next vRowRange
    Set CreateSubResultInput = vRecords
End Function

Private Function CreateGroupDataInput() As Collection
    Dim vRecords As Collection: Set vRecords = New Collection
    Dim vWindText As String: vWindText = Range_Wind().Value

    With New TrackSubResultGroupModel
        Call .Initialize( _
                Range_TargetRoundId().Value _
                , Range_TargetGroup().Value _
                , IIf(CompetitionRepository.ReadSettinng().OperationWindGauge, WindProcessFunction.ConvertWind(vWindText), Nothing) _
                , Now _
                , Range_GroupId().Value _
        )

        Call vRecords.Add(.Self(), .Id)
        Range_GroupId().Value = .Id
    End With

    Set CreateGroupDataInput = vRecords
End Function

Private Sub WriteToStartList(RoundId As Long, StartList As StartListGroupModel, ResultList As TrackResultOrderModels)
    Dim record As StartListOrderModel
    Dim inputRow As Range: Set inputRow = Range_OldStartList(True)
    For Each record In StartList.Orders.All()
        Dim resultData As TrackResultOrderModel: Set resultData = ResultList.FindByOrder(RoundId, record.Group, record.Order)

        Dim resultId As String
        If (resultData Is Nothing) Then
            resultId = UuidFactory.GenerateUuid()
        Else
            resultId = resultData.Id
        End If

        inputRow.Cells(1, OrderListColumnEnum.IDX_RESULT_ID).value = resultId

        inputRow.Cells(1, OrderListColumnEnum.IDX_ORD).value = record.Order
        inputRow.Cells(1, OrderListColumnEnum.IDX_PERSON_SEX).value = record.Person.Sex
        inputRow.Cells(1, OrderListColumnEnum.IDX_PERSON_BIB).value = "'" & record.Person.Bib
        inputRow.Cells(1, OrderListColumnEnum.IDX_PERSON_PERSONAL_NAME).value = record.Person.PersonalName
        inputRow.Cells(1, OrderListColumnEnum.IDX_PERSON_AGE).value = record.Person.Age
        inputRow.Cells(1, OrderListColumnEnum.IDX_PERSON_PERSONAL_PHONETIC).value = record.Person.PersonalPhonetic
        inputRow.Cells(1, OrderListColumnEnum.IDX_PERSON_PERSONAL_LATIN).value = record.Person.PersonalLatin
        inputRow.Cells(1, OrderListColumnEnum.IDX_PERSON_TEAM_NAME).value = record.Person.TeamName
        inputRow.Cells(1, OrderListColumnEnum.IDX_PERSON_TEAM_PLACE).value = record.Person.TeamPlace
        inputRow.Cells(1, OrderListColumnEnum.IDX_PERSON_PERSONAL_COUNTRY).value = record.Person.PersonalCountry
        inputRow.Cells(1, OrderListColumnEnum.IDX_PERSON_PERSONAL_GRADE).value = record.Person.PersonalGrade
        inputRow.Cells(1, OrderListColumnEnum.IDX_ACCUMALTIVE).value = StringUtil.JoinCollectionToString(record.Accumaltive, ", ")

        Set inputRow = inputRow.Offset(1, 0)
    Next record
    Call Range_OldStartList(True)
    Call LoadStartOrderIdList(RoundId, StartList.GroupNumber, True)
End Sub

Private Sub WriteToResultSheetFromGroup(RoundId As Long, Group As Long)
    ' 組データのリセット
    Range_GroupId().ClearContents
    Range_Wind().ClearContents

    Dim groupData As TrackSubResultGroupModel: Set groupData = TrackSubResultGroupRepository.ReadAllRecords().FindByGroup(RoundId, Group)
    If (groupData Is Nothing) Then
        ' 組データ未作成のため処理終了
        Exit Sub
    End If

    ' グループデータのロード
    Range_GroupId().Value = groupData.Id

    If Not (CompetitionRepository.ReadSettinng().OperationWindGauge) Then
        ' 風速を測る種目ではないため、何もしない
    ElseIf (groupData.Wind Is Nothing) Then
        ' 風力データなし.
    Else
        ' 風速データのロード
        Range_Wind().Value = groupData.Wind.Value
    End If
End Sub

Private Sub WriteToResultSheetFromStartList(RoundId As Long, StartList As StartListGroupModel, ResultList As TrackResultOrderModels, ResultSubList As TrackSubResultOrderModels)
    Dim record As StartListOrderModel
    Dim inputRow As Range: Set inputRow = Range_OldResultList(true)
    Dim vStartListOrders As StartListOrderModels: Set vStartListOrders = StartList.Orders
    Dim vPlace As Long: vPlace = 1

    If (ResultList Is Nothing) Then
        Set vStartListOrders = vStartListOrders.SortByRankable()
    End If

    For Each record In vStartListOrders.All()
        If (ResultList Is Nothing) Then
            If (record.NotStarted) Then
                Call LoadResultOrderData(inputRow, record)
            Else
                Call LoadResultOrderData(inputRow)
                If Not (record.DemoEntry) Then
                    inputRow.Cells(1,ResultSheetColumnEnum.IDX_RANK).Value = vPlace
                    vPlace = vPlace + 1
                End If
            End If
        Else
            Dim resultData As TrackResultOrderModel: Set resultData = ResultList.FindByOrder(RoundId, record.Group, record.Order)
            Dim resultSubData As TrackSubResultOrderModel
            If Not (resultData Is Nothing) Then
                Set resultSubData = ResultSubList.FindByResultId(resultData.Id, 0)
            End If

            Call LoadResultOrderData(inputRow, record, resultData, resultSubData)
        End If
        Set inputRow = inputRow.Offset(1, 0)
    Next record

    With Range_OldResultList(true)
        Range(.Columns(ResultSheetColumnEnum.IDX_START_ORD), .Columns(ResultSheetColumnEnum.IDX_RANK)).NumberFormatLocal = ""
        Range(.Columns(ResultSheetColumnEnum.IDX_REAL_RESULT), .Columns(ResultSheetColumnEnum.IDX_PERSON_PERSONAL_NAME)).NumberFormatLocal = "@"
        Range(.Columns(ResultSheetColumnEnum.IDX_PERSON_PERSONAL_PHONETIC), .Columns(ResultSheetColumnEnum.IDX_OFFICIAL_RESULT)).NumberFormatLocal = "@"
        Range(.Columns(ResultSheetColumnEnum.IDX_PERSONAL_ID), .Columns(ResultSheetColumnEnum.IDX_RESULT_ID)).NumberFormatLocal = "@"
        .Columns(ResultSheetColumnEnum.IDX_PERSON_AGE).NumberFormatLocal = ""
        .Columns(ResultSheetColumnEnum.IDX_OFFICIAL_SCORE).NumberFormatLocal = ""
        .Columns(ResultSheetColumnEnum.IDX_UPDATED_AT).NumberFormatLocal = "yyyy-MM-ddThh:mm:ss"

        .Columns(ResultSheetColumnEnum.IDX_REAL_RESULT).HorizontalAlignment = xlRight
        .Columns(ResultSheetColumnEnum.IDX_OFFICIAL_RESULT).HorizontalAlignment = xlRight

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
    End With
End Sub

Private Sub WriteToSplitTimeList(RoundId As Long, Group As Long, StartList As StartListGroupModel, ResultList As TrackResultOrderModels, ResultSubList As TrackSubResultOrderModels)
    Dim resultTableRowIndex As Long: resultTableRowIndex = Range_OldResultList(true).Rows.Count
    Dim inputRow As Range: Set inputRow = Range_OldResultList().Offset(resultTableRowIndex - 1,0).Resize(1)

    Dim checkPoints As Collection: Set checkPoints = CompetitionRepository.ReadSettinng.SplitTimePoints
    Dim subIndex As Long

    If (checkPoints.Count < 1) Then
        ' スプリットタイム入力はないためここで終了.
        With inputRow.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        Exit Sub
    End If
    
    For subIndex = 1 To checkPoints.Count
        Set inputRow = inputRow.Offset(1, 0)

        Dim splitTimeData As TrackSubResultOrderModel
        If Not (ResultSubList Is Nothing) Then
            Set splitTimeData = resultSubList.FindByGroupSubNumber(RoundId, Group, subIndex)
        End If

        inputRow.Cells(ResultSheetColumnEnum.IDX_REMARK).Value = checkPoints.Item(subIndex) & "m"
        inputRow.Cells(ResultSheetColumnEnum.IDX_RANK).Value = "L"
        inputRow.Cells(ResultSheetColumnEnum.IDX_SUB_NUMBER).Value = subIndex
        If Not (splitTimeData Is Nothing) Then
            Dim startListData As StartListOrderModel: Set startListData = StartList.FindByOrder(splitTimeData.Order)
            Dim resultData As TrackResultOrderModel: Set resultData = ResultList.FindByOrder(RoundId, Group, splitTimeData.Order)

            inputRow.Cells(1, ResultSheetColumnEnum.IDX_SUB_RESULT_ID).value = splitTimeData.Id
            inputRow.Cells(1, ResultSheetColumnEnum.IDX_REAL_RESULT).value = "'" & splitTimeData.SubResult.ToString()
            inputRow.Cells(1, ResultSheetColumnEnum.IDX_RESULT_ID).value = resultData.Id
            inputRow.Cells(1, ResultSheetColumnEnum.IDX_UPDATED_AT).value = splitTimeData.UpdatedAt

            inputRow.Cells(1, ResultSheetColumnEnum.IDX_START_ORD).value = splitTimeData.Order
            inputRow.Cells(1, ResultSheetColumnEnum.IDX_PERSON_SEX).value = startListData.Person.Sex
            inputRow.Cells(1, ResultSheetColumnEnum.IDX_PERSON_BIB).value = "'" & startListData.Person.Bib
            inputRow.Cells(1, ResultSheetColumnEnum.IDX_PERSON_PERSONAL_NAME).value = startListData.Person.PersonalName
            inputRow.Cells(1, ResultSheetColumnEnum.IDX_PERSON_AGE).value = startListData.Person.Age
            inputRow.Cells(1, ResultSheetColumnEnum.IDX_PERSON_PERSONAL_PHONETIC).value = startListData.Person.PersonalPhonetic
            inputRow.Cells(1, ResultSheetColumnEnum.IDX_PERSON_PERSONAL_LATIN).value = startListData.Person.PersonalLatin
            inputRow.Cells(1, ResultSheetColumnEnum.IDX_PERSON_TEAM_NAME).value = startListData.Person.TeamName
            inputRow.Cells(1, ResultSheetColumnEnum.IDX_PERSON_TEAM_PLACE).value = startListData.Person.TeamPlace
            inputRow.Cells(1, ResultSheetColumnEnum.IDX_PERSON_PERSONAL_COUNTRY).value = startListData.Person.PersonalCountry
            inputRow.Cells(1, ResultSheetColumnEnum.IDX_PERSON_PERSONAL_GRADE).value = startListData.Person.PersonalGrade            
        Else
            inputRow.Cells(1, ResultSheetColumnEnum.IDX_SUB_RESULT_ID).value = UuidFactory.GenerateUuid()
            inputRow.Cells(1, ResultSheetColumnEnum.IDX_RESULT_ID).value = "?"
        End If
    Next subIndex

    With Range_OldResultList(true).Offset(resultTableRowIndex, 0).Resize(checkPoints.Count)
        Range(.Columns(ResultSheetColumnEnum.IDX_START_ORD), .Columns(ResultSheetColumnEnum.IDX_RANK)).NumberFormatLocal = ""
        Range(.Columns(ResultSheetColumnEnum.IDX_REAL_RESULT), .Columns(ResultSheetColumnEnum.IDX_PERSON_PERSONAL_NAME)).NumberFormatLocal = "@"
        Range(.Columns(ResultSheetColumnEnum.IDX_PERSON_PERSONAL_PHONETIC), .Columns(ResultSheetColumnEnum.IDX_OFFICIAL_RESULT)).NumberFormatLocal = "@"
        Range(.Columns(ResultSheetColumnEnum.IDX_PERSONAL_ID), .Columns(ResultSheetColumnEnum.IDX_RESULT_ID)).NumberFormatLocal = "@"
        .Columns(ResultSheetColumnEnum.IDX_PERSON_AGE).NumberFormatLocal = ""
        .Columns(ResultSheetColumnEnum.IDX_OFFICIAL_SCORE).NumberFormatLocal = ""
        .Columns(ResultSheetColumnEnum.IDX_UPDATED_AT).NumberFormatLocal = "yyyy-MM-ddThh:mm:ss"

        .Columns(ResultSheetColumnEnum.IDX_REAL_RESULT).HorizontalAlignment = xlRight
        .Columns(ResultSheetColumnEnum.IDX_OFFICIAL_RESULT).HorizontalAlignment = xlRight

        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlHairLine
            .ColorIndex = xlAutomatic
        End With

        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End With
End Sub

Private Sub RollBackGroupInput()
    Range_TargetRoundNameInput().Value = EventRoundRepository.ReadAllRounds().Item(Range_TargetRoundId().Value).Name
    Range_TargetGroupInput().Value = Range_TargetGroup().Value
End Sub

Public Function OnUpdatedResultData(Spot As Range) As Boolean
    OnUpdatedResultData = False
    If (Application.Intersect(Spot, Range_OldResultList()) Is Nothing) Then
        Exit Function
    End If

    Dim currentRow As Range: Set currentRow = Range_OldResultList().Rows(Spot.Row - TABLE_ROW_OFFSET)
    currentRow.Cells(1, ResultSheetColumnEnum.IDX_UPDATED_AT).Value = ""
    OnUpdatedResultData = True
End Function

Public Function OnUpdatedRealResult(Spot As Range) As Boolean
    OnUpdatedRealResult = False
    If (Application.Intersect(Spot, Range_OldResultList().Columns(ResultSheetColumnEnum.IDX_REAL_RESULT)) Is Nothing) Then
        Exit Function
    End If

    OnUpdatedRealResult = True
    Dim currentRow As Range: Set currentRow = Range_OldResultList().Rows(Spot.Row - TABLE_ROW_OFFSET)
    Dim putValue As String: putValue = Spot.Value
    Dim vPrecision As Long
    
    If (Spot.Row <= TABLE_ROW_OFFSET + Range_OldStartList().Rows.Count) Then 
        vPrecision = CompetitionRepository.ReadSettinng().ResultPrecision
    Else
        vPrecision = CompetitionRepository.ReadSettinng().SplitTimePrecision
    End If

    Dim vRealResult As TimeResultValue
    Set vRealResult = TimeProcessFunction.ParseFromNumberString(Spot.Value, vPrecision)
    If (vRealResult Is Nothing) Then
        If (vPrecision = 0) Then
            If (Right(putValue, Len(CompetitionRepository.ReadSettinng().SecondDelimiter)) <> CompetitionRepository.ReadSettinng().SecondDelimiter) Then
                putValue = putValue & CompetitionRepository.ReadSettinng().SecondDelimiter
            End If
        End If
        Set vRealResult = TimeProcessFunction.ParseTimeResultValue(putValue)
    End If

    Spot.Value = "'" & vRealResult.ToString()
    Dim vOfficialResult As TimeResultValue: Set vOfficialResult = vRealResult.RoundValue(vPrecision)
    currentRow.Cells(1, ResultSheetColumnEnum.IDX_OFFICIAL_RESULT).Value = "'" & vOfficialResult.ToString()

    If (CompetitionRepository.ReadSettinng.ScoringCombinedEvents) Then
        ' 混成の点数.
        Dim point As Variant
        If (vOfficialResult.Valid) Then 
            point = CDec(CompetitionRepository.ReadSettinng.ScoringIndexNumber2) - CDec(vOfficialResult.Value) / CDec(1000)
            point = point ^ CDec(CompetitionRepository.ReadSettinng.ScoringIndexNumber3)
            point = CDec(CompetitionRepository.ReadSettinng.ScoringIndexNumber1) * point
        Else
            point = CDec(0)
        End If

        currentRow.Cells(1, ResultSheetColumnEnum.IDX_OFFICIAL_SCORE).Value = Int(point)
    End If
End Function

Public Function OnUpdateOrder(Spot As Range) As Boolean
    OnUpdateOrder = False
    If (Application.Intersect(Spot, Range_OldResultList().Columns(ResultSheetColumnEnum.IDX_START_ORD)) Is Nothing) Then
        Exit Function
    End If

    Dim currentRow As Range: Set currentRow = Range_OldResultList().Rows(Spot.Row - TABLE_ROW_OFFSET)
    Dim orderKey As String: orderKey = Range_TargetRoundId().value & "__" & Range_TargetGroup().Value & "__" & Spot.Value()
    Dim startData As StartListOrderModel: Set startData = LoadStartOrderList(Range_TargetRoundId().value, Range_TargetGroup().Value).FindByOrder(Spot.Value)
    Dim resultData As TrackResultOrderModel: Set resultData = LoadResultOrderList(Range_TargetRoundId().value, Range_TargetGroup().Value).FindByOrder(Range_TargetRoundId().value, Range_TargetGroup().Value, Spot.Value())

    Dim resultSubData As TrackSubResultOrderModel
    If Not (resultData Is Nothing) Then
        Set resultSubData = LoadResultSubOrderList(Range_TargetRoundId().value, Range_TargetGroup().Value).FindByResultId(resultData.Id, 0)
    End If

    If (Spot.Row <= TABLE_ROW_OFFSET + Range_OldStartList().Rows.Count) Then
        ' 成績の登録
        Call LoadResultOrderData(currentRow, startData, resultData, resultSubData)
    Else
        ' ラップタイム表の登録
        Call LoadSplitTimeOrderData(currentRow, startData, resultData)
    End If
    OnUpdateOrder = True
End Function

Public Sub OnWindInputSelect(Target As Range)
    If Not (CompetitionRepository.ReadSettinng().OperationWindGauge) Then
        If Not (Application.Intersect(Target, Range_Wind()) Is Nothing) Then
            Range_TargetGroupInput.Select
        End If
    End If
End Sub

Public Function ThisSheet() As Worksheet
    Set ThisSheet = ResultInputSheet
End Function

Public Function Range_OldResultList(Optional Reload As Boolean = False) As Range
    If (Not (m_ResultList Is Nothing) And Not Reload) Then
        Set Range_OldResultList = m_ResultList
    End If

    Dim rowCount As Long: rowCount = Application.WorksheetFunction.CountA(ThisSheet().Columns(ResultSheetColumnEnum.IDX_RESULT_ID + RESULT_LIST_TABLE_COLUMN_OFFSET))
    Set m_ResultList = ThisSheet().Cells(TABLE_ROW_OFFSET, RESULT_LIST_TABLE_COLUMN_OFFSET + 1).Offset(1, 0).Resize(rowCount - IIf(rowCount > 1, 1, 0), 28)
    Set Range_OldResultList = m_ResultList
End Function

Public Function Range_OldStartList(Optional Reload As Boolean = False) As Range
    If (Not (m_StartList Is Nothing) And Not Reload) Then
        Set Range_OldStartList = m_StartList
    End If

    Dim rowCount As Long: rowCount = Application.WorksheetFunction.CountA(ThisSheet().Columns(OrderListColumnEnum.IDX_ORD + ORDER_LIST_TABLE_COLUMN_OFFSET))
    rowCount = IIf(rowCount > 1, rowCount, 1)
    Set m_StartList = ThisSheet().Cells(TABLE_ROW_OFFSET, ORDER_LIST_TABLE_COLUMN_OFFSET + 1).Offset(1, 0).Resize(rowCount - IIf(rowCount > 2, 2, 1), 10)
    Set Range_OldStartList = m_StartList
End Function

Public Function Range_TargetRoundNameInput() As Range
    Set Range_TargetRoundNameInput = ThisSheet().Cells(3, 4)
End Function  

Public Function Range_TargetGroupInput() As Range
    Set Range_TargetGroupInput = ThisSheet().Cells(4, 4)
End Function
    
Public Function Range_TargetRoundId() As Range
    Set Range_TargetRoundId = ThisSheet().Cells(3, 5)
End Function

Public Function Range_TargetGroup() As Range
    Set Range_TargetGroup = ThisSheet().Cells(4, 5)
End Function

Public Function Range_GroupId() As Range
    Set Range_GroupId = ThisSheet().Cells(4, 6)
End Function

Public Function Range_Wind() As Range
    Set Range_Wind = ThisSheet().Cells(5, 4)
End Function

Private Sub LoadResultOrderData(RowToLoad As Range, Optional StartOrderData As StartListOrderModel = Nothing, Optional ResultOrderData As TrackResultOrderModel = Nothing, Optional ResultSubData As TrackSubResultOrderModel = Nothing)
    If (StartOrderData Is Nothing) Then
        Call RowToLoad.Clear()
        RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_RESULT_ID).value = "?"
        Exit Sub
    End If

    Dim vRemarks As String: vRemarks = ""
    If (ResultOrderData Is Nothing) Then
        If (StartOrderData.NotStarted) Then
            vRemarks = vRemarks & ", " & "DNS"
            RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_NOT_STARTED).value = "*"
        End If

        If (StartOrderData.DemoEntry) Then
            vRemarks = vRemarks & ", " & "OPN"
            RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_DEMO_ENTRY).value = "*"
        End If

        vRemarks = Replace(vRemarks, ", ", "", , 1)

        RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_REMARK).value = vRemarks
        RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_RESULT_ID).value = LoadStartOrderIdList(Range_TargetRoundId().Value, StartOrderData.Group).Item(Range_TargetRoundId().Value & "__" & StartOrderData.Group & "__" & StartOrderData.Order)
        RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_SUB_RESULT_ID).value = UuidFactory.GenerateUuid()
        RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_SUB_NUMBER).value = 0
    Else
        If (ResultOrderData.NotStarted) Then
            vRemarks = vRemarks & ", " & "DNS"
            RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_NOT_STARTED).value = "*"
        End If

        If (ResultOrderData.NotFinished) Then
            vRemarks = vRemarks & ", " & "DNF"
            RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_NOT_FINISHED).value = "*"
        End If

        If (ResultOrderData.DemoEntry) Then
            vRemarks = vRemarks & ", " & "OPN"
            RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_DEMO_ENTRY).value = "*"
        End If

        If Not (ResultOrderData.Remark = "") Then
            vRemarks = vRemarks & ", " & ResultOrderData.Remark
        End If

        vRemarks = Replace(vRemarks, ", ", "", , 1)
        RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_RANK).value = IIF(ResultOrderData.Rank <= 0, "", ResultOrderData.Rank)
        RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_REAL_RESULT).value = "'" & ResultOrderData.RealResult.ToString()
        RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_OFFICIAL_RESULT).value = "'" & ResultOrderData.Result.ToString()
        RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_REACTION_TIME).value = IIf(ResultOrderData.ReactionTime <= 0, "", ResultOrderData.ReactionTime)
        RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_REMARK).value = vRemarks
        RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_DISQUALIFIED_REASON).value = ResultOrderData.DisqualifiedReason
        RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_ADVANCED_TO_NEXT_ROUND).value = ResultOrderData.AdvancedNextRound
        RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_RESULT_ID).value = ResultOrderData.Id
        RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_UPDATED_AT).value = ResultOrderData.UpdatedAt

        If (ResultSubData Is Nothing) Then 
            RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_SUB_RESULT_ID).value = UuidFactory.GenerateUuid()
        Else
            RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_SUB_RESULT_ID).value = ResultSubData.Id
        End If

        RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_SUB_NUMBER).value = 0

        If (CompetitionRepository.ReadSettinng().ScoringCombinedEvents) Then
            RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_OFFICIAL_SCORE).value = ResultOrderData.Score
        Else
            RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_OFFICIAL_SCORE).value = ""
        End If
    End If

    RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_START_ORD).value = StartOrderData.Order
    RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_PERSON_SEX).value = StartOrderData.Person.Sex
    RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_PERSON_BIB).value = "'" & StartOrderData.Person.Bib
    RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_PERSON_PERSONAL_NAME).value = StartOrderData.Person.PersonalName
    RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_PERSON_AGE).value = StartOrderData.Person.Age
    RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_PERSON_PERSONAL_PHONETIC).value = StartOrderData.Person.PersonalPhonetic
    RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_PERSON_PERSONAL_LATIN).value = StartOrderData.Person.PersonalLatin
    RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_PERSON_TEAM_NAME).value = StartOrderData.Person.TeamName
    RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_PERSON_TEAM_PLACE).value = StartOrderData.Person.TeamPlace
    RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_PERSON_PERSONAL_COUNTRY).value = StartOrderData.Person.PersonalCountry
    RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_PERSON_PERSONAL_GRADE).value = StartOrderData.Person.PersonalGrade
    RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_ACCUMALTIVE).value = StringUtil.JoinCollectionToString(StartOrderData.Accumaltive, ", ")
End Sub

Private Sub LoadSplitTimeOrderData(RowToLoad As Range, Optional StartOrderData As StartListOrderModel = Nothing, Optional ResultOrderData As TrackResultOrderModel = Nothing)
    If (StartOrderData Is Nothing) Then
        Call RowToLoad.Clear()
        RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_RESULT_ID).value = "?"
        Exit Sub        
    End If

    If (ResultOrderData Is Nothing) Then
        RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_RESULT_ID).value = LoadStartOrderIdList(Range_TargetRoundId().Value, StartOrderData.Group).Item(Range_TargetRoundId().Value & "__" & StartOrderData.Group & "__" & StartOrderData.Order)
    Else
        RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_RESULT_ID).value = ResultOrderData.Id
    End If

    RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_PERSON_SEX).value = StartOrderData.Person.Sex
    RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_PERSON_BIB).value = "'" & StartOrderData.Person.Bib
    RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_PERSON_PERSONAL_NAME).value = StartOrderData.Person.PersonalName
    RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_PERSON_AGE).value = StartOrderData.Person.Age
    RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_PERSON_PERSONAL_PHONETIC).value = StartOrderData.Person.PersonalPhonetic
    RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_PERSON_PERSONAL_LATIN).value = StartOrderData.Person.PersonalLatin
    RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_PERSON_TEAM_NAME).value = StartOrderData.Person.TeamName
    RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_PERSON_TEAM_PLACE).value = StartOrderData.Person.TeamPlace
    RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_PERSON_PERSONAL_COUNTRY).value = StartOrderData.Person.PersonalCountry
    RowToLoad.Cells(1, ResultSheetColumnEnum.IDX_PERSON_PERSONAL_GRADE).value = StartOrderData.Person.PersonalGrade
End Sub

Private Function LoadStartOrderList(TargetRoundId As Long, TargetGroup As Long) As StartListGroupModel
    If (Not(m_StartOrderList Is Nothing)) Then
        If ((m_StartOrderList.RoundId = TargetRoundId) And (m_StartOrderList.GroupNumber = TargetGroup)) Then
            Set LoadStartOrderList = m_StartOrderList
            Exit Function
        End If
    End If

    Set m_StartOrderList = New StartListGroupModel
    Call m_StartOrderList.Initialize(TargetRoundId, TargetGroup, ProgramEntryRepository.ReadAllStartList(true).FilterByGroup(TargetGroup, TargetRoundId))
    Set LoadStartOrderList = m_StartOrderList
End Function

Private Function LoadStartOrderIdList(TargetRoundId As Long, TargetGroup As Long, Optional Reload As Boolean = False) As Object
    If (Not(m_StartOrderIdList Is Nothing) And (Reload = False)) Then
        Set LoadStartOrderIdList = m_StartOrderIdList
    End If

    Dim currentRow As Range
    Dim tempStartOrderIdList As Object: Set tempStartOrderIdList = CreateObject("Scripting.Dictionary")
    For Each currentRow In Range_OldStartList().Rows()
        Call tempStartOrderIdList.Add(TargetRoundId & "__" & TargetGroup & "__" & currentRow.Cells(1, OrderListColumnEnum.IDX_ORD).value, currentRow.Cells(1, OrderListColumnEnum.IDX_RESULT_ID).value)
    Next currentRow

    Set m_StartOrderIdList = tempStartOrderIdList
    Set LoadStartOrderIdList = m_StartOrderIdList
End Function

Private Function LoadResultOrderList(TargetRoundId As Long, TargetGroup As Long, Optional Reload As Boolean = False) As TrackResultOrderModels
    If (Not(m_ResultOrderList Is Nothing)) Then
        If (Reload = False) Then
            Set LoadResultOrderList = m_ResultOrderList
            Exit Function
        End If
    End If

    Set m_ResultOrderList = TrackOrderResultsReposotory.ReadAllTrackResult().FilterByRoundGroup(TargetRoundId, TargetGroup).SortByOrder()
    Set LoadResultOrderList = m_ResultOrderList
End Function

Private  Function LoadResultSubOrderList(TargetRoundId As Long, TargetGroup As Long, Optional Reload As Boolean = False) As TrackSubResultOrderModels
    If (Not (m_ResultSubOrderList Is Nothing)) Then
        If (Reload = False) Then
            Set LoadResultSubOrderList = m_ResultSubOrderList
            Exit Function
        End If
    End If

    Set m_ResultSubOrderList = TrackSubResultOrderRepository.ReadAllRecords().FilterByGroup(TargetRoundId, TargetGroup)
    Set LoadResultSubOrderList = m_ResultSubOrderList
End Function
