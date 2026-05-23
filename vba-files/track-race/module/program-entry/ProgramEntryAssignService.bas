Attribute VB_Name = "ProgramEntryAssignService"
'namespace=vba-files/track-race/modules/program-entry
Option Explicit
Option Private Module

Public Function PickUpQualifiersByResults(TargetRoundId As Long, TargetNumber As Long) As StartListOrderModels
    Dim allResults As TrackResultOrderModels: Set allResults = TrackOrderResultsReposotory.ReadAllTrackResult().FilterByRound(TargetRoundId)
    Dim pickedResults As TrackResultOrderModels
    Dim resultRecord As TrackResultOrderModel

    If (TargetNumber < 1) Then 
        Set pickedResults = allResults.FilterQualified(TargetRoundId)
    Else
        Set pickedResults = New TrackResultOrderModels
        For Each resultRecord In allResults.All()
            If (resultRecord.AdvancedNextRound <> "") Then
                Call pickedResults.Add(resultRecord)
            ElseIf (resultRecord.IsRankable() And resultRecord.ResultRank <= TargetNumber) Then
                Call pickedResults.Add(resultRecord)
            End If
        Next resultRecord 
    End If

    Dim currentStartList As StartListOrderModels: Set currentStartList = ProgramEntryRepository.ReadAllStartList(true)
    Dim startListRecords As StartListOrderModels: Set startListRecords = New StartListOrderModels
    For Each resultRecord In pickedResults.All()
        Dim qualifiedOrder As StartListOrderModel: Set qualifiedOrder = currentStartList.FindOrder(TargetRoundId, resultRecord.Group, resultRecord.Order)

        Dim vQualifyingEntry As PersonalEntryModel: Set vQualifyingEntry = New PersonalEntryModel
        Call vQualifyingEntry.Initialize( _
                qualifiedOrder.Person.Id _
                , 0 _
                , 0 _
                , qualifiedOrder.DemoEntry _
                , resultRecord.Result _
                , qualifiedOrder.Qualified1 _
                , qualifiedOrder.Qualified2 _
                , qualifiedOrder.Qualified3 _
        )

        Dim accumaltiveArray() As String: accumaltiveArray = Split(StringUtil.JoinCollectionToString(qualifiedOrder.Accumaltive, "__"), "__")
        With New StartListOrderModel
            Call .Initialize( _
                    qualifiedOrder.Person _
                    , vQualifyingEntry _
                    , resultRecord.ResultRank _
                    , TargetRoundId + 1 _
                    , accumaltiveArray _
                    , false _
            )

            Call startListRecords.Add(.Self())
        End With
    Next resultRecord

    Set PickUpQualifiersByResults = startListRecords
End Function
