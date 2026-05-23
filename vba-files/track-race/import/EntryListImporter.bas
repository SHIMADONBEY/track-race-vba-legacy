Attribute VB_Name = "EntryListImporter"
'namespace=vba-files/track-race/import
Option Explicit

Public Sub ImportEntries(Entries, Persons)
    Dim pPersons As AthletePersonModels: Set pPersons = JsonConverter.ParseJsonTo(Persons, New AthletePersonModels)

    Dim vParentParser As EntriesParser          : Set vParentParser = New EntriesParser
    Dim vChildParser As TimeResultEntryParser   : Set vChildParser = New TimeResultEntryParser
    Call vChildParser.Initialize(CompetitionRepository.ReadSettinng().MinuteDelimiter, CompetitionRepository.ReadSettinng().SecondDelimiter)
    Call vParentParser.Initialize(vChildParser)

    Dim pEntries As PersonalEntryModels: Set pEntries = JsonConverter.ParseJsonTo(Entries, New PersonalEntryModels, vParentParser)

    Dim vStartList As StartListOrderModels  : Set vStartList = ProgramEntryRepository.ReadAllStartList()
    Dim vStartEntriesPerPerson As Object    : Set vStartEntriesPerPerson = CollectionUtil.GroupingBy(vStartList.All(), New StartListPersonClassifier)
    Dim vEntriesPerPerson As Object         : Set vEntriesPerPerson = pEntries.GroupByPerson()

    Dim vPerson As AthletePersonModel
    For Each vPerson In pPersons.All()
        Dim vPersonId As String: vPersonId = vPerson.Id
        Dim vElements As Collection: Set vElements = New Collection
        Dim aAccumultive() As String

        If (vStartEntriesPerPerson.Exists(vPersonId)) Then 
            Dim vItem As StartListOrderModel
            For Each vItem In vStartEntriesPerPerson.Item(vPersonId)
                aAccumultive = Split(StringUtil.JoinCollectionToString(vItem.Accumaltive, ","), ",")
                With New StartListOrderModel
                    Call .Initialize( _
                            vPerson _
                            , vItem.Entry _
                            , vItem.QualifiedRank _
                            , vItem.RoundId _
                            , aAccumultive _
                            , vItem.NotStarted _
                            , vItem.Group _
                            , vItem.Order _
                            , vItem.Id _
                    )

                    Call vElements.Add(.Self())
                End With
            Next vItem 
        Else
            Dim vEntry As PersonalEntryModel
            For Each vEntry In vEntriesPerPerson.Item(vPersonId).All()
                ReDim aAccumultive(0 To 0)
                With New StartListOrderModel
                    Call .Initialize( _
                            vPerson _
                            , vEntry _
                            , 0 _
                            , 1 _
                            , aAccumultive _
                            , False _
                    )

                    Call vElements.Add(.Self())
                End With
            Next vEntry 
        End If

        Call ProgramEntryRepository.Upsert(vElements)
    Next vPerson 
End Sub
