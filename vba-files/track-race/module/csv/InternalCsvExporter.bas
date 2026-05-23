Attribute VB_Name = "InternalCsvExporter"
Option Explicit
Option Private Module
'namespace=vba-files/track-race/module/csv

Public Sub WriteCsv()
    Dim vCompetition As CompetitionInfoModel    : Set vCompetition = CompetitionRepository.ReadCompetition()
    Dim vEvent As CompetitionEventModel         : Set vEvent = CompetitionRepository.ReadSettinng().CompetitionEvent

    Dim vFileName As String: vFileName = CompetitionRepository.ReadSettinng.HtmlTemplate.OutputDirectoryPath & "\TRK_" & vEvent.EventKey & ".csv"
    Dim vFileStream As Object: Set vFileStream = CreateObject("ADODB.Stream")
    vFileStream.Open
    vFileStream.Charset = "SJIS"
    Call CsvExporter.ExportToCsv(vFileStream, vCompetition, vEvent)
    Call vFileStream.SaveToFile(vFileName, 2)
    vFileStream.Close
End Sub