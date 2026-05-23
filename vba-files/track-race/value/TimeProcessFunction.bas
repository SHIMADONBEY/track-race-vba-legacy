Attribute VB_Name = "TimeProcessFunction"
'namespace=vba-files/track-race/value
Option Explicit
Option Private Module

Public Function ParseTimeResultValue(ByVal Expression As String) As IResultValue
    Dim pEventSetting As TrackEventSettingModel : Set pEventSetting = CompetitionRepository.ReadSettinng()
    Set ParseTimeResultValue = ResultValueParser.Parse(Expression, EventTypeConstants.TRACK, pEventSetting.MinuteDelimiter, pEventSetting.SecondDelimiter, "")
End Function

Public Function ParseFromNumberString(ByVal NumberString As Variant, ByVal Precision As Long) As IResultValue
    If (CStr(NumberString) Like "*[!0-9]*") Then
        ' It is not parsed because it does not consist only of numbers.
        Set ParseFromNumberString = Nothing
        Exit Function
    End If

    Dim vMinuteDelim As String: vMinuteDelim = CompetitionRepository.ReadSettinng.MinuteDelimiter
    Dim vSecondDelim As String: vSecondDelim = CompetitionRepository.ReadSettinng.SecondDelimiter
    Dim vLength As Long: vLength = Len(NumberString)

    Dim vFormat As String: vFormat = "0" & StringUtil.ConvertFormatString(vSecondDelim) & String(Precision, "0")
    If (vLength > 2 + Precision) Then
        vFormat = "0" & StringUtil.ConvertFormatString(vMinuteDelim) & "0" & vFormat
    End If

    Dim vFormattedString As String: vFormattedString = Format(NumberString, vFormat)
    Set ParseFromNumberString = ParseTimeResultValue(vFormattedString)
End Function

' Public Sub ConvertTimeValueTest()
'     Debug.Print ConvertTimeValue("9""95", "'", """") ' 9950
'     Debug.Print ConvertTimeValue("59""2", "'", """") ' 59200
'     Debug.Print ConvertTimeValue("1:03.17", ":", ".") ' 63170
'     Debug.Print ConvertTimeValue("2:34.5678", ":", ".") ' 154567
'     Debug.Print ConvertTimeValue("1:12:34.5678", ":", ".") ' 2147483647
'     Debug.Print ConvertTimeValue("DNS", ":", ".") ' 2147483647
' End Sub

' Public Sub FormatTimeValueTest()
'     Debug.Print FormatTimeValue(9950, 1, "", """") ' 10"0
'     Debug.Print FormatTimeValue(59980, 2, ":", ".") ' 59.98
'     Debug.Print FormatTimeValue(59981, 2, ":", ".") ' 59.99
'     Debug.Print FormatTimeValue(59991, 2, ":", ".") ' 1:00.00
'     Debug.Print FormatTimeValue(65432, 2, "•ª", "•b") ' 1•ª05•b44
'     Debug.Print FormatTimeValue(65432, 2, "", "•b") ' 65•b44
' End Sub
