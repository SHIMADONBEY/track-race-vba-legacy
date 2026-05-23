Attribute VB_Name = "ResultValueParser"
'namespace=vba-files/common/value
Option Explicit
Option Private Module

Public Function Parse(Expression As String, EventType As String, MinuteDelimiter As String, SecondDelimiter As String, MetreDelimiter As String) As IResultValue
    Dim vParsedValue As IResultValue

    Select Case EventType 
        Case EventTypeConstants.TRACK, EventTypeConstants.RELAY, EventTypeConstants.ROAD
            ' Use TimeResultValue
            Set vParsedValue = New TimeResultValue
            Call vParsedValue.Parse(Expression, MinuteDelimiter, SecondDelimiter)
        Case EventTypeConstants.HEIGHT_JUMP, EventTypeConstants.LENGTH_JUMP, EventTypeConstants.THROW
            ' TODO: For Field Events
            Set vParsedValue = Nothing
            ' Call vParsedValue.Parse(Expression, MetreDelimiter)
        Case EventTypeConstants.COMBINED_EVENTS
            ' TODO: For Combined Events
            Set vParsedValue = Nothing
            ' Call vParsedValue.Parse(Expression)
        Case Else
            Set vParsedValue = Nothing
    End Select

    Set Parse = vParsedValue
End Function
