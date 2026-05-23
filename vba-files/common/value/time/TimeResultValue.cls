VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "TimeResultValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/common/value/time
Option Explicit
Implements IResultValue

Private Const MAX_DIGITS As Long = 3
Private Const MINUTE_PER_SECOND As Long = 60
Private Const INVALID_VALUE As Long = 2147483647

Private m_Value As Long
Private m_SecondDelimiter As String
Private m_MinuteDelimiter As String
Private m_Scale As Long

Private Sub Class_Initialize()
    m_Value = 2147483647
End Sub

Public Sub ReadValue(ByVal Value As Long, Optional ByVal ValueScale As Long = -2147483648, Optional ByVal MinuteDelimiter As String = "", Optional ByVal SecondDelimiter As String = "")
    m_Value = Value
    m_Scale = ValueScale
    m_MinuteDelimiter = MinuteDelimiter
    m_SecondDelimiter = SecondDelimiter
End Sub

Public Sub Parse(ByVal Expression As String, ByVal MinuteDelimiter As String, ByVal SecondDelimiter As String)
    If (MinuteDelimiter = "") Then
        Err.Raise (5)
    End If
    m_MinuteDelimiter = MinuteDelimiter

    If (SecondDelimiter = "") Then
        Err.Raise (5)
    End If
    m_SecondDelimiter = SecondDelimiter

    Dim rexp As Object: Set rexp = CreateObject("VBScript.RegExp")

    Dim noMinutePattern As String: noMinutePattern = "^(\d{1,})" & EscapeValue(SecondDelimiter) & "(\d{0," & MAX_DIGITS & "})"
    Dim minutePattern As String: minutePattern = "^(\d{1,})" & EscapeValue(MinuteDelimiter) & "(\d{2})" & EscapeValue(SecondDelimiter) & "(\d{0," & MAX_DIGITS & "})"

    Dim targetExpression As String: targetExpression = ""

    rexp.Pattern = noMinutePattern
    If (rexp.Test(Expression)) Then
        m_Scale = Len(rexp.Execute(Expression).Item(0).SubMatches.Item(1))
        m_MinuteDelimiter = ""
        targetExpression = "0" & MinuteDelimiter & rexp.Execute(Expression).Item(0).Value & String(MAX_DIGITS, "0")
    End If

    rexp.Pattern = minutePattern
    If (rexp.Test(Expression)) Then
        m_Scale = Len(rexp.Execute(Expression).Item(0).SubMatches.Item(2))
        targetExpression = rexp.Execute(Expression).Item(0).Value & String(MAX_DIGITS, "0")
    End If

    If (targetExpression = "") Then
        m_Value = INVALID_VALUE
        Exit Sub
    End If

    Dim parsingPattern As String: parsingPattern = "^(\d+)" & EscapeValue(MinuteDelimiter) & "(\d+)" & EscapeValue(SecondDelimiter) & "(\d{" & MAX_DIGITS & "})0{0," & MAX_DIGITS & "}$"
    rexp.Pattern = parsingPattern
    targetExpression = rexp.Execute(targetExpression).Item(0).Value

    Dim minuteValue As Long: minuteValue = CLng(rexp.Replace(targetExpression, "$1"))
    Dim secondValue As Long: secondValue = CLng(rexp.Replace(targetExpression, "$2"))
    Dim miniValue As Long: miniValue = CLng(rexp.Replace(targetExpression, "$3"))

    m_Value = (minuteValue * MINUTE_PER_SECOND + secondValue) * 10 ^ MAX_DIGITS + miniValue

    Set rexp = Nothing
End Sub

Public Property Get Value() As Long
    Value = m_Value
End Property

Public Property Get Valid() As Boolean
    Valid = Not ((m_Value < 0) Or (m_Value = INVALID_VALUE))
End Property

Public Function RoundValue(Digits As Long) As TimeResultValue
    If Not(Me.Valid()) Then
        Set RoundValue = Me
        Exit Function
    End If

    If ((Digits > MAX_DIGITS) Or Digits > m_Scale Or Digits < 0) Then
        Err.Raise(5)
    End If

    Dim addend As Long: addend = IIf(Digits = MAX_DIGITS, 0, 10 ^ (MAX_DIGITS - Digits) - 1)
    Dim timeValueAdded As Long: timeValueAdded = m_Value + addend
    Dim valueUnit As Long: valueUnit = 10 ^ (MAX_DIGITS - Digits)

    Dim roundedValue As Long: roundedValue = timeValueAdded - (timeValueAdded Mod valueUnit)
    Dim rounded As New TimeResultValue: Call rounded.ReadValue(roundedValue, Digits, m_MinuteDelimiter, m_SecondDelimiter)
    Set RoundValue = rounded
End Function

Public Function Compare(ByRef Other As TimeResultValue) As Integer
    If Other Is Nothing Then
        Compare = -1
    Else
        Compare = Sgn(Me.Value - Other.Value)
    End If
End Function

Public Function ToString(Optional ByVal MinuteDelimiter As String = "", Optional ByVal SecondDelimiter As String = "") As String
    If (Not Me.Valid) Then
        ToString = ""
        Exit Function
    End If

    If MinuteDelimiter = "" Then
        MinuteDelimiter = m_MinuteDelimiter
    End If

    If SecondDelimiter = "" Then
        SecondDelimiter = m_SecondDelimiter
    End If

    ToString = ""

    Dim minuteValue As Long: minuteValue = Fix(m_Value / (MINUTE_PER_SECOND * 10 ^ MAX_DIGITS))
    Dim secondValue As Long: secondValue = Fix(m_Value / (10 ^ MAX_DIGITS)) Mod MINUTE_PER_SECOND
    Dim miniValue As Long: miniValue = Fix((m_Value Mod (10 ^ MAX_DIGITS)) / (10 ^ (MAX_DIGITS - m_Scale)))

    ToString = SecondDelimiter
    If (m_Scale > 0) Then
        ' 桁精度が0より大きい場合は秒未満をくっつける
        ToString = ToString & Format(miniValue, String(m_Scale, "0"))
    End If

    If (MinuteDelimiter = "") Then
        ToString = Replace(Str(minuteValue * MINUTE_PER_SECOND + secondValue), " ", "") & ToString
    Else
        If (minuteValue > 0) Then
            ToString = Replace(Str(minuteValue), " ", "") & MinuteDelimiter & Format(secondValue, "00") & ToString
        Else
            ToString = Replace(Str(secondValue), " ", "") & ToString
        End If
    End If
End Function

Private Function EscapeValue(ByVal Expression) As String
    EscapeValue = Expression

    EscapeValue = Replace(EscapeValue, "\", "\\")
    EscapeValue = Replace(EscapeValue, ".", "\.")
    EscapeValue = Replace(EscapeValue, "^", "\^")
    EscapeValue = Replace(EscapeValue, "$", "\$")
    EscapeValue = Replace(EscapeValue, "|", "\|")
    EscapeValue = Replace(EscapeValue, "[", "\[")
    EscapeValue = Replace(EscapeValue, "]", "\]")
    EscapeValue = Replace(EscapeValue, "(", "\(")
    EscapeValue = Replace(EscapeValue, ")", "\)")
    EscapeValue = Replace(EscapeValue, "{", "\{")
    EscapeValue = Replace(EscapeValue, "}", "\}")
    EscapeValue = Replace(EscapeValue, "+", "\+")
    EscapeValue = Replace(EscapeValue, "*", "\*")
    EscapeValue = Replace(EscapeValue, "?", "\?")
End Function

Private Property Get IResultValue_Valid() As Boolean
    IResultValue_Valid = Me.Valid
End Property

Private Property Get IResultValue_Value() As Long
    IResultValue_Value = Me.Value
End Property

Private Function IResultValue_ToString() As String
    IResultValue_ToString = Me.ToString()
End Function

Private Function IResultValue_Compare(ByRef Other As IResultValue) As Integer
    Dim thatValue As TimeResultValue: Set thatValue = Other
    ' タイムは速いほう(m_Valueの値が小さい)が成績が良いので、 -1をかける.
    IResultValue_Compare = -1 * Me.Compare(thatValue)
End Function

Private Sub IResultValue_Parse(ByVal Expression As String, ParamArray Delimiters() As Variant)
    Dim vStartIndex As Long: vStartIndex = LBound(Delimiters)
    Call Me.Parse(Expression, Delimiters(vStartIndex + 0), Delimiters(vStartIndex + 1))
End Sub
