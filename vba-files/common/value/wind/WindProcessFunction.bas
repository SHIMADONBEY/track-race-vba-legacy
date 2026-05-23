Attribute VB_Name = "WindProcessFunction"
'namespace=vba-files/common/value/wind
Option Explicit
Option Private Module

Public Function ConvertWind(Expression As Variant) As WindValue
    If (Expression = "" Or IsEmpty(Expression)) Then
        Set ConvertWind = Nothing
        Exit Function
    End If

    Dim putValue As WindValue: Set putValue = New WindValue
    Call putValue.Initialize(CDec(Expression))
    Set ConvertWind = putValue
End Function

Public Function WIND_2() As WindValue
    Set WIND_2 = ConvertWind("2.0")
End Function
