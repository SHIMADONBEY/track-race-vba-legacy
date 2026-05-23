Attribute VB_Name = "StringUtil"
'namespace=vba-files/common/util
Option Explicit
Option Private Module

Public Function JoinArrayToString(Expression As Variant, Delimiter As String) As String
    Dim i As Long
    Dim first As Boolean: first = True
    For i = LBound(Expression) To UBound(Expression)
        If Not(first) Then
            JoinArrayToString = JoinArrayToString & Delimiter
        End If
        
        JoinArrayToString = JoinArrayToString & Expression(i)
        first = first And False
    Next i
End Function

Public Function JoinCollectionToString(Expression As Collection, Delimiter As String) As String
    Dim i As Long
    Dim first As Boolean: first = True
    For i = 1 To Expression.Count()
        If Not(first) Then
            JoinCollectionToString = JoinCollectionToString & Delimiter
        End If
        
        JoinCollectionToString = JoinCollectionToString & Expression.Item(i)
        first = first And False
    Next i
End Function

Public Function ConvertFormatString(Expression As String) As String
    ConvertFormatString = Expression
    ConvertFormatString = Replace(ConvertFormatString, "\", "\\")
    ConvertFormatString = Replace(ConvertFormatString, "a", "\a")
    ConvertFormatString = Replace(ConvertFormatString, "c", "\c")
    ConvertFormatString = Replace(ConvertFormatString, "d", "\d")
    ConvertFormatString = Replace(ConvertFormatString, "e", "\e")
    ConvertFormatString = Replace(ConvertFormatString, "E", "\E")
    ConvertFormatString = Replace(ConvertFormatString, "h", "\h")
    ConvertFormatString = Replace(ConvertFormatString, "m", "\m")
    ConvertFormatString = Replace(ConvertFormatString, "n", "\n")
    ConvertFormatString = Replace(ConvertFormatString, "p", "\p")
    ConvertFormatString = Replace(ConvertFormatString, "q", "\q")
    ConvertFormatString = Replace(ConvertFormatString, "s", "\s")
    ConvertFormatString = Replace(ConvertFormatString, "t", "\t")
    ConvertFormatString = Replace(ConvertFormatString, "w", "\w")
    ConvertFormatString = Replace(ConvertFormatString, "y", "\y")
    ConvertFormatString = Replace(ConvertFormatString, "/", "\/")
    ConvertFormatString = Replace(ConvertFormatString, ":", "\:")
    ConvertFormatString = Replace(ConvertFormatString, "#", "\#")
    ConvertFormatString = Replace(ConvertFormatString, "%", "\%")
    ConvertFormatString = Replace(ConvertFormatString, ",", "\,")
    ConvertFormatString = Replace(ConvertFormatString, ".", "\.")
    ConvertFormatString = Replace(ConvertFormatString, "@", "\@")
    ConvertFormatString = Replace(ConvertFormatString, "&", "\")
    ConvertFormatString = Replace(ConvertFormatString, "<", "\<")
    ConvertFormatString = Replace(ConvertFormatString, ">", "\>")
    ConvertFormatString = Replace(ConvertFormatString, "!", "\!")
    ConvertFormatString = Replace(ConvertFormatString, """", "\""")
End Function
