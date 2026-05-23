VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "IResultValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/common/value
Option Explicit

Public Property Get Valid() As Boolean
End Property

Public Property Get Value() As Long
End Property

Public Function ToString() As String
End Function

Public Function Compare(ByRef Other As IResultValue) As Integer
End Function

Public Sub Parse(ByVal Expression As String, ParamArray Delimiters() As Variant)
End Sub

' (Templates Start)============================================================
' Implements IResultValue

' Private Property Get IResultValue_Valid() As Boolean
' End Property

' Private Property Get IResultValue_Value() As Long
' End Property

' Private Function IResultValue_ToString() As String
' End Function

' Private Function IResultValue_Compare(ByRef Other As IResultValue) As Integer
' End Function

' Private Sub IResultValue_Parse(ByVal Expression As String, ParamArray Delimiters() As Variant)
' End Sub
' (Templates End)============================================================
