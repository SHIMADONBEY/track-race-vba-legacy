VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "IObjectComparator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/common/collection
Option Explicit

'/*
'IObjectComparator#Compare()
'Compare two objects.
'
'**Syntax**
'
'+ LeftHandObject: An Object  
'+ RightHandObject: An Object
'
'**Return**
'A negative integer if the first argument is less than the second argument, 0 if both are equal, 
'and a positive integer if the first argument is greater than the second argument.
'*/
Public Function Compare(LeftHandObject As Object, RightHandObject As Object) As Integer
End Function

' (Template Start) ============================================================
' Implements IObjectComparator

' Private  Function IObjectComparator_Compare(LeftHandObject As Object, RightHandObject As Object) As Integer
' End Function

' (Template End) ==============================================================
