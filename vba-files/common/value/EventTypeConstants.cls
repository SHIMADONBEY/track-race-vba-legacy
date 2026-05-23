VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "EventTypeConstants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'namespace=vba-files/common/value
Option Explicit

Public Function TRACK() As String
    TRACK = "TRK"
End Function

Public Function RELAY() As String
    RELAY = "REL"
End Function

Public Function HEIGHT_JUMP() As String
    HEIGHT_JUMP = "HIG"
End Function

Public Function LENGTH_JUMP() As String
    LENGTH_JUMP = "LNG"
End Function

Public Function THROW() As String
    THROW = "THR"
End Function

Public Function COMBINED_EVENTS() As String
    COMBINED_EVENTS = "DEC"
End Function

Public Function ROAD() As String
    ROAD = "ROD"
End Function
