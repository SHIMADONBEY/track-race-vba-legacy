VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "CodeMasterRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'namespace=vba-files/track-race/Repository
Option Explicit

Private m_SexCode As Object
Private m_WindScalingCode As Object
Private m_ReactionTimeCode As Object
Private m_ResultUnitCode As Object
Private m_UsingScoreCode As Object
Private m_GroupStrategyCode As Object
Private m_OrderStrategyCode As Object

Public Property Get SexCode(Optional Reload As Boolean = False) As Object
    If (Reload Or m_SexCode Is Nothing) Then
        Set m_SexCode = CreateObject("Scripting.Dictionary")
        Dim vRowRange As Range
        For Each vRowRange In Range("Code_Sex").Rows()
            Call m_SexCode.Add(vRowRange.Cells(1, 1).Value, vRowRange.Cells(1, 2).Value)
        Next vRowRange 
    End If

    Set SexCode = m_SexCode
End Property

Public Property Get WindScalingCode(Optional Reload As Boolean = False) As Object
    If (Reload Or m_WindScalingCode Is Nothing) Then
        Set m_WindScalingCode = CreateObject("Scripting.Dictionary")
        Dim vRowRange As Range
        For Each vRowRange In Range("Code_Wind").Rows()
            Call m_WindScalingCode.Add(vRowRange.Cells(1, 1).Value, vRowRange.Cells(1, 2).Value)
        Next vRowRange 
    End If

    Set WindScalingCode = m_WindScalingCode
End Property

Public Property Get ReactionTimeCode(Optional Reload As Boolean = False) As Object
    If (Reload Or m_ReactionTimeCode Is Nothing) Then
        Set m_ReactionTimeCode = CreateObject("Scripting.Dictionary")
        Dim vRowRange As Range
        For Each vRowRange In Range("Code_Reaction").Rows()
            Call m_ReactionTimeCode.Add(vRowRange.Cells(1, 1).Value, vRowRange.Cells(1, 2).Value)
        Next vRowRange 
    End If

    Set ReactionTimeCode = m_ReactionTimeCode
End Property

Public Property Get ResultUnitCode(Optional Reload As Boolean = False) As Object
    If (Reload Or m_ResultUnitCode Is Nothing) Then
        Set m_ResultUnitCode = CreateObject("Scripting.Dictionary")
        Dim vRowRange As Range
        For Each vRowRange In Range("Code_Result_Unit").Rows()
            Call m_ResultUnitCode.Add(vRowRange.Cells(1, 1).Value, vRowRange.Cells(1, 2).Value)
        Next vRowRange 
    End If

    Set ResultUnitCode = m_ResultUnitCode
End Property

Public Property Get UsingScoreCode(Optional Reload As Boolean = False) As Object
    If (Reload Or m_UsingScoreCode Is Nothing) Then
        Set m_UsingScoreCode = CreateObject("Scripting.Dictionary")
        Dim vRowRange As Range
        For Each vRowRange In Range("Code_Point_1").Rows()
            Call m_UsingScoreCode.Add(vRowRange.Cells(1, 1).Value, vRowRange.Cells(1, 2).Value)
        Next vRowRange 
    End If

    Set UsingScoreCode = m_UsingScoreCode
End Property

Public Property Get GroupStrategyCode(Optional Reload As Boolean = False) As Object
    If (Reload Or m_GroupStrategyCode Is Nothing) Then
        Set m_GroupStrategyCode = CreateObject("Scripting.Dictionary")
        Dim vRowRange As Range
        For Each vRowRange In Range("Code_Group_Strategy").Rows()
            Call m_GroupStrategyCode.Add(vRowRange.Cells(1, 1).Value, vRowRange.Cells(1, 2).Value)
        Next vRowRange 
    End If

    Set GroupStrategyCode = m_GroupStrategyCode
End Property

Public Property Get OrderStrategyCode(Optional Reload As Boolean = False) As Object
    If (Reload Or m_OrderStrategyCode Is Nothing) Then
        Set m_OrderStrategyCode = CreateObject("Scripting.Dictionary")
        Dim vRowRange As Range
        For Each vRowRange In Range("Code_Lane_Strategy").Rows()
            Call m_OrderStrategyCode.Add(vRowRange.Cells(1, 1).Value, vRowRange.Cells(1, 2).Value)
        Next vRowRange 
    End If

    Set OrderStrategyCode = m_OrderStrategyCode
End Property

Public Sub Refresh()
    Dim vCodeDictionary As Object
    Set vCodeDictionary = Me.SexCode(True)
    Set vCodeDictionary = Me.WindScalingCode(True)
    Set vCodeDictionary = Me.ReactionTimeCode(True)
    Set vCodeDictionary = Me.ResultUnitCode(True)
    Set vCodeDictionary = Me.UsingScoreCode(True)
    Set vCodeDictionary = Me.GroupStrategyCode(True)
    Set vCodeDictionary = Me.OrderStrategyCode(True)
End Sub
