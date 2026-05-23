VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EntryListSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Sub BtnProgramEntry_Clicked()
    On Error GoTo caught_failure_program_entry
        Call MacroProcessorFactory.CreateInstance()

        Dim vTargetRound As String: vTargetRound = Range_TargetRound().Value
        
        Call ProgramEntryAssignerController.ProgramEntries(vTargetRound)

        GoTo finally_program_entry
    On Error GoTo 0
caught_failure_program_entry:
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()

    Select Case Err.Number 
        Case CustomErrorCodeEnum.TargetRoundNotSelected 
            Call MessageFactory.Generate("SE022").ToMsgBox()
        Case CustomErrorCodeEnum.ResultAlreadyRegistred 
            Call MessageFactory.Generate("SE018").ToMsgBox()
        Case Else 
            Call MessageFactory.Generate("SF001").ToMsgBox()
    End Select

    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
finally_program_entry:
    Call MacroProcessorFactory.DestroyInstance()
End Sub

Public Sub BtnWriteStartList_Clicked()
    On Error GoTo caught_failure_writing_start_list
        Call MacroProcessorFactory.CreateInstance()

        Dim vTargetRound As String: vTargetRound = Range_TargetRound().Value
        Call ProgramEntryAssignerController.WriteStartList(vTargetRound)

        Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
        Call MessageFactory.Generate("SI014").ToMsgBox(Range_TargetRound().Value)
        Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
        GoTo finally_program_entry
    On Error GoTo 0
caught_failure_writing_start_list:
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
    Call MessageFactory.Generate("SF001").ToMsgBox()
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
finally_program_entry:
    Call MacroProcessorFactory.DestroyInstance()
End Sub

Public Sub BtnPickUpQualifiers_Clicked()
    On Error GoTo caught_failure_pick_up_qualifiers
        Call MacroProcessorFactory.CreateInstance()

        Dim vTargetRound As String: vTargetRound = Range_TargetRound().Value

        Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
        If (MessageFactory.Generate("SQ010").ToMsgBox() = vbYes) Then
            Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
            Call ProgramEntryAssignerController.PickUpQualifiers(vTargetRound)
        Else
            Dim vTargetNumber As Long: vTargetNumber = Application.InputBox(MessageFactory.Generate("SI011").Prompt, Title:="êièoé“êîì¸óÕ", Type:=1)
            Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
            Call ProgramEntryAssignerController.PickUpQualifiers(vTargetRound, vTargetNumber)
        End If
        GoTo finally_pick_up_qualifiers
    On Error GoTo 0
caught_failure_pick_up_qualifiers:
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating

    Select Case Err.Number
    Case CustomErrorCodeEnum.QualifiersNotAssigned
        MsgBox Err.Description, vbOKOnly + vbExclamation, MessageFactory.Generate("SW012").Title
    Case Else
        Call MessageFactory.Generate("SF001").ToMsgBox()
    End Select

    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating
finally_pick_up_qualifiers:
    Call MacroProcessorFactory.DestroyInstance()
End Sub

Private Function Range_TargetRound() As Range
    Set Range_TargetRound = Me.Cells(2, 3)
End Function
