VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventSettingSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Call MacroProcessorFactory.CreateInstance
    On Error GoTo on_cell_changed_failed
        Dim vSpotCell As Range
        For Each vSpotCell In Target
            Call EventSettingController.OnRoundCountUpdated(vSpotCell)
            Call EventSettingController.OnLaneCountUpdated(vSpotCell)
        Next vSpotCell
    GoTo finally_cell_changed
on_cell_changed_failed:
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()

    Select Case Err.Number
        Case CustomErrorCodeEnum.InvalidRoundCount
            Call MessageFactory.Generate("SE002").ToMsgBox()
        Case Else 
            Call MessageFactory.Generate("SF001").ToMsgBox()
    End Select

    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
finally_cell_changed:
    Call MacroProcessorFactory.DestroyInstance
End Sub

Private Sub Worksheet_Deactivate()
    Call MacroProcessorFactory.CreateInstance
    On Error GoTo on_failed_deactivation
        Call EventSettingController.UpdateEventData
        Call EventSettingController.UpdateEventRounds
    GoTo finally_deactivation
on_failed_deactivation:
    Me.Activate
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating

    Select Case Err.Number 
        Case CustomErrorCodeEnum.DuplicateRoundName
            MsgBox Err.Description, vbExclamation + vbOKOnly, MessageFactory.Generate("SW017").Title
        Case Else 
            Call MessageFactory.Generate("SF001").ToMsgBox()
    End Select

    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating
finally_deactivation:
    Call MacroProcessorFactory.DestroyInstance
End Sub
