VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "GameRecordSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Call MacroProcessorFactory.CreateInstance()
    On Error GoTo on_worksheet_change_failed
        Dim changedCell As Range
        For Each changedCell In Target.Cells
            Call GameRecordController.OnUpdatedRecord(changedCell)
        Next changedCell 
    GoTo finally_worksheet_change
on_worksheet_change_failed:
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
    Call MessageFactory.Generate("SF001").ToMsgBox()
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
finally_worksheet_change:
    Call MacroProcessorFactory.DestroyInstance()
End Sub

Private Sub Worksheet_Deactivate()
    Call MacroProcessorFactory.CreateInstance()
    On Error GoTo on_worksheet_deactivate_failed
        Call GameRecordController.UpdateGameRecords()
    GoTo finally_worksheet_deactivate
on_worksheet_deactivate_failed:
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
    Call MessageFactory.Generate("SF001").ToMsgBox()
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
finally_worksheet_deactivate:
    Call MacroProcessorFactory.DestroyInstance()
End Sub
