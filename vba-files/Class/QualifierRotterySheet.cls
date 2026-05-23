VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QualifierRotterySheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Sub BtnQualifyAllClicked()
    On Error GoTo on_btn_qualify_all_clicked_failed
    Call MacroProcessorFactory.CreateInstance()
    Call QualifierDrawerController.QualifyAll(Rng_TargetRoundId.Value)

    ResultInputController.Get_ThisSheet().Select
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
    Call MessageFactory.Generate("SI006").ToMsgBox(Me.Cells(2, 5).Value)
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()

    GoTo finally_btn_qualify_all_clicked
on_btn_qualify_all_clicked_failed:
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
    Call MessageFactory.Generate("SF001").ToMsgBox()
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
finally_btn_qualify_all_clicked:
    On Error GoTo 0
    Call MacroProcessorFactory.DestroyInstance()
End Sub

Public Sub BtnQualifyClicked()
    If (Rng_TargetNumber.Value <> Rng_SelectedCount.Value) Then
        Call MessageFactory.Generate("SE008").ToMsgBox(Rng_TargetNumber.Value, Rng_SelectedCount.Value)
        Exit Sub
    End If

    On Error GoTo on_btn_qualify_clicked_failed
    Call MacroProcessorFactory.CreateInstance()
    Call QualifierDrawerController.SelectQualifiers(Rng_TargetRoundId.Value, Rng_TargetNumber.Value)

    ResultInputController.Get_ThisSheet().Select
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
    Call MessageFactory.Generate("SI006").ToMsgBox(Me.Cells(2, 5).Value)
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()

    Goto finally_btn_qualify_clicked_failed
on_btn_qualify_clicked_failed:
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
    Call MessageFactory.Generate("SF001").ToMsgBox()
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
finally_btn_qualify_clicked_failed:
    Call MacroProcessorFactory.DestroyInstance()
End Sub

Private Function Rng_TargetRoundId() As Range
    Set Rng_TargetRoundId = Cells(2, 6)
End Function

Private Function Rng_TargetNumber() As Range
    Set Rng_TargetNumber = Cells(3, 5)
End Function

Private Function Rng_SelectedCount() As Range
    Set Rng_SelectedCount = Cells(9, 2)
End Function
