VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResultInputSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Sub BtnGroupSelect_Clicked()

    On Error GoTo on_group_selection_failed

    Cells(1, 1).Select
    ActiveWindow.FreezePanes = False
    Dim mp As MacroProcessor: Set mp = MacroProcessorFactory.GetInstance()

    Dim vRoundName As String    : vRoundName = ResultInputController.Range_TargetRoundNameInput().Value
    Dim vGroupNumber As Long    : vGroupNumber = ResultInputController.Range_TargetGroupInput().Value

    If Not (ResultInputController.ChangingGroup(vRoundName, vGroupNumber)) Then
        GoTo finally_BtnGroupSelect_Clicked
    End If

    ' èàóù
    Call ResultInputController.InitializeResultList(vRoundName, vGroupNumber)

    Dim loadingResult As Boolean: loadingResult = ResultInputController.ResultHasRegistered()
    If (Not loadingResult) Then
        Call mp.SwitchScreenUpdating()
        loadingResult = (MessageFactory.Generate("SQ004").ToMsgBox() = vbYes)
        Call mp.SwitchScreenUpdating()
    End If

    Call ResultInputController.ReadResults(vRoundName, vGroupNumber, loadingResult)

    On Error GoTo 0
    GoTo finally_BtnGroupSelect_Clicked
on_group_selection_failed:
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
    Select Case Err.Number 
        Case CustomErrorCodeEnum.NoEntryAssigned 
            MsgBox Err.Description, vbOKOnly + vbCritical, MessageFactory.Generate("SE021").Title  
        Case Else 
            Call MessageFactory.Generate("SF001").ToMsgBox()
    End Select    
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
finally_BtnGroupSelect_Clicked:
    Set mp = Nothing
    Call MacroProcessorFactory.DestroyInstance()
    Cells(11, 1).Select
    ActiveWindow.FreezePanes = True
    Cells(11, 2).Select
End Sub

Public Sub BtnRegisterResult_Clicked()
    Dim mp As MacroProcessor: Set mp = MacroProcessorFactory.GetInstance()
    On Error GoTo on_register_result_failed 

    Dim vRoundName As String    : vRoundName = ResultInputController.Range_TargetRoundNameInput().Value
    Dim vGroupNumber As Long : vGroupNumber = ResultInputController.Range_TargetGroupInput().Value

    Call ResultInputController.RegisterResults()
    Call ResultInputController.InitializeResultList(vRoundName, vGroupNumber)

    Call mp.SwitchScreenUpdating()
    Call MessageFactory.Generate("SI015").ToMsgBox(vRoundName, vGroupNumber)
    Call mp.SwitchScreenUpdating()

    Call ResultInputController.ReadResults(vRoundName, vGroupNumber, true)

    On Error GoTo 0
    GoTo finally_BtnRegisterResult_Clicked
on_register_result_failed:
    Call mp.SwitchScreenUpdating()
    Select Case Err.Number 
        Case CustomErrorCodeEnum.ResultOrderDuplicated 
            MsgBox Err.Description, vbOKOnly + vbCritical, MessageFactory.Generate("SE020").Title
        Case Else 
            Call MessageFactory.Generate("SF001").ToMsgBox()
    End Select
    Call mp.SwitchScreenUpdating()
finally_BtnRegisterResult_Clicked:
    Set mp = Nothing
    Call MacroProcessorFactory.DestroyInstance()
End Sub

Public Sub BtnRanking_Clicked()
    If Not (MessageFactory.Generate("SQ009").ToMsgBox() = vbYes) Then
        Exit Sub
    End If

    Dim mp As MacroProcessor: Set mp = MacroProcessorFactory.GetInstance()
    On Error GoTo on_btn_ranking_cliecked_failed
        Dim vRoundName As String: vRoundName = ResultInputController.Range_TargetRoundNameInput().Value
        Call ResultInputController.CalculateRanking(vRoundName)

        Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
        Call MessageFactory.Generate("SI006").ToMsgBox(vRoundName)
        Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
    On Error GoTo 0
    GoTo finally_btn_ranking_cliecked
on_btn_ranking_cliecked_failed:
    Call mp.SwitchScreenUpdating()
    Call MessageFactory.Generate("SF001").ToMsgBox()
    Call mp.SwitchScreenUpdating()
finally_btn_ranking_cliecked:
    Set mp = Nothing
    Call MacroProcessorFactory.DestroyInstance()
End Sub

Public Sub BtnRankingList_Clicked()
    Dim mp As MacroProcessor: Set mp = MacroProcessorFactory.GetInstance()
    On Error GoTo on_btn_ranking_list_cliecked_failed
        Dim vRoundName As String: vRoundName = ResultInputController.Range_TargetRoundNameInput().Value
        Call ResultInputController.WriteRanking(vRoundName)

        Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
        Call MessageFactory.Generate("SI016").ToMsgBox(vRoundName)
        Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
    On Error GoTo 0
    GoTo finally_btn_ranking_list_cliecked
on_btn_ranking_list_cliecked_failed:
    Call mp.SwitchScreenUpdating()
    Call MessageFactory.Generate("SF001").ToMsgBox()
    Call mp.SwitchScreenUpdating()
finally_btn_ranking_list_cliecked:
    Set mp = Nothing
    Call MacroProcessorFactory.DestroyInstance()
End Sub

Public Sub BtnResultList_Clicked()
    Dim mp As MacroProcessor: Set mp = MacroProcessorFactory.GetInstance()
    On Error GoTo on_btn_ranking_cliecked_failed

    Dim vRoundName As String: vRoundName = ResultInputController.Range_TargetRoundNameInput().Value
    Call ResultInputController.WriteResultList(vRoundName)
    Debug.Print "Updated"
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
    Call MessageFactory.Generate("SI013").ToMsgBox(vRoundName)
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
    On Error GoTo 0
    GoTo finally_btn_ranking_cliecked
on_btn_ranking_cliecked_failed:
    Call mp.SwitchScreenUpdating()
    Call MessageFactory.Generate("SF001").ToMsgBox()
    Call mp.SwitchScreenUpdating()
finally_btn_ranking_cliecked:
    Set mp = Nothing
    Call MacroProcessorFactory.DestroyInstance()
End Sub

Public Sub BtnCsvReport_Clicked()
    Call MacroProcessorFactory.CreateInstance()
    On Error GoTo on_btn_csv_report_failed
        Call InternalCsvExporter.WriteCsv()
        Call MessageFactory.Generate("SI018").ToMsgBox()
    GoTo finally_btn_csv_report
on_btn_csv_report_failed:
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()

    Select Case Err.Number 
        Case CustomErrorCodeEnum.NoEntryAssigned
            MsgBox Err.Description, vbOKOnly + vbCritical, MessageFactory.Generate("SE019").Title
        Case Else 
            Call MessageFactory.Generate("SF001").ToMsgBox()
    End Select

    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
finally_btn_csv_report:
    Call MacroProcessorFactory.DestroyInstance()
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim changedCell As Range
    Dim mp As MacroProcessor: Set mp = MacroProcessorFactory.GetInstance()
    On Error Goto on_worksheet_change_failed
    For Each changedCell In Target.Cells
        If (ResultInputController.OnUpdatedResultData(changedCell)) Then
            If (ResultInputController.OnUpdateOrder(changedCell)) Then 
                Debug.Print "Processed: " & changedCell.Address
            ElseIf(ResultInputController.OnUpdatedRealResult(changedCell)) Then
                Debug.Print "Processed: " & changedCell.Address
            Else
                Debug.Print "Processed: " & changedCell.Address
            End If
        End If
    Next changedCell 
    On Error Goto 0
    Goto finally_worksheet_change_failed
on_worksheet_change_failed:
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
    Call MessageFactory.Generate("SF001").ToMsgBox()
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
finally_worksheet_change_failed:
    Set mp = Nothing
    Call MacroProcessorFactory.DestroyInstance()
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Call MacroProcessorFactory.CreateInstance()
    On Error GoTo on_failed_selection_change
        Call ResultInputController.OnWindInputSelect(Target)
    GoTo finally_selection_change
on_failed_selection_change:
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()
    Call MessageFactory.Generate("SF001").ToMsgBox()
    Call MacroProcessorFactory.GetInstance().SwitchScreenUpdating()    
finally_selection_change:
    Call MacroProcessorFactory.DestroyInstance()
End Sub
