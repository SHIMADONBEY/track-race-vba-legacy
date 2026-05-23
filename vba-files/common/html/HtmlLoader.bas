Attribute VB_Name = "HtmlLoader"
'namespace=vba-files/common/html
Option Explicit
Option Private Module

Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, _
    ByVal Operation As String, _
    ByVal Filename As String, _
    Optional ByVal Parameters As String, _
    Optional ByVal Directory As String, _
    Optional ByVal WindowStyle As Long = vbMinimizedFocus _
) As Long

Public Sub OpenHtmlFile(FilePath As String)
    Dim vExecutedResult As Long: vExecutedResult = ShellExecute(0, "open", FilePath)
    If (vExecutedResult <= 32) Then
        Debug.Print "Shell Execution Error: " & vExecutedResult
        Err.Source = "HtmlLoader.OpenHtmlFile"
        Select Case vExecutedResult 
            Case 0 
                Err.Raise 7
            Case 2, 3 
                Err.Raise 53
            Case Else
                Err.Raise 51
        End Select
    End If
End Sub
