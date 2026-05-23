VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "TrackHtmlOutputer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/module/html/output
Option Explicit

Private m_HtmlOutputer As HtmlOutputer

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    Set m_HtmlOutputer = New HtmlOutputer
    m_Initialized = False
End Sub

Private Sub Class_Terminate()
    Set m_HtmlOutputer = Nothing
End Sub

Public Sub Initialize(HtmlTemplateWriter As IHtmlTemplateWriter)
    If m_Initialized Then Err.Raise(91)
    Call m_HtmlOutputer.Initialize(HtmlTemplateWriter)
    m_Initialized = True
End Sub

Public Function WriteFile(ByVal FileName As String)
    Dim vOutputPath As String: vOutputPath = CompetitionRepository.ReadSettinng.HtmlTemplate.OutputDirectoryPath
    If (Dir(vOutputPath, vbDirectory) = "") Then
        MkDir vOutputPath
    End If

    Dim vOutputFilePath As String: vOutputFilePath = vOutputPath & "\" & FileName
    Call m_HtmlOutputer.WriteFile(vOutputFilePath)

    WriteFile = vOutputFilePath
End Function
