VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HtmlTemplateWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/common/html/input
Option Explicit

Private m_HtmlText As String
Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
    ' TODO: Write Process To Initialize
End Sub

Private Sub Class_Terminate()
    ' TODO: Write Process To Terminate
End Sub

Public Sub Initialize(TemplateFilePath As String)
    If m_Initialized Then Err.Raise (91)

    Dim fileStream As Object: Set fileStream = CreateObject("ADODB.Stream")

    With fileStream
        .Charset = "UTF-8"
        Call .Open
        Call .LoadFromFile(TemplateFilePath)
        m_HtmlText = .ReadText
        Call .Close
    End With

    m_Initialized = True
End Sub

'Public Property Get HtmlCollection() As Object
Public Property Get HtmlCollection() As Object
    If Not m_Initialized Then Err.Raise (91)

    Set HtmlCollection = CreateObject("htmlfile")
    Call HtmlCollection.Write(m_HtmlText)
End Property
