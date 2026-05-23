VERSION 1.0 CLASS
  BEGIN
  MultiUse = -1  'True
  END
  Attribute VB_Name = "HtmlOutputer"
  Attribute VB_GlobalNameSpace = False
  Attribute VB_Creatable = False
  Attribute VB_PredeclaredId = False
  Attribute VB_Exposed = False
'namespace=vba-files/common/html/output
Option Explicit
Private m_HtmlTemplateWriter As IHtmlTemplateWriter

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
    ' TODO: Write Process To Initialize
End Sub

Private Sub Class_Terminate()
    ' TODO: Write Process To Terminate
End Sub

Public Sub Initialize(HtmlTemplateWriter As IHtmlTemplateWriter)
    If m_Initialized Then Err.Raise(91)
    Set m_HtmlTemplateWriter = HtmlTemplateWriter
    m_Initialized = True
End Sub

Public Property Get HtmlTemplateWriter() As IHtmlTemplateWriter
    If Not m_Initialized Then Err.Raise(91)
    Set HtmlTemplateWriter = m_HtmlTemplateWriter
End Property

Public Sub WriteFile(ByVal FileName As String)
    Call m_HtmlTemplateWriter.WriteHtml()

    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        Call .Open
        Call .WriteText(m_HtmlTemplateWriter.HtmlDocument.documentElement.OuterHtml)
        Call .SaveToFile(FileName, 2)
        Call .Close
    End With
End Sub
