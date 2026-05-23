VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HtmlResultRankingWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'namespace=vba-files/track-race/modules/html/output/result-ranking
Option Explicit
Implements IHtmlTemplateWriter

'Private m_Html As Object
Private m_Html As Object
Private m_OutputData As ResultRankingOutputModel
Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
End Sub

Private Sub IHtmlTemplateWriter_WriteHtml()
    Call WriteHtml()
End Sub

Private Property Get IHtmlTemplateWriter_HtmlDocument() As Object
    If Not m_Initialized Then Err.Raise(91)
    Set IHtmlTemplateWriter_HtmlDocument = m_Html
End Property

Public Sub Initialize(ByRef HtmlDoc As Object, ByRef OutputData As ResultRankingOutputModel)
    Set m_Html = HtmlDoc
    Set m_OutputData = OutputData
    m_Initialized = True
End Sub

Public Function WriteHtml() As Object
    
    Dim jsonElement As Object: Set jsonElement = m_Html.getElementById("json-param")
    
    jsonElement.InnerText = m_OutputData.ToJson()

    Set WriteHtml = m_Html
End Function
