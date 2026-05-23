VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HtmlTemplatePathsModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/module/event/model
Option Explicit
Private m_StartListTemplatePath As String
Private m_ResultTemlatePath As String
Private m_EntryListTemplatePath As String
Private m_TrackRankingListPath As String
Private m_OutputDirectoryPath As String

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
    ' TODO: Write Process To Initialize
End Sub

Private Sub Class_Terminate()
    ' TODO: Write Process To Terminate
End Sub

Public Sub Initialize(StartListTemplatePath As String, ResultTemlatePath As String, EntryListTemplatePath As String, TrackRankingListPath As String, OutputDirectoryPath As String)
    If m_Initialized Then Err.Raise (91)
    m_StartListTemplatePath = StartListTemplatePath
    m_ResultTemlatePath = ResultTemlatePath
    m_EntryListTemplatePath = EntryListTemplatePath
    m_TrackRankingListPath = TrackRankingListPath
    m_OutputDirectoryPath = OutputDirectoryPath
    m_Initialized = True
End Sub

Public Property Get StartListTemplatePath() As String
    If Not m_Initialized Then Err.Raise (91)
    StartListTemplatePath = m_StartListTemplatePath
End Property

Public Property Get ResultTemlatePath() As String
    If Not m_Initialized Then Err.Raise (91)
    ResultTemlatePath = m_ResultTemlatePath
End Property

Public Property Get EntryListTemplatePath() As String
    If Not m_Initialized Then Err.Raise (91)
    EntryListTemplatePath = m_EntryListTemplatePath
End Property

Public Property Get TrackRankingListPath() As String
    If Not m_Initialized Then Err.Raise (91)
    TrackRankingListPath = m_TrackRankingListPath
End Property

Public Property Get OutputDirectoryPath() As String
    If Not m_Initialized Then Err.Raise (91)
    OutputDirectoryPath = m_OutputDirectoryPath
End Property

Public Function Self() As HtmlTemplatePathsModel
    If Not m_Initialized Then Err.Raise(91)
    Set Self = Me
End Function