VERSION 1.0 CLASS
BEGIN
MultiUse = -1 ' True
END
Attribute VB_Name = "MessageModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/common/message
Option Explicit
Implements IJsonConvertable

Private m_Code As String
Private m_MessageType As String
Private m_Title As String
Private m_Prompt As String

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
End Sub

Private Sub Class_Terminate()
End Sub

Public Sub Initialize(Code As String, MessageType As String, Title As String, Prompt As String)
    If m_Initialized Then Err.Raise(91)
    m_Code = Code
    m_MessageType = MessageType
    m_Title = Title
    m_Prompt = Prompt
    m_Initialized = True
End Sub

Public Property Get Code() As String
    If Not m_Initialized Then Err.Raise(91)
    Code = m_Code
End Property

Public Property Get MessageType() As String
    If Not m_Initialized Then Err.Raise(91)
    MessageType = m_MessageType
End Property

Public Property Get Title() As String
    If Not m_Initialized Then Err.Raise(91)
    Title = m_Title
End Property

Public Property Get Prompt(ParamArray Arguments() As Variant) As String
    If Not m_Initialized Then Err.Raise(91)
    Dim vReplacedPrompt As String: vReplacedPrompt = m_Prompt

    Dim pArguments() As Variant
    Dim loBound As Long: loBound = LBound(Arguments)
    If (loBound > UBound(Arguments)) Then
        ' Empty Arguments
        pArguments = Arguments
    ElseIf (IsArray(Arguments(loBound))) Then 
        pArguments = Arguments(loBound)
        loBound = LBound(pArguments())
    Else
        pArguments = Arguments
    End If

    Dim i As Long
    For i = 0 To UBound(pArguments) - loBound
        vReplacedPrompt = Replace(vReplacedPrompt, "{" & i & "}", pArguments(i + loBound))
    Next i

    Prompt = vReplacedPrompt
End Property

Public Function ToMsgBox(ParamArray Arguments() As Variant)
    Dim pMessageType As String: pMessageType = Me.MessageType
    Select Case pMessageType 
        Case "FATAL", "ERROR" 
            ToMsgBox = MsgBox(Me.Prompt(Arguments), vbCritical + vbOKOnly, Me.Title)
        Case "WARN" 
            ToMsgBox = MsgBox(Me.Prompt(Arguments), vbExclamation + vbOKOnly, Me.Title)
        Case "INFO" 
            ToMsgBox = MsgBox(Me.Prompt(Arguments), vbInformation + vbOKOnly, Me.Title)
        Case "QUESTION" 
            ToMsgBox = MsgBox(Me.Prompt(Arguments), vbQuestion + vbYesNo, Me.Title)
        Case Else 
            ' ESCÉLÅ[ÇâüÇµÇΩèÍçáÇ∆ìôâøÇ∆Ç∑ÇÈ.
            ToMsgBox = vbCancel
    End Select
End Function

Public Function Self() 
    If Not m_Initialized Then Err.Raise(91)
    Set Self = Me
End Function

Private Function IJsonConvertable_ToJson() As String
    Dim jsonDictionary As Object: Set jsonDictionary = CreateObject("Scripting.Dictionary")
    Call jsonDictionary.Add("code", Me.Code)
    Call jsonDictionary.Add("message_type", Me.MessageType)
    Call jsonDictionary.Add("title", Me.Title)
    Call jsonDictionary.Add("prompt", Me.Prompt)
    IJsonConvertable_ToJson = JsonConverter.ConvertToJson(jsonDictionary)
End Function

Public Function ToJson() As String
    ToJson = IJsonConvertable_ToJson()
End Function
