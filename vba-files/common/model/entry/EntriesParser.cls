VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "EntriesParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/common/model/entry
Option Explicit
Implements IJsonParser

Private m_Parser As IJsonParser

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
End Sub

Private Sub Class_Terminate()
End Sub

Public Sub Initialize(Parser As IJsonParser)
    If m_Initialized Then Err.Raise 91, "EntriesParser.Initialize"
    Set m_Parser = Parser
    m_Initialized = True
End Sub

Private Function IJsonParser_Parse(JsonString As String, ByVal NewObject As Object) As Object
    If Not m_Initialized Then Err.Raise 91, "EntriesParser.Parse"

    Dim pParseObject As PersonalEntryModels: Set pParseObject = NewObject
    Dim vJsonElement As Object
    For Each vJsonElement In JsonConverter.ParseJson(JsonString)
        Call pParseObject.Add(JsonConverter.ParseJsonTo(JsonConverter.ConvertToJson(vJsonElement), New PersonalEntryModel, m_Parser))
    Next vJsonElement 

    Set IJsonParser_Parse = pParseObject
End Function
