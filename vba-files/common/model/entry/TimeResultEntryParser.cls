VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "TimeResultEntryParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/common/model/entry
Option Explicit
Implements IJsonParser

Private m_MinuteDelimiter As String
Private m_SecondDelimiter As String

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
End Sub

Private Sub Class_Terminate()
End Sub

Public Sub Initialize(MinuteDelimiter As String, SecondDelimiter As String)
    If m_Initialized Then Err.Raise 91, "TimeResultEntryParser.Initialize"
    m_MinuteDelimiter = MinuteDelimiter
    m_SecondDelimiter = SecondDelimiter
    m_Initialized = True
End Sub

Private Function IJsonParser_Parse(JsonString As String, ByVal NewObject As Object) As Object
    If Not m_Initialized Then Err.Raise 91, "TimeResultEntryParser.Parse"

    Dim pEntryObject As PersonalEntryModel: Set pEntryObject = NewObject
    With JsonConverter.ParseJson(JsonString)
        Dim vQualifiedValues As Collection: Set vQualifiedValues = New Collection

        Dim vResultExpression As Variant
        For Each vResultExpression In .Item("qualified_values")
            Dim vQualifiedValue As IResultValue: Set vQualifiedValue = New TimeResultValue
            Call vQualifiedValue.Parse(vResultExpression, m_MinuteDelimiter, m_SecondDelimiter)
            Call vQualifiedValues.Add(vQualifiedValue)
        Next vResultExpression

        Call pEntryObject.Initialize( _
                .Item("personal_id") _
                , .Item("entry_id") _
                , .Item("event_id") _
                , .Item("demo_entry") _
                , vQualifiedValues.Item(1) _
                , vQualifiedValues.Item(2) _
                , vQualifiedValues.Item(3) _
        )
    End With

    Set IJsonParser_Parse = pEntryObject
End Function
