VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "TrackGameRecordJsonParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'namespace=vba-files/track-race/model/game-record
Option Explicit

Implements IJsonParser

Private Function IJsonParser_Parse(JsonString As String, ByVal NewObject As Object) As Object
    Debug.Print JsonString
    Dim pNewObject As GameRecordModel: Set pNewObject = NewObject

    With JsonConverter.ParseJson(JsonString)
        Dim vRecord As IResultValue: Set vRecord = TimeProcessFunction.ParseTimeResultValue(.Item("record"))
        Call pNewObject.Initialize( _
                .Item("event_id") _
                , .Item("serial") _
                , .Item("name") _
                , .Item("abbrebiation") _
                , vRecord _
                , .Item("athlete_name") _
                , .Item("team_name") _
                , .Item("team_place") _
                , .Item("team_country") _
                , .Item("recorded_year") _
                , .Item("recorded_month") _
                , .Item("recorded_day") _
                , .Item("record_scale") _
                , .Item("target_grade") _
                , .Item("comment") _
                , .Item("recorded_ymd") _
                , .Item("id") _
        )
    End With

    Set IJsonParser_Parse = pNewObject
End Function
