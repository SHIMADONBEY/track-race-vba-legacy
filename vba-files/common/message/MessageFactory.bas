Attribute VB_Name = "MessageFactory"
'namespace=vba-files/common/message
Option Explicit
Option Private Module

Private Enum MessageTableColumnEnum
    IDX_CODE = 1
    IDX_MSG_TYPE
    IDX_SERIAL
    IDX_TITLE
    IDX_PROMPT
End Enum

Private m_MessageDictionary As Object

Public Sub Load()
    Dim vMessages As Object: Set vMessages = CreateObject("Scripting.Dictionary")
    Dim vRowRange As Range

    For Each vRowRange In Range("SystemMessageList").Rows()
        Dim vMessage As MessageModel: Set vMessage = ReadFromRow(vRowRange)
        If (vMessage Is Nothing) Then 
            ' DO NOTHING        
        Else
            Call vMessages.Add(vMessage.Code, vMessage)
        End If
    Next vRowRange 

    Set m_MessageDictionary = vMessages
End Sub

Public Function Generate(Code As String, Optional Reload As Boolean = False) As MessageModel
    If (Reload Or m_MessageDictionary Is Nothing) Then
        Call Load()
    End If

    Set Generate = m_MessageDictionary.Item(Code)
End Function

Private Function ReadFromRow(RowRange As Range) As MessageModel
    Dim pCode As String: pCode = RowRange.Cells(1, MessageTableColumnEnum.IDX_CODE).Value
    If (pCode = "") Then
        Set ReadFromRow = Nothing
        Exit Function
    End If

    With New MessageModel
        Call .Initialize( _
                pCode _
                , RowRange.Cells(1, MessageTableColumnEnum.IDX_MSG_TYPE).Value _
                , RowRange.Cells(1, MessageTableColumnEnum.IDX_TITLE).Value _
                , RowRange.Cells(1, MessageTableColumnEnum.IDX_PROMPT).Value _
        )

        Set ReadFromRow = .Self()
    End With
End Function
