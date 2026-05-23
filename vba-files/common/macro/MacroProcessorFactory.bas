Attribute VB_Name = "MacroProcessorFactory"
'namespace=vba-files/common/macro
Option Explicit

Private m_Instance As MacroProcessor

Public Function GetInstance() As MacroProcessor
    If (m_Instance Is Nothing) Then
        Call CreateInstance()
    End If

    Set GetInstance = m_Instance
End Function

Public Sub DestroyInstance()
    If Not (m_Instance Is Nothing) Then
        Set m_Instance = Nothing
    End If
End Sub

Public Sub CreateInstance()
    If (m_Instance Is Nothing) Then
        Set m_Instance = New MacroProcessor
    End If
End Sub