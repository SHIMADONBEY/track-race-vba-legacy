VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "MacroProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'namespace=vba-files/common/macro 
Option Explicit

Private m_ScreenUpdating As Boolean
Private m_Calculation As XlCalculation
Private m_EnableEvents As Boolean

Private m_CurrentScreenUpdating As Boolean

Private Sub Class_Initialize()
    m_ScreenUpdating = Application.ScreenUpdating
    m_Calculation = Application.Calculation
    m_EnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    m_CurrentScreenUpdating = Application.ScreenUpdating

    Debug.Print Format(Now, "yyyy-MM-ddThh:nn:ss") & " - Macro Started."
End Sub

Private Sub Class_Terminate()
    Application.ScreenUpdating = m_ScreenUpdating
    Application.Calculation = m_Calculation
    Application.EnableEvents = m_EnableEvents

    Debug.Print Format(Now, "yyyy-MM-ddThh:nn:ss") & " - Macro Ended. (" & Err.Number & ") " & Err.Description

    If Not (Err.Number = 0) Then
        Debug.Print Format(Now, "yyyy-MM-ddThh:nn:ss") & " - Error occured at " & Err.Source
    End If
End Sub

Public Sub SwitchScreenUpdating()
    Application.ScreenUpdating = Not(m_CurrentScreenUpdating)
    m_CurrentScreenUpdating = Application.ScreenUpdating
End Sub
