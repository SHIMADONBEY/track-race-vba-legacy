VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "WindValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/common/value/wind
Option Explicit
Private m_Value As Variant

Private m_Initialized As Boolean

Private Sub Class_Initalize()
    m_Initialized = false
End Sub

Private Sub Class_Terminate()
End Sub

Public Sub Initialize(Value As Variant)
    If (m_Initialized) Then Err.Raise(91)

    Dim roundedValue As Long: roundedValue = Int(CDec(Value) * CDec(10) + CDec(0.999))
    m_Value = CDec(roundedValue) / CDec(10)

    m_Initialized = True
End Sub

Public Property Get Value()
    If Not (m_Initialized) Then Err.Raise(91)
    Value = m_Value
End Property

Public Function Compare(ByRef Other As WindValue)
    If (Other Is Nothing) Then 
        Compare = -1
    Else
        Compare = Sgn(Me.Value - Other.Value)
    End If
End Function
