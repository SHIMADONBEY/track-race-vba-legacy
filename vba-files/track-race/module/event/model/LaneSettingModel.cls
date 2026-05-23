VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LaneSettingModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/track-race/module/event/model
Option Explicit
Private m_LaneCount As Long
Private m_TopLanes As Variant
Private m_MiddleLanes As Variant
Private m_BottomLanes As Variant
Private m_ExtraLanes As Variant

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
End Sub

Private Sub Class_Terminate()
End Sub

Public Sub Initialize(LaneCount As Long, TopLanes As String, MiddleLanes As String, BottomLanes As String, ExtraLanes As String)
    If m_Initialized Then Err.Raise (91)
    m_LaneCount = LaneCount

    Dim splitTop() As String: splitTop = Split(TopLanes, ",")
    Dim splitMiddle() As String: splitMiddle = Split(MiddleLanes, ",")
    Dim splitBottom() As String: splitBottom = Split(BottomLanes, ",")
    Dim splitExtra() As String: splitExtra = Split(ExtraLanes, ",")

    Dim i As Long
    If (TopLanes = "") Then
        m_TopLanes = Array()
    Else
        ReDim m_TopLanes(LBound(splitTop) To UBound(splitTop))
        For i = LBound(splitTop) To UBound(splitTop)
            m_TopLanes(i) = CLng(splitTop(i))
        Next i
    End If

    If (MiddleLanes = "") Then
        m_MiddleLanes = Array()
    Else
        ReDim m_MiddleLanes(LBound(splitMiddle) To UBound(splitMiddle))
        For i = LBound(splitMiddle) To UBound(splitMiddle)
            m_MiddleLanes(i) = CLng(splitMiddle(i))
        Next i
    End If

    If (BottomLanes = "") Then
        m_BottomLanes = Array()
    Else
        ReDim m_BottomLanes(LBound(splitBottom) To UBound(splitBottom))
        For i = LBound(splitBottom) To UBound(splitBottom)
            m_BottomLanes(i) = CLng(splitBottom(i))
        Next i
    End If

    If (ExtraLanes = "") Then
        m_ExtraLanes = Array()
    Else
        ReDim m_ExtraLanes(LBound(splitExtra) To UBound(splitExtra))
        For i = LBound(splitExtra) To UBound(splitExtra)
            m_ExtraLanes(i) = CLng(splitExtra(i))
        Next i        
    End If

    m_Initialized = True
End Sub

Public Property Get LaneCount() As Long
    If Not m_Initialized Then Err.Raise (91)
    LaneCount = m_LaneCount
End Property

Public Property Get TopLanes() As Variant
    If Not m_Initialized Then Err.Raise (91)
    TopLanes = m_TopLanes
End Property

Public Property Get MiddleLanes() As Variant
    If Not m_Initialized Then Err.Raise (91)
    MiddleLanes = m_MiddleLanes
End Property

Public Property Get BottomLanes() As Variant
    If Not m_Initialized Then Err.Raise (91)
    BottomLanes = m_BottomLanes
End Property

Public Property Get ExtraLanes() As Variant
    If Not m_Initialized Then Err.Raise (91)
    ExtraLanes = m_ExtraLanes
End Property

Public Function Self() As LaneSettingModel
    If Not m_Initialized Then Err.Raise(91)
    Set Self = Me
End Function
