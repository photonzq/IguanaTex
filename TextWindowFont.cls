VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextWindowFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==============================================================================
' TextWindowFont: Font wrapper for Scintilla editor
' ==============================================================================

Option Explicit

#If VBA7 Then
    Private m_hSci As LongPtr
#Else
    Private m_hSci As Long
#End If
Private m_target As Object
Private m_size As Double

#If VBA7 Then
Public Sub InitWin(ByVal hSci As LongPtr, ByVal targetCtrl As Object)
#Else
Public Sub InitWin(ByVal hSci As Long, ByVal targetCtrl As Object)
#End If
    m_hSci = hSci
    Set m_target = targetCtrl
    If m_size = 0 Then
        If Not targetCtrl Is Nothing Then
            On Error Resume Next
            m_size = targetCtrl.Font.Size
            On Error GoTo 0
        End If
        If m_size = 0 Then m_size = 11
    End If
End Sub

Public Property Get Size() As Double
    If m_hSci <> 0 Then
        Size = m_size
    ElseIf Not m_target Is Nothing Then
        On Error Resume Next
        Size = m_target.Font.Size
        On Error GoTo 0
    Else
        Size = m_size
    End If
End Property

Public Property Let Size(ByVal value As Double)
    m_size = value
    If m_hSci <> 0 Then
        SetupScintillaEditor m_hSci, CLng(value)
    End If
    If Not m_target Is Nothing Then
        On Error Resume Next
        m_target.Font.Size = value
        On Error GoTo 0
    End If
End Property
