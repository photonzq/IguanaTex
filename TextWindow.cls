VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Replacement for TextBox

Option Explicit

#If Mac Then
Private Declare PtrSafe Function TWInit _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   () As LongLong

Private Declare PtrSafe Function TWTerm _
 Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
(ByVal Handle As LongLong, ByVal b As LongLong, ByVal c As LongLong, ByVal d As LongLong) As LongLong

Private Declare PtrSafe Function TWSetResizeTarget _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   (ByVal Handle As LongLong, ByVal b As LongLong, ByVal c As LongLong, ByVal d As LongLong) As LongLong

Private Declare PtrSafe Function TWResize _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   (ByVal Handle As LongLong, ByVal b As LongLong, ByVal c As LongLong, ByVal d As LongLong) As LongLong

Private Declare PtrSafe Function TWShow _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   (ByVal Handle As LongLong, ByVal b As LongLong, ByVal c As LongLong, ByVal d As LongLong) As LongLong

Private Declare PtrSafe Function TWHide _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   (ByVal Handle As LongLong, ByVal b As LongLong, ByVal c As LongLong, ByVal d As LongLong) As LongLong

Private Declare PtrSafe Function TWFocus _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   (ByVal Handle As LongLong, ByVal b As LongLong, ByVal c As LongLong, ByVal d As LongLong) As LongLong

Private Declare PtrSafe Function TWGetByteLength _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   (ByVal Handle As LongLong, ByVal b As LongLong, ByVal c As LongLong, ByVal d As LongLong) As LongLong

Private Declare PtrSafe Function TWGetBytes _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   (ByVal Handle As LongLong, ByVal buffer As LongPtr, ByVal length As LongLong, ByVal d As LongLong) As LongLong

Private Declare PtrSafe Function TWSetBytes _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   (ByVal Handle As LongLong, ByVal data As LongPtr, ByVal length As LongLong, ByVal d As LongLong) As LongLong

Private Declare PtrSafe Function TWGetSelStart _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   (ByVal Handle As LongLong, ByVal b As LongLong, ByVal c As LongLong, ByVal d As LongLong) As LongLong

Private Declare PtrSafe Function TWSetSelStart _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   (ByVal Handle As LongLong, ByVal value As LongLong, ByVal c As LongLong, ByVal d As LongLong) As LongLong

Private Declare PtrSafe Function TWGetWordWrap _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   (ByVal Handle As LongLong, ByVal b As LongLong, ByVal c As LongLong, ByVal d As LongLong) As LongLong

Private Declare PtrSafe Function TWSetWordWrap _
    Lib "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib" _
   (ByVal Handle As LongLong, ByVal value As LongLong, ByVal c As LongLong, ByVal d As LongLong) As LongLong

Private Const UNUSED As LongLong = 0

' members
Private mHandle As LongLong
Private mFont As New TextWindowFont

' properties
Public Property Get Font() As TextWindowFont
    Set Font = mFont
End Property

Public Property Get Text() As String
    Dim length As LongLong
    length = TWGetByteLength(mHandle, UNUSED, UNUSED, UNUSED)
    
    If length <= 0 Then
        Text = ""
        Exit Property
    End If

    Dim bytes() As Byte
    Dim usedLength As LongLong
    ReDim bytes(CLng(length - 1))
    usedLength = TWGetBytes(mHandle, VarPtr(bytes(0)), length, UNUSED)

    If usedLength <= 0 Then
        Text = ""
        Exit Property
    End If

    ReDim Preserve bytes(CLng(usedLength - 1))
    Text = bytes
End Property

Public Property Let Text(value As String)
    Dim bytes() As Byte
    Dim length As LongLong
    bytes = value
    length = ArrayLength(bytes)
    If length > 0 Then
         TWSetBytes mHandle, VarPtr(bytes(0)), length, UNUSED
    Else
         TWSetBytes mHandle, 0, 0, UNUSED
    End If
End Property

Public Property Get SelStart() As Integer
    SelStart = CInt(TWGetSelStart(mHandle, UNUSED, UNUSED, UNUSED))
End Property

Public Property Let SelStart(value As Integer)
    TWSetSelStart mHandle, CLngLng(value), UNUSED, UNUSED
End Property

Public Property Get WordWrap() As Boolean
    WordWrap = TWGetWordWrap(mHandle, UNUSED, UNUSED, UNUSED)
End Property

Public Property Let WordWrap(value As Boolean)
    TWSetWordWrap mHandle, value, UNUSED, UNUSED
End Property

' constructor/destructor
Private Sub Class_Initialize()
    mHandle = TWInit
    mFont.Init mHandle
End Sub

Private Sub Class_Terminate()
    TWTerm mHandle, UNUSED, UNUSED, UNUSED
End Sub

' methods
Public Sub Show()
    TWShow mHandle, UNUSED, UNUSED, UNUSED
End Sub

Public Sub Hide()
    TWHide mHandle, UNUSED, UNUSED, UNUSED
End Sub

Public Sub SetResizeTarget(target As Object, parentForm As Object)
    Dim backup As Control
    Set backup = parentForm.ActiveControl
    target.SetFocus
    TWSetResizeTarget mHandle, UNUSED, UNUSED, UNUSED
    If Not backup Is Nothing Then
        backup.SetFocus
    End If
End Sub

Public Sub ResizeAsTarget()
    TWResize mHandle, UNUSED, UNUSED, UNUSED
End Sub

Public Sub SetFocus()
    TWFocus mHandle, UNUSED, UNUSED, UNUSED
End Sub

#Else
' ==============================================================================
' Windows Scintilla Implementation
' ==============================================================================

#If VBA7 Then
    Private m_hSci As LongPtr
    Private m_hWndParent As LongPtr
#Else
    Private m_hSci As Long
    Private m_hWndParent As Long
#End If

Private m_target As Object
Private m_parentForm As Object
Private m_fallback As Boolean
Private m_Font As New TextWindowFont
Private m_initialText As String
Private m_initialWordWrap As Boolean
Private m_wordWrapSet As Boolean

' properties
Public Property Get Font() As TextWindowFont
    Set Font = m_Font
End Property

Public Property Get Text() As String
    If m_hSci <> 0 Then
        Text = ScintillaGetText(m_hSci)
    ElseIf Not m_target Is Nothing Then
        Text = m_target.Text
    Else
        Text = m_initialText
    End If
End Property

Public Property Let Text(value As String)
    m_initialText = value
    If m_hSci <> 0 Then
        ScintillaSetText m_hSci, value
    End If
    If Not m_target Is Nothing Then
        On Error Resume Next
        m_target.Text = value
        On Error GoTo 0
    End If
End Property

Public Property Get SelStart() As Long
    If m_hSci <> 0 Then
        SelStart = CLng(SendMessage(m_hSci, SCI_GETCURRENTPOS, 0, 0))
    ElseIf Not m_target Is Nothing Then
        SelStart = m_target.SelStart
    Else
        SelStart = 0
    End If
End Property

Public Property Let SelStart(value As Long)
    If m_hSci <> 0 Then
        SendMessage m_hSci, SCI_SETCURRENTPOS, CLng(value), 0
        SendMessage m_hSci, SCI_SETSEL, CLng(value), CLng(value)
    End If
    If Not m_target Is Nothing Then
        On Error Resume Next
        m_target.SelStart = value
        On Error GoTo 0
    End If
End Property

Public Property Get WordWrap() As Boolean
    If m_hSci <> 0 Then
        WordWrap = (SendMessage(m_hSci, SCI_GETWRAPMODE, 0, 0) = SC_WRAP_WORD)
    ElseIf Not m_target Is Nothing Then
        WordWrap = m_target.WordWrap
    Else
        WordWrap = m_initialWordWrap
    End If
End Property

Public Property Let WordWrap(value As Boolean)
    m_initialWordWrap = value
    m_wordWrapSet = True
    If m_hSci <> 0 Then
        If value Then
            SendMessage m_hSci, SCI_SETWRAPMODE, SC_WRAP_WORD, 0
        Else
            SendMessage m_hSci, SCI_SETWRAPMODE, SC_WRAP_NONE, 0
        End If
    End If
    If Not m_target Is Nothing Then
        On Error Resume Next
        m_target.WordWrap = value
        On Error GoTo 0
    End If
End Property

' constructor/destructor
Private Sub Class_Initialize()
    m_fallback = False
    m_initialWordWrap = True
    m_wordWrapSet = False
End Sub

Private Sub Class_Terminate()
    If m_hSci <> 0 Then
        StopBracketMatchTimer
        DestroyWindow m_hSci
        m_hSci = 0
    End If
End Sub

' methods
Public Sub Show()
    If m_hSci <> 0 Then
        ShowWindow m_hSci, 5 ' SW_SHOW
    ElseIf Not m_target Is Nothing Then
        m_target.Visible = True
    End If
End Sub

Public Sub Hide()
    If m_hSci <> 0 Then
        ShowWindow m_hSci, 0 ' SW_HIDE
    ElseIf Not m_target Is Nothing Then
        m_target.Visible = False
    End If
End Sub

Public Sub SetFocus()
    If m_hSci <> 0 Then
        SetFocusAPI m_hSci
        SendMessage m_hSci, SCI_SETFOCUS, 1, 0
    ElseIf Not m_target Is Nothing Then
        m_target.SetFocus
    End If
End Sub

Public Sub SetResizeTarget(target As Object, parentForm As Object)
    Set m_target = target
    Set m_parentForm = parentForm

    If m_fallback Then Exit Sub

    ' 1. Check/Load Scintilla & Lexilla DLLs
    If Not EnsureScintillaLoaded() Then
        m_fallback = True
        If Not m_target Is Nothing Then m_target.Visible = True
        Exit Sub
    End If

    ' 2. Get UserForm Window Handle
    m_hWndParent = ScintillaGetFormHwnd(parentForm)
    If m_hWndParent = 0 Then
        m_fallback = True
        If Not m_target Is Nothing Then m_target.Visible = True
        Exit Sub
    End If

    ' 3. Create Scintilla Window if not already created
    If m_hSci = 0 Then
        Dim x As Long, y As Long, w As Long, h As Long
        CalculateTargetRect target, parentForm, x, y, w, h

        #If VBA7 Then
            m_hSci = CreateWindowExA(WS_EX_CLIENTEDGE, "Scintilla", vbNullString, _
                WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS, _
                x, y, w, h, m_hWndParent, 0, 0, 0)
        #Else
            m_hSci = CreateWindowExA(WS_EX_CLIENTEDGE, "Scintilla", vbNullString, _
                WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS, _
                x, y, w, h, m_hWndParent, 0, 0, 0)
        #End If

        If m_hSci = 0 Then
            m_fallback = True
            If Not m_target Is Nothing Then m_target.Visible = True
            Exit Sub
        End If

        ' Initialize Scintilla with LaTeX Lexer and styles
        SetupScintillaEditor m_hSci, 11

        ' Apply initial text
        If Len(m_initialText) > 0 Then
            ScintillaSetText m_hSci, m_initialText
        ElseIf Len(m_target.Text) > 0 Then
            ScintillaSetText m_hSci, m_target.Text
        End If

        ' Apply initial word wrap
        If m_wordWrapSet Then
            WordWrap = m_initialWordWrap
        End If

        ' Initialize font wrapper
        m_Font.InitWin m_hSci, m_target

        ' Start bracket matching timer
        StartBracketMatchTimer m_hSci

        ' Hide the placeholder TextBox
        On Error Resume Next
        m_target.Visible = False
        On Error GoTo 0
    Else
        ' Already created, just resize
        ResizeAsTarget
    End If
End Sub

Public Sub ResizeAsTarget()
    If m_hSci <> 0 And Not m_target Is Nothing And Not m_parentForm Is Nothing Then
        Dim x As Long, y As Long, w As Long, h As Long
        CalculateTargetRect m_target, m_parentForm, x, y, w, h
        MoveWindow m_hSci, x, y, w, h, 1
    End If
End Sub

Private Sub CalculateTargetRect(target As Object, parentForm As Object, _
                               ByRef outX As Long, ByRef outY As Long, _
                               ByRef outW As Long, ByRef outH As Long)
    Dim clientRc As RECT
    GetClientRect m_hWndParent, clientRc

    Dim clientW As Double, clientH As Double
    clientW = clientRc.Right - clientRc.Left
    clientH = clientRc.Bottom - clientRc.Top

    Dim scaleX As Double, scaleY As Double
    If parentForm.InsideWidth > 0 Then
        scaleX = clientW / parentForm.InsideWidth
    Else
        scaleX = 1.333333
    End If

    If parentForm.InsideHeight > 0 Then
        scaleY = clientH / parentForm.InsideHeight
    Else
        scaleY = 1.333333
    End If

    Dim ptLeft As Double, ptTop As Double
    ptLeft = target.Left
    ptTop = target.Top

    ' If inside a Page of MultiPage, add container offsets
    If TypeName(target.Parent) = "Page" Then
        On Error Resume Next
        ptLeft = ptLeft + target.Parent.Parent.Left + 2
        ' Offset for MultiPage tab strip (~20-22 pt)
        ptTop = ptTop + target.Parent.Parent.Top + 22
        On Error GoTo 0
    End If

    outX = CLng(ptLeft * scaleX)
    outY = CLng(ptTop * scaleY)
    outW = CLng(target.Width * scaleX)
    outH = CLng(target.Height * scaleY)
End Sub

#End If
