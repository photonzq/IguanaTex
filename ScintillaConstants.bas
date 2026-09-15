Attribute VB_Name = "ScintillaConstants"
Option Explicit

' ==============================================================================
' ScintillaConstants: Win32 API declarations, Scintilla messages, and styles
' Supports both 32-bit (x86) and 64-bit (x64) Microsoft Office.
' ==============================================================================

#If VBA7 Then
    Public Declare PtrSafe Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As LongPtr
    Public Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
    Public Declare PtrSafe Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As LongPtr
    Public Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr

    Public Declare PtrSafe Function CreateWindowExA Lib "user32" ( _
        ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, _
        ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, _
        ByVal nWidth As Long, ByVal nHeight As Long, _
        ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, _
        ByVal hInstance As LongPtr, ByVal lpParam As LongPtr) As LongPtr

    Public Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    Public Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
    Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
    Public Declare PtrSafe Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As LongPtr) As LongPtr
    Public Declare PtrSafe Function GetFocusAPI Lib "user32" Alias "GetFocus" () As LongPtr
    Public Declare PtrSafe Function IsWindow Lib "user32" (ByVal hWnd As LongPtr) As Long

    Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" ( _
        ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

    Public Declare PtrSafe Function SendMessageStr Lib "user32" Alias "SendMessageA" ( _
        ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As String) As LongPtr

    Public Declare PtrSafe Function SendMessageBytes Lib "user32" Alias "SendMessageA" ( _
        ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByRef lParam As Any) As LongPtr

    Public Declare PtrSafe Function CreateLexer Lib "Lexilla.dll" (ByVal name As String) As LongPtr

    Public Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Public Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As RECT) As Long
    Public Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As RECT) As Long

    Public Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    Public Declare PtrSafe Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal uIDEvent As LongPtr) As Long

    Public Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
        ByVal CodePage As Long, ByVal dwFlags As Long, _
        ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, _
        ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, _
        ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long

    Public Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
        ByVal CodePage As Long, ByVal dwFlags As Long, _
        ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, _
        ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long

    Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
    Public Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
#Else
    Public Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
    Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
    Public Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
    Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

    Public Declare Function CreateWindowExA Lib "user32" ( _
        ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, _
        ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, _
        ByVal nWidth As Long, ByVal nHeight As Long, _
        ByVal hWndParent As Long, ByVal hMenu As Long, _
        ByVal hInstance As Long, ByVal lpParam As Long) As Long

    Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
    Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
    Public Declare Function GetFocusAPI Lib "user32" Alias "GetFocus" () As Long
    Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long

    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
        ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Public Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" ( _
        ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

    Public Declare Function SendMessageBytes Lib "user32" Alias "SendMessageA" ( _
        ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

    Public Declare Function CreateLexer Lib "Lexilla.dll" (ByVal name As String) As Long

    Public Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
    Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long

    Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Public Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal uIDEvent As Long) As Long

    Public Declare Function WideCharToMultiByte Lib "kernel32" ( _
        ByVal CodePage As Long, ByVal dwFlags As Long, _
        ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, _
        ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, _
        ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

    Public Declare Function MultiByteToWideChar Lib "kernel32" ( _
        ByVal CodePage As Long, ByVal dwFlags As Long, _
        ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, _
        ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

    Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
    Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
#End If

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' Window Styles
Public Const WS_CHILD As Long = &H40000000
Public Const WS_VISIBLE As Long = &H10000000
Public Const WS_TABSTOP As Long = &H10000
Public Const WS_CLIPCHILDREN As Long = &H2000000
Public Const WS_CLIPSIBLINGS As Long = &H4000000
Public Const WS_VSCROLL As Long = &H200000
Public Const WS_HSCROLL As Long = &H100000
Public Const WS_BORDER As Long = &H800000
Public Const WS_EX_CLIENTEDGE As Long = &H200

' Scintilla Messages
Public Const SCI_SETTEXT As Long = 2181
Public Const SCI_GETTEXT As Long = 2182
Public Const SCI_GETLENGTH As Long = 2006
Public Const SCI_SETILEXER As Long = 4033
Public Const SCI_STYLESETFONT As Long = 2056
Public Const SCI_STYLESETSIZE As Long = 2055
Public Const SCI_STYLESETFORE As Long = 2051
Public Const SCI_STYLESETBACK As Long = 2052
Public Const SCI_STYLESETBOLD As Long = 2053
Public Const SCI_STYLESETITALIC As Long = 2054
Public Const SCI_STYLESETUNDERLINE As Long = 2059
Public Const SCI_STYLECLEARALL As Long = 2050
Public Const SCI_SETMARGINWIDTHN As Long = 2242
Public Const SCI_SETMARGINTYPEN As Long = 2240
Public Const SCI_SETMARGINMASKN As Long = 2244
Public Const SCI_SETWRAPMODE As Long = 2268
Public Const SCI_GETWRAPMODE As Long = 2269
Public Const SCI_BRACEMATCH As Long = 2353
Public Const SCI_BRACEHIGHLIGHT As Long = 2351
Public Const SCI_BRACEBADLIGHT As Long = 2352
Public Const SCI_GETCURRENTPOS As Long = 2008
Public Const SCI_SETSEL As Long = 2160
Public Const SCI_GETSELECTIONSTART As Long = 2143
Public Const SCI_GETSELECTIONEND As Long = 2145
Public Const SCI_SETCURRENTPOS As Long = 2141
Public Const SCI_GETCHARAT As Long = 2007
Public Const SCI_SETCODEPAGE As Long = 2037
Public Const SCI_SETTABWIDTH As Long = 2036
Public Const SCI_SETINDENTATIONGUIDES As Long = 2132
Public Const SCI_SETCARETLINEVISIBLE As Long = 2096
Public Const SCI_SETCARETLINEBACK As Long = 2098
Public Const SCI_SETFOCUS As Long = 2380

' Style indices
Public Const STYLE_DEFAULT As Long = 32
Public Const STYLE_LINENUMBER As Long = 33
Public Const STYLE_BRACELIGHT As Long = 34
Public Const STYLE_BRACEBAD As Long = 35

' Margin types
Public Const SC_MARGIN_NUMBER As Long = 1
Public Const SC_CP_UTF8 As Long = 65001
Public Const SC_WRAP_NONE As Long = 0
Public Const SC_WRAP_WORD As Long = 1

' LaTeX Lexer Token Styles (SCE_L_*)
Public Const SCE_L_DEFAULT As Long = 0
Public Const SCE_L_COMMAND As Long = 1
Public Const SCE_L_TAG As Long = 2
Public Const SCE_L_MATH As Long = 3
Public Const SCE_L_COMMENT As Long = 4
Public Const SCE_L_TAG2 As Long = 5
Public Const SCE_L_MATH2 As Long = 6
Public Const SCE_L_COMMENT2 As Long = 7
Public Const SCE_L_VERBATIM As Long = 8
Public Const SCE_L_SHORTCMD As Long = 9
Public Const SCE_L_SPECIAL As Long = 10
Public Const SCE_L_CMDOPT As Long = 11
Public Const SCE_L_ERROR As Long = 12

' Module-level DLL handles and active editor
#If VBA7 Then
    Private m_hLexilla As LongPtr
    Private m_hScintilla As LongPtr
    Public g_hActiveScintilla As LongPtr
    Private m_TimerID As LongPtr
#Else
    Private m_hLexilla As Long
    Private m_hScintilla As Long
    Public g_hActiveScintilla As Long
    Private m_TimerID As Long
#End If

Private m_IsLoaded As Boolean

' ------------------------------------------------------------------------------
' DLL Management
' ------------------------------------------------------------------------------

Public Function EnsureScintillaLoaded() As Boolean
    If m_IsLoaded Then
        EnsureScintillaLoaded = True
        Exit Function
    End If

    Dim archFolder As String
    #If Win64 Then
        archFolder = "x64"
    #Else
        archFolder = "x86"
    #End If

    Dim searchPaths(1 To 5) As String
    Dim addinPath As String
    addinPath = GetAddInDirectory()

    searchPaths(1) = addinPath & "\lib\" & archFolder
    searchPaths(2) = addinPath & "\" & archFolder
    searchPaths(3) = addinPath
    searchPaths(4) = Environ$("APPDATA") & "\Microsoft\AddIns\lib\" & archFolder
    searchPaths(5) = "C:\Users\" & Environ$("USERNAME") & "\Desktop\AI Playground\IguanaTex\lib\" & archFolder

    Dim i As Long, candidateDir As String
    Dim lexPath As String, sciPath As String
    Dim found As Boolean

    For i = 1 To 5
        candidateDir = searchPaths(i)
        lexPath = candidateDir & "\Lexilla.dll"
        sciPath = candidateDir & "\Scintilla.dll"
        If FileExists(lexPath) And FileExists(sciPath) Then
            found = True
            Exit For
        End If
    Next i

    If Not found Then
        EnsureScintillaLoaded = False
        Exit Function
    End If

    #If VBA7 Then
        m_hLexilla = LoadLibraryA(lexPath)
        m_hScintilla = LoadLibraryA(sciPath)
    #Else
        m_hLexilla = LoadLibraryA(lexPath)
        m_hScintilla = LoadLibraryA(sciPath)
    #End If

    If m_hLexilla <> 0 And m_hScintilla <> 0 Then
        m_IsLoaded = True
        EnsureScintillaLoaded = True
    Else
        EnsureScintillaLoaded = False
    End If
End Function

Private Function GetAddInDirectory() As String
    On Error Resume Next
    Dim ai As AddIn
    For Each ai In Application.AddIns
        If InStr(1, ai.Name, "IguanaTex", vbTextCompare) > 0 Then
            GetAddInDirectory = ai.path
            Exit Function
        End If
    Next ai
    GetAddInDirectory = ActivePresentation.path
    If GetAddInDirectory = vbNullString Then
        GetAddInDirectory = CurDir$()
    End If
End Function

Private Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir$(filePath, vbNormal Or vbReadOnly Or vbHidden) <> vbNullString)
End Function

' ------------------------------------------------------------------------------
' Form hWnd Retrieval
' ------------------------------------------------------------------------------

#If VBA7 Then
Public Function ScintillaGetFormHwnd(ByVal frm As Object) As LongPtr
#Else
Public Function ScintillaGetFormHwnd(ByVal frm As Object) As Long
#End If
    Dim hWnd As LongPtr
    hWnd = FindWindowA("ThunderDFrame", frm.Caption)
    If hWnd = 0 Then
        hWnd = FindWindowA(vbNullString, frm.Caption)
    End If
    ScintillaGetFormHwnd = hWnd
End Function

' ------------------------------------------------------------------------------
' Setup Scintilla with LaTeX Lexer & Styles
' ------------------------------------------------------------------------------

#If VBA7 Then
Public Sub SetupScintillaEditor(ByVal hSci As LongPtr, Optional ByVal fontSize As Long = 11)
    Dim pLexer As LongPtr
#Else
Public Sub SetupScintillaEditor(ByVal hSci As Long, Optional ByVal fontSize As Long = 11)
    Dim pLexer As Long
#End If
    If hSci = 0 Then Exit Sub

    ' 1. UTF-8 Encoding
    SendMessage hSci, SCI_SETCODEPAGE, SC_CP_UTF8, 0

    ' 2. Bind Lexilla LaTeX Lexer
    On Error Resume Next
    pLexer = CreateLexer("latex")
    On Error GoTo 0
    If pLexer <> 0 Then
        SendMessage hSci, SCI_SETILEXER, 0, pLexer
    End If

    ' 3. Base Font & Colors (Consolas)
    SendMessageStr hSci, SCI_STYLESETFONT, STYLE_DEFAULT, "Consolas" & vbNullChar
    SendMessage hSci, SCI_STYLESETSIZE, STYLE_DEFAULT, fontSize
    SendMessage hSci, SCI_STYLESETFORE, STYLE_DEFAULT, RGB(30, 30, 30)
    SendMessage hSci, SCI_STYLESETBACK, STYLE_DEFAULT, RGB(255, 255, 255)
    SendMessage hSci, SCI_STYLECLEARALL, 0, 0

    ' 4. Line Number Margin (Margin 0)
    SendMessage hSci, SCI_SETMARGINTYPEN, 0, SC_MARGIN_NUMBER
    SendMessage hSci, SCI_SETMARGINWIDTHN, 0, 38
    SendMessageStr hSci, SCI_STYLESETFONT, STYLE_LINENUMBER, "Consolas" & vbNullChar
    SendMessage hSci, SCI_STYLESETSIZE, STYLE_LINENUMBER, 9
    SendMessage hSci, SCI_STYLESETFORE, STYLE_LINENUMBER, RGB(140, 140, 140)
    SendMessage hSci, SCI_STYLESETBACK, STYLE_LINENUMBER, RGB(245, 245, 247)

    ' 5. LaTeX Syntax Highlighting Styles
    ' Commands (\begin, \end, \frac, \alpha, etc.) -> Blue, Bold
    SendMessage hSci, SCI_STYLESETFORE, SCE_L_COMMAND, RGB(0, 70, 210)
    SendMessage hSci, SCI_STYLESETBOLD, SCE_L_COMMAND, 1

    ' Tags / Environments ({equation}, {align}) -> Dark Cyan, Bold
    SendMessage hSci, SCI_STYLESETFORE, SCE_L_TAG, RGB(0, 130, 130)
    SendMessage hSci, SCI_STYLESETBOLD, SCE_L_TAG, 1

    ' Inline Math ($...$) -> Dark Green
    SendMessage hSci, SCI_STYLESETFORE, SCE_L_MATH, RGB(20, 128, 20)

    ' Display Math ($$...$$, \[...\]) -> Forest Green, Bold
    SendMessage hSci, SCI_STYLESETFORE, SCE_L_MATH2, RGB(0, 110, 0)
    SendMessage hSci, SCI_STYLESETBOLD, SCE_L_MATH2, 1

    ' Comments (% ...) -> Gray / Slate, Italic
    SendMessage hSci, SCI_STYLESETFORE, SCE_L_COMMENT, RGB(120, 130, 140)
    SendMessage hSci, SCI_STYLESETITALIC, SCE_L_COMMENT, 1
    SendMessage hSci, SCI_STYLESETFORE, SCE_L_COMMENT2, RGB(120, 130, 140)
    SendMessage hSci, SCI_STYLESETITALIC, SCE_L_COMMENT2, 1

    ' Short commands (\\, \{, \}) -> Dark Blue
    SendMessage hSci, SCI_STYLESETFORE, SCE_L_SHORTCMD, RGB(0, 50, 160)

    ' Special symbols (&, ^, _, ~) -> Crimson / Dark Red
    SendMessage hSci, SCI_STYLESETFORE, SCE_L_SPECIAL, RGB(180, 30, 30)

    ' Command options ([12pt], [h!]) -> Dark Golden / Orange
    SendMessage hSci, SCI_STYLESETFORE, SCE_L_CMDOPT, RGB(160, 80, 0)

    ' 6. Bracket Matching Styles
    ' Matched braces -> Soft blue background, bold
    SendMessage hSci, SCI_STYLESETBOLD, STYLE_BRACELIGHT, 1
    SendMessage hSci, SCI_STYLESETBACK, STYLE_BRACELIGHT, RGB(190, 225, 255)
    SendMessage hSci, SCI_STYLESETFORE, STYLE_BRACELIGHT, RGB(0, 40, 120)

    ' Unmatched brace -> Soft red background, bold red
    SendMessage hSci, SCI_STYLESETBOLD, STYLE_BRACEBAD, 1
    SendMessage hSci, SCI_STYLESETBACK, STYLE_BRACEBAD, RGB(255, 200, 200)
    SendMessage hSci, SCI_STYLESETFORE, STYLE_BRACEBAD, RGB(200, 0, 0)

    ' 7. Caret & Editor Usability
    SendMessage hSci, SCI_SETCARETLINEVISIBLE, 1, 0
    SendMessage hSci, SCI_SETCARETLINEBACK, RGB(246, 248, 254), 0
    SendMessage hSci, SCI_SETTABWIDTH, 2, 0
    SendMessage hSci, SCI_SETINDENTATIONGUIDES, 1, 0
    SendMessage hSci, SCI_SETWRAPMODE, SC_WRAP_WORD, 0
End Sub

' ------------------------------------------------------------------------------
' Text Encoding (UTF-8)
' ------------------------------------------------------------------------------

#If VBA7 Then
Public Function ScintillaGetText(ByVal hSci As LongPtr) As String
#Else
Public Function ScintillaGetText(ByVal hSci As Long) As String
#End If
    If hSci = 0 Then Exit Function
    Dim textLen As Long
    textLen = SendMessage(hSci, SCI_GETLENGTH, 0, 0)
    If textLen <= 0 Then
        ScintillaGetText = ""
        Exit Function
    End If

    Dim utf8Bytes() As Byte
    ReDim utf8Bytes(0 To textLen)
    SendMessageBytes hSci, SCI_GETTEXT, textLen + 1, utf8Bytes(0)

    ' Convert UTF-8 to VBA BSTR (UTF-16)
    Dim wideLen As Long
    #If VBA7 Then
        wideLen = MultiByteToWideChar(SC_CP_UTF8, 0, VarPtr(utf8Bytes(0)), textLen, 0, 0)
        If wideLen > 0 Then
            Dim result As String
            result = Space$(wideLen)
            MultiByteToWideChar SC_CP_UTF8, 0, VarPtr(utf8Bytes(0)), textLen, StrPtr(result), wideLen
            ScintillaGetText = result
        End If
    #Else
        wideLen = MultiByteToWideChar(SC_CP_UTF8, 0, VarPtr(utf8Bytes(0)), textLen, 0, 0)
        If wideLen > 0 Then
            Dim result As String
            result = Space$(wideLen)
            MultiByteToWideChar SC_CP_UTF8, 0, VarPtr(utf8Bytes(0)), textLen, StrPtr(result), wideLen
            ScintillaGetText = result
        End If
    #End If
End Function

#If VBA7 Then
Public Sub ScintillaSetText(ByVal hSci As LongPtr, ByVal text As String)
#Else
Public Sub ScintillaSetText(ByVal hSci As Long, ByVal text As String)
#End If
    If hSci = 0 Then Exit Sub
    If Len(text) = 0 Then
        SendMessageStr hSci, SCI_SETTEXT, 0, "" & vbNullChar
        Exit Sub
    End If

    ' Convert VBA BSTR (UTF-16) to UTF-8
    Dim utf8Len As Long
    #If VBA7 Then
        utf8Len = WideCharToMultiByte(SC_CP_UTF8, 0, StrPtr(text), Len(text), 0, 0, 0, 0)
        If utf8Len > 0 Then
            Dim utf8Bytes() As Byte
            ReDim utf8Bytes(0 To utf8Len)
            WideCharToMultiByte SC_CP_UTF8, 0, StrPtr(text), Len(text), VarPtr(utf8Bytes(0)), utf8Len, 0, 0
            utf8Bytes(utf8Len) = 0
            SendMessageBytes hSci, SCI_SETTEXT, 0, utf8Bytes(0)
        End If
    #Else
        utf8Len = WideCharToMultiByte(SC_CP_UTF8, 0, StrPtr(text), Len(text), 0, 0, 0, 0)
        If utf8Len > 0 Then
            Dim utf8Bytes() As Byte
            ReDim utf8Bytes(0 To utf8Len)
            WideCharToMultiByte SC_CP_UTF8, 0, StrPtr(text), Len(text), VarPtr(utf8Bytes(0)), utf8Len, 0, 0
            utf8Bytes(utf8Len) = 0
            SendMessageBytes hSci, SCI_SETTEXT, 0, utf8Bytes(0)
        End If
    #End If
End Sub

' ------------------------------------------------------------------------------
' Bracket Matching & Timer
' ------------------------------------------------------------------------------

#If VBA7 Then
Public Sub CheckBraceMatch(ByVal hSci As LongPtr)
#Else
Public Sub CheckBraceMatch(ByVal hSci As Long)
#End If
    If hSci = 0 Then Exit Sub
    If IsWindow(hSci) = 0 Then Exit Sub

    Dim pos As Long, matchPos As Long, ch As Long
    pos = SendMessage(hSci, SCI_GETCURRENTPOS, 0, 0)

    If pos > 0 Then
        ch = SendMessage(hSci, SCI_GETCHARAT, pos - 1, 0)
        If IsBraceChar(ch) Then
            matchPos = SendMessage(hSci, SCI_BRACEMATCH, pos - 1, 0)
            If matchPos <> -1 Then
                SendMessage hSci, SCI_BRACEHIGHLIGHT, pos - 1, matchPos
                Exit Sub
            Else
                SendMessage hSci, SCI_BRACEBADLIGHT, pos - 1, 0
                Exit Sub
            End If
        End If
    End If

    ch = SendMessage(hSci, SCI_GETCHARAT, pos, 0)
    If IsBraceChar(ch) Then
        matchPos = SendMessage(hSci, SCI_BRACEMATCH, pos, 0)
        If matchPos <> -1 Then
            SendMessage hSci, SCI_BRACEHIGHLIGHT, pos, matchPos
            Exit Sub
        Else
            SendMessage hSci, SCI_BRACEBADLIGHT, pos, 0
            Exit Sub
        End If
    End If

    SendMessage hSci, SCI_BRACEHIGHLIGHT, -1, -1
End Sub

Private Function IsBraceChar(ByVal ch As Long) As Boolean
    Select Case ch
        Case 40, 41   ' ( )
            IsBraceChar = True
        Case 91, 93   ' [ ]
            IsBraceChar = True
        Case 123, 125 ' { }
            IsBraceChar = True
        Case Else
            IsBraceChar = False
    End Select
End Function

#If VBA7 Then
Public Sub StartBracketMatchTimer(ByVal hSci As LongPtr)
#Else
Public Sub StartBracketMatchTimer(ByVal hSci As Long)
#End If
    g_hActiveScintilla = hSci
    If m_TimerID = 0 And hSci <> 0 Then
        m_TimerID = SetTimer(0, 0, 150, AddressOf ScintillaTimerProc)
    End If
End Sub

Public Sub StopBracketMatchTimer()
    If m_TimerID <> 0 Then
        KillTimer 0, m_TimerID
        m_TimerID = 0
    End If
    g_hActiveScintilla = 0
End Sub

#If VBA7 Then
Public Sub ScintillaTimerProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long)
#Else
Public Sub ScintillaTimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
#End If
    On Error Resume Next
    If g_hActiveScintilla <> 0 Then
        CheckBraceMatch g_hActiveScintilla
    End If
End Sub
