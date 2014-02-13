Attribute VB_Name = "rtf"
Option Explicit

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCaret Lib "user32" (ByVal hwnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function SetCaretBlinkTime Lib "user32" (ByVal wMSeconds As Long) As Long
Private Declare Function GetCaretBlinkTime Lib "user32" () As Long

'Private Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long) As Long
'Private Declare Function SetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal code As Long, ByVal nPos As Long, ByVal fRedraw As Boolean) As Long
'Private Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal redraw As Boolean) As Long
'Private Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
'Private Const SB_HORZ = 0
'Private Const SB_VERT = 1
'Private Const SIF_RANGE = &H1
'Private Const SIF_PAGE = &H2
'Private Const SIF_POS = &H4
'Private Const SIF_DISABLENOSCROLL = &H8
'Private Const SIF_TRACKPOS = &H10
'Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
'
'Private Type SCROLLINFO
'    cbSize As Long
'    fMask As Long
'    nMin As Long
'    nMax As Long
'    nPage As Long
'    nPos As Long
'    nTrackPos As Long
'End Type

Private Type Rect
    left As Long
    Top As Long
    right As Long
    Bottom As Long
End Type

Private Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

Public Enum tmMsgs
        EM_UNDO = &HC7
        EM_CANUNDO = &HC6
        EM_SETWORDBREAKPROC = &HD0
        EM_SETTABSTOPS = &HCB
        EM_SETSEL = &HB1
        EM_SETRECTNP = &HB4
        EM_SETRECT = &HB3
        EM_SETREADONLY = &HCF
        EM_SETPASSWORDCHAR = &HCC
        EM_SETMODIFY = &HB9
        EM_SCROLLCARET = &HB7
        EM_SETHANDLE = &HBC
        EM_SCROLL = &HB5
        EM_REPLACESEL = &HC2
        EM_LINESCROLL = &HB6
        EM_LINELENGTH = &HC1
        EM_LINEINDEX = &HBB
        EM_LINEFROMCHAR = &HC9
        EM_LIMITTEXT = &HC5
        EM_GETWORDBREAKPROC = &HD1
        EM_GETTHUMB = &HBE
        EM_GETRECT = &HB2
        EM_GETSEL = &HB0
        EM_GETPASSWORDCHAR = &HD2
        EM_GETMODIFY = &HB8
        EM_GETLINECOUNT = &HBA
        EM_GETLINE = &HC4
        EM_GETHANDLE = &HBD
        EM_GETFIRSTVISIBLELINE = &HCE
        EM_FMTLINES = &HC8
        EM_EMPTYUNDOBUFFER = &HCD
        EM_SETMARGINS = &HD3
End Enum

Private Const LF_FACESIZE = 32

Private Type CHARFORMAT2
    cbSize As Integer                           '2
    wPad1 As Integer                            '4
    dwMask As Long                              '8
    dwEffects As Long                           '12
    yHeight As Long                             '16
    yOffset As Long                             '20
    crTextColor As Long                         '24
    bCharSet As Byte                            '25
    bPitchAndFamily As Byte                     '26
    szFaceName(0 To LF_FACESIZE - 1) As Byte    '58
    wPad2 As Integer                            '60

    ' Additional stuff supported by RICHEDIT20
    wWeight As Integer         'Font weight (LOGFONT value)
    sSpacing As Integer        'Amount to space between letters
    crBackColor As Long        'Background color
    lLCID As Long              'Locale ID
    dwReserved As Long         'Reserved. Must be 0
    sStyle As Integer          'Style handle
    wKerning As Integer        'Twip size above which to kern char pair
    bUnderlineType As Byte     'Underline type
    bAnimation As Byte         'Animated text like marching ants
    bRevAuthor As Byte         'Revision author index
    bReserved1 As Byte
End Type

Private Const EM_SETCHARFORMAT = (&H400 + 68)
Private Const CFM_BACKCOLOR = &H4000000
Private Const WM_VScroll = &H115
Private Const WM_CHAR = &H102
Private Const EC_LEFTMARGIN = &H1
Private Const EC_RIGHTMARGIN = &H2

Private myTopLine As Long
Private TrackingScroll As Boolean
Private OverRidingTabs As Boolean
Private OverrideTabNow As Boolean
Private working As Boolean

Sub SetWindowUpdate(rtf As Object, Optional disable As Boolean = True)
    LockWindowUpdate IIf(disable, rtf.hwnd, 0)
End Sub

Sub SetLineColor(index As Long, rtf As Object, color As ColorConstants, Optional bold As Boolean = False)
    On Error Resume Next
    SelectLine index, rtf
    rtf.SelColor = color
    rtf.SelBold = bold
End Sub

Public Sub SelectLine(index As Long, rtf As Object)
    On Error Resume Next
    rtf.selStart = LineStartPos(index, rtf)
    rtf.selLength = LineLength(rtf)
End Sub

Function CurrentLine(rtf As Object) As Long
    CurrentLine = SendMessageLong(rtf.hwnd, EM_LINEFROMCHAR, rtf.selStart, 0&)
End Function

Function LineLength(rtf As Object) As Long
    LineLength = SendMessageLong(rtf.hwnd, EM_LINELENGTH, rtf.selStart, 0&)
End Function

Function LineStartPos(lineIndex As Long, rtf As Object) As Long
    LineStartPos = SendMessageLong(rtf.hwnd, EM_LINEINDEX, lineIndex, 0&)
End Function

'Public Sub GotoLine(Line As Long)
'On Error Resume Next
'StartPos = SendMessageLong(rtfCodebox.hWnd, EM_LINEINDEX, Line - 1, 0&)
'
'rtfCodebox.SelStart = StartPos
'rtfCodebox.SelLength = 0
'End Sub


Sub ScrollPage(txtA As Object, txtB As Object, Optional up As Boolean = False)
        
    Dim cnt As Long, topA As Long, topB As Long
    cnt = VisibleLines(txtA) - 1
    
    topA = TopLineIndex(txtA)
    topB = TopLineIndex(txtB)
    
    ScrollToLine txtA, topA + IIf(up, cnt, -cnt)
    ScrollToLine txtB, topB + IIf(up, cnt, -cnt)
    
End Sub

Function TopLineIndex(x As Object) As Long
    TopLineIndex = SendMessage(x.hwnd, EM_GETFIRSTVISIBLELINE, 0, ByVal 0&)
End Function

Function VisibleLines(x As Object) As Long
    Dim udtRect As Rect, tm As TEXTMETRIC
    Dim hdc As Long, lFont As Long, lOrgFont As Long
    Const WM_GETFONT As Long = &H31
    
    SendMessage x.hwnd, EM_GETRECT, 0, udtRect

    lFont = SendMessage(x.hwnd, WM_GETFONT, 0, 0)
    hdc = GetDC(x.hwnd)

    If lFont <> 0 Then
        lOrgFont = SelectObject(hdc, lFont)
    End If

    GetTextMetrics hdc, tm
    
    If lFont <> 0 Then
        lFont = SelectObject(hdc, lOrgFont)
    End If

    VisibleLines = (udtRect.Bottom - udtRect.Top) \ tm.tmHeight

    ReleaseDC x.hwnd, hdc

End Function

Sub ScrollToLine(t As Object, x As Integer)
     x = x - TopLineIndex(t)
     ScrollIncremental t, , x
End Sub

Sub ScrollIncremental(t As Object, Optional horz As Integer = 0, Optional vert As Integer = 0)
    'lParam&  The low-order 2 bytes specify the number of vertical
    '          lines to scroll. The high-order 2 bytes specify the
    '          number of horizontal columns to scroll. A positive
    '          value for lParam& causes text to scroll upward or to the
    '          left. A negative value causes text to scroll downward or
    '          to the right.
    ' r&       Indicates the number of lines actually scrolled.
    
    Dim r As Long
    r = CLng(&H10000 * horz) + vert
    r = SendMessage(t.hwnd, EM_LINESCROLL, 0, ByVal r)

End Sub

Sub SelBackcolor(rtb As RichTextBox, lngColor As Long)
    Dim tCF2 As CHARFORMAT2
    tCF2.dwMask = CFM_BACKCOLOR
    tCF2.crBackColor = lngColor
    tCF2.cbSize = Len(tCF2)
    SendMessage rtb.hwnd, EM_SETCHARFORMAT, 1, tCF2
End Sub

Sub DynamicHighLight(mrtb As RichTextBox, word As String, cSel As Collection, Optional backColor As Long, Optional foreColor As Long, Optional doBold As Boolean)
 
    Dim lngPos As Long
    Dim ss As Long
    Dim sl As Long
    Dim cs As CSelection
    Dim topLine As Long
    
    If working Then Exit Sub
    
    working = True
    LockWindowUpdate mrtb.hwnd
    
    'save current selection offsets
    topLine = TopLineIndex(mrtb)
    ss = mrtb.selStart
    sl = mrtb.selLength
    
    'remove all previous selections from display..
    For Each cs In cSel
        cs.Undo mrtb
    Next
    
    Set cSel = New Collection
     
    lngPos = mrtb.Find(word, 0)
     
    While lngPos > 0
        
        Set cs = New CSelection
        cs.LoadSel mrtb
        cSel.Add cs
        
        If foreColor > 0 Then mrtb.SelColor = foreColor
        If doBold = True Then mrtb.SelBold = True
        SelBackcolor mrtb, backColor
        
        lngPos = mrtb.Find(word, lngPos + mrtb.selLength)
        
    Wend
     
    mrtb.selStart = ss
    mrtb.selLength = sl
    ScrollToLine mrtb, CInt(topLine)
    
    LockWindowUpdate 0
    working = False
    
End Sub

