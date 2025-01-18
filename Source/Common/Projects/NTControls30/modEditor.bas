Attribute VB_Name = "modEditor"
Option Explicit

Public Type RangeType
    StartPos As Long
    StopPos As Long
End Type

Public Type TextRange
    Range As RangeType
    lpStr As Long
End Type

Public Type ColorRange
    StartLoc As Long
    Forecolor As Long
    BackColor As Long
End Type

Public Type UndoType
    CodePage As Long
    PriorTextData As NTNodes10.Strands
    AfterTextData As NTNodes10.Strands
    PriorSelRange As RangeType
    AfterSelRange As RangeType
End Type

Public Enum TextModes
    TM_PLAINTEXT = 1
    TM_RICHTEXT = 2
    TM_SINGLELEVELUNDO = 4
    TM_MULTILEVELUNDO = 8
    TM_SINGLECODEPAGE = 16
    TM_MULTICODEPAGE = 32
End Enum

Public Type SetTextEx
    flags As Long
    CodePage As Long
End Type
'
'Public Type GetTextEx
'    cb As Long
'    flags As Long
'    codepage As Long
'    lpDefaultChar As Long
'    lpUsedDefChar As Long
'End Type
'
'Public Type GetTextLengthEx
'    flags As Long
'    codepage As Long
'End Type

Public Const ST_DEFAULT = 0
Public Const ST_KEEPUNDO = 1
Public Const ST_SELECTION = 2
Public Const ST_NEWCHARS = 3
Public Const ST_UNICODE = 4


Public Const GT_DEFAULT = 0
Public Const GT_USECRLF = 1
Public Const GT_SELECTION = 2

Public Const GTL_DEFAULT = 0
Public Const GTL_USECRLF = 1
Public Const GTL_PRECISE = 2
Public Const GTL_CLOSE = 4
Public Const GTL_NUMCHARS = 8
Public Const GTL_NUMBYTES = 16

Public Const SB_BOTH = 3
Public Const SB_BOTTOM = 7
Public Const SB_CTL = 2
Public Const SB_ENDSCROLL = 8
Public Const SB_HORZ = 0
Public Const SB_LEFT = 6
Public Const SB_LINEDOWN = 1
Public Const SB_LINELEFT = 0
Public Const SB_LINERIGHT = 1
Public Const SB_LINEUP = 0
Public Const SB_PAGEDOWN = 3
Public Const SB_PAGELEFT = 2
Public Const SB_PAGERIGHT = 3
Public Const SB_PAGEUP = 2
Public Const SB_RIGHT = 7
Public Const SB_TOP = 6
Public Const SB_THUMBPOSITION = 4
Public Const SB_THUMBTRACK = 5
Public Const SB_VERT = 1

Public Const SC_VSCROLL = &HF070&
Public Const SC_HSCROLL = &HF080&


Public Const WM_COMMAND = &H111
Public Const WM_ERASEBKGND = &H14

Public Const WM_USER = &H400

Public Const WM_MOUSEMOVE = &H200


Public Const WM_MOUSEWHEEL = &H20A ' window message for mouse wheel
Public Const WM_VSCROLL = &H115
Public Const WM_HSCROLL = &H114
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CHAR = &H102
Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const WM_CLEAR = &H303
Public Const WM_PAINT = &HF
Public Const WM_GETTEXT = &HD
Public Const EM_GETSEL = &HB0&
Public Const EM_SETSEL = &HB1&
Public Const EM_CANUNDO = &HC6&
Public Const EM_UNDO = &HC7&
Public Const EM_REDO = (WM_USER + 84)
Public Const EM_CANREDO = (WM_USER + 85)

Public Const WM_UNDO = &H304

Public Const WM_SETCURSOR = &H20
Public Const WM_MOUSEACTIVATE = &H21

Public Const GWL_WNDPROC = (-4)

Public Const EN_HSCROLL = &H601
Public Const EN_VSCROLL = &H602
Public Const EM_EXGETSEL = (WM_USER + 52)
Public Const EM_SCROLL As Long = &HB5
Public Const EM_GETLINECOUNT As Long = &HBA
Public Const EM_LINESCROLL = &HB6
Public Const EM_GETLINE = &HC4&
Public Const EM_LINEINDEX = &HBB&
Public Const EM_LINELENGTH = &HC1&
Public Const EM_SETTEXTMODE = (WM_USER + 89)
Public Const EM_GETSCROLLPOS = (WM_USER + 221)
Public Const EM_SETSCROLLPOS = (WM_USER + 222)
Public Const EM_GETFIRSTVISIBLELINE = &HCE&
Public Const EM_EXLINEFROMCHAR = (WM_USER + 54)
Public Const EM_GETTEXTRANGE = (WM_USER + 75)
Public Const EM_SCROLLCARET = &HB7&
Public Const EM_CHARFROMPOS = &HD7&
Public Const EM_POSFROMCHAR = &HD6&
Public Const EM_EXSETSEL = (WM_USER + 55)
Public Const EM_SETTEXTEX = (WM_USER + 97)
Public Const EM_GETTEXTLENGTHEX = (WM_USER + 95)
Public Const EM_HIDESELECTION = (WM_USER + 63)
Public Const EM_SETUNDOLIMIT = (WM_USER + 82)
Public Const EM_LINEFROMCHAR = &HC9&

Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2
Public Const VK_MBUTTON = &H4

Public Const VK_LSHIFT = &HA0
Public Const VK_RSHIFT = &HA1

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer 'Gets state of one key
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Declare Function GetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long) As Long

Public Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageStruct Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLngPtr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageMemory Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, wParam As Any) As Variant

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
'Public Declare Function PtInRect Lib "user32" (lpRect As Rect, ByVal lLeft As Long, ByVal lTop As Long) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal lLeft As Long, ByVal lTop As Long) As Long


Public Declare Function RectInRegion Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Any) As Long

Public Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Dest As Any, ByRef Source As Any, ByVal Length As Long)

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_GETKEYBOARDSPEED = 10
Public Const SPI_GETKEYBOARDDELAY = 22
Public Const SPI_SETKEYBOARDSPEED = 11
Public Const SPI_SETKEYBOARDDELAY = 23
Public Const SPIF_SENDCHANGE = &H2

Public Const SM_CYSIZE = 31&   'Titlebar height
Public Const SM_CXSIZEFRAME = 32&
Public Const SM_CYSIZEFRAME = 33&

Public Const SM_CXBORDER = 5&  'Borders width
Public Const SM_CYCAPTION = 4& 'caption height

Public Const SIF_RANGE = &H1
Public Const SIF_PAGE = &H2
Public Const SIF_POS = &H4
Public Const SIF_DISABLENOSCROLL = &H8
Public Const SIF_TRACKPOS = &H10
Public Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)


Public Declare Function SetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Public Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal N As Long, lpScrollInfo As SCROLLINFO) As Long
Public Declare Function GetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
Public Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal N As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
Public Declare Function SetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal bRedraw As Long) As Long
Public Declare Function ScrollWindow Lib "user32" (ByVal hWnd As Long, ByVal XAmount As Long, ByVal YAmount As Long, lpRect As RECT, lpClipRect As RECT) As Long


Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const MOUSE_MOVED = &H1

Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Public Const LOGPIXELSX = 88
Private Const POINTS_PER_INCH As Long = 100

Public Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

Private SubClassed As New VBA.Collection

Public Function UndoSize(ByRef InArray() As UndoType, Optional ByVal InBytes As Boolean = False) As Long
On Error GoTo dimerror

    Static dimcheck As Long

    If UBound(InArray) = -1 Or LBound(InArray) = -1 Then
        UndoSize = 0
    Else
        UndoSize = (UBound(InArray) + -CInt(Not CBool(-LBound(InArray)))) * IIf(InBytes, LenB(InArray(LBound(InArray))), 1)
    End If
    Exit Function
startover:
    Err.Clear
    On Error GoTo -1
    On Error GoTo 0
    On Error GoTo dimerror
    If UBound(InArray, dimcheck) = -1 Or LBound(InArray, dimcheck) = -1 Then
        UndoSize = 0
    Else
        UndoSize = (UBound(InArray, dimcheck) + -CInt(Not CBool(-LBound(InArray, dimcheck)))) * IIf(InBytes, LenB(InArray(LBound(InArray, dimcheck), LBound(InArray, dimcheck - 1))), 1)
    End If
    
    Exit Function
dimerror:
    If dimcheck = 0 Then
        dimcheck = 2
        Err.Clear
        GoTo startover
    End If
    UndoSize = 0
End Function

Public Sub Hook(ByRef UserControl As IControl)
    If ((IsCompiled Or IsRunMode) Or ((Not IsCompiled) And IsRunningMode)) Then
        If UserControl.hProc = 0 Then
            SubClassed.Add UserControl, "H" & UserControl.hWnd
            UserControl.hProc = GetWindowLong(UserControl.hWnd, GWL_WNDPROC)
            SetWindowLong UserControl.hWnd, GWL_WNDPROC, AddressOf WinProc
        End If
    Else
       Unhook UserControl
    End If
End Sub

Public Sub Unhook(ByRef UserControl As IControl)
    If UserControl.hProc <> 0 Then
        SetWindowLong UserControl.hWnd, GWL_WNDPROC, UserControl.hProc
        UserControl.hProc = 0
    End If
    If SubClassed.Count > 0 Then
        SubClassed.Remove "H" & UserControl.hWnd
    End If

End Sub

Public Function PT(X, Y) As POINTAPI
    PT.X = CLng(X)
    PT.Y = CLng(Y)
End Function

Public Function Range(ByVal StartPos As Long, ByVal StopPos As Long) As RangeType
    Range.StartPos = StartPos
    Range.StopPos = StopPos
End Function

Public Function RemoveNextArg(ByRef TheParams As Variant, ByVal TheSeperator As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare, Optional ByVal TrimResult As Boolean = True) As String
    If TrimResult Then
        If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
            RemoveNextArg = Trim(Left(TheParams, InStr(1, TheParams, TheSeperator, Compare) - 1))
            TheParams = Trim(Mid(TheParams, InStr(1, TheParams, TheSeperator, Compare) + Len(TheSeperator)))
        Else
            RemoveNextArg = Trim(TheParams)
            TheParams = ""
        End If
    Else
        If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
            RemoveNextArg = Left(TheParams, InStr(1, TheParams, TheSeperator, Compare) - 1)
            TheParams = Mid(TheParams, InStr(1, TheParams, TheSeperator, Compare) + Len(TheSeperator))
        Else
            RemoveNextArg = TheParams
            TheParams = ""
        End If
    End If
End Function

Public Function RangeLength(ByRef Range As RangeType) As Long
    RangeLength = IIf((Range.StopPos >= Range.StartPos), (Range.StopPos - Range.StartPos), IIf((Range.StartPos = -1) And (Range.StopPos = 0), -1, 0))
End Function

Public Function GetTextRange(ByVal hWnd As Long, ByVal StartPos As Long, ByVal StopPos As Long) As String
    If (StopPos - StartPos) > 0 Then
        Dim txt As TextRange
        Dim str As String
        str = Space((StopPos - StartPos))
        txt.Range.StartPos = StartPos
        txt.Range.StopPos = StopPos
        RtlMoveMemory txt.lpStr, ByVal VarPtr(str), 4&
        If SendMessageLngPtr(hWnd, EM_GETTEXTRANGE, 0, ByVal VarPtr(txt)) Then
            RtlMoveMemory ByVal VarPtr(str), txt.lpStr, 4&
            GetTextRange = Left(StrConv(str, vbUnicode), StopPos - StartPos)
        End If
    End If
End Function

'Public Sub SetScrollbarsByVisibility(ByRef frm As TextBox, Optional ByVal SelStartWithLength As Boolean = True)
'    'frm.VScroll.AutoRedraw = False
'    'frm.HScroll.AutoRedraw = False
'
'    Dim lLength As Long
'    Dim lLine As Long
'    Dim textSel As POINTAPI
'    Dim lIndex  As Long
'    Dim useCursor As Long
'    Dim txt As String
'    lLine = SendMessageLngPtr(frm.RichTextBox.hWnd, EM_GETFIRSTVISIBLELINE, 0, 0&)
'    If lLine > frm.VScroll.max Then
'        If frm.VScroll.Value <> frm.VScroll.max Then
'            frm.VScroll.Value = frm.VScroll.max
'        End If
'    ElseIf frm.VScroll.Value <> lLine Then
'        frm.VScroll.Value = lLine
'    End If
'    SendMessageStruct frm.RichTextBox.hWnd, EM_EXGETSEL, 0, textSel
'    useCursor = IIf(SelStartWithLength, textSel.Y, textSel.X)
'    lLine = SendMessageLngPtr(frm.RichTextBox.hWnd, EM_EXLINEFROMCHAR, 0, useCursor)
'    lIndex = SendMessageLngPtr(frm.RichTextBox.hWnd, EM_LINEINDEX, lLine, 0&)
'    lLength = SendMessageLngPtr(frm.RichTextBox.hWnd, EM_LINELENGTH, useCursor, 0&)
'    txt = GetTextRange(frm.RichTextBox.hWnd, lIndex, lIndex + lLength)
'    lLength = frm.TextWidth(Left(txt, useCursor - lIndex))
'    If frm.TextWidth(Left(txt, useCursor - lIndex)) - frm.HScroll.Value < 0 Or _
'        frm.TextWidth(Left(txt, useCursor - lIndex)) - frm.HScroll.Value > frm.RichTextBox.Width Then
'        lLength = lLength - (frm.RichTextBox.Width \ 4)
'        If lLength >= frm.HScroll.Min And lLength <= frm.HScroll.max Then
'            If frm.HScroll.Value <> lLength Then
'                frm.HScroll.Value = lLength
'            End If
'        ElseIf lLength > frm.HScroll.max Then
'            If frm.HScroll.Value <> frm.HScroll.max Then
'                frm.HScroll.Value = frm.HScroll.max
'            End If
'        ElseIf lLength < frm.HScroll.Min Then
'            If frm.HScroll.Value <> frm.HScroll.Min Then
'                frm.HScroll.Value = frm.HScroll.Min
'            End If
'        End If
'    End If
'
'    'frm.VScroll.AutoRedraw = True
'    'frm.HScroll.AutoRedraw = True
'End Sub
Public Sub AdjustScollbarValue(ByRef ScrollBar As ScrollBar, ByVal Offset As Long)
    'ScrollBar.AutoRedraw = False
    
    If Offset > 0 Then
        If ScrollBar.Value + Offset <= ScrollBar.Max Then
            If ScrollBar.Value <> ScrollBar.Value + Offset Then
                ScrollBar.Value = ScrollBar.Value + Offset
            End If
        ElseIf ScrollBar.Value <> ScrollBar.Max Then
            ScrollBar.Value = ScrollBar.Max
        End If
    ElseIf Offset < 0 Then
        If ScrollBar.Value + Offset >= ScrollBar.Min Then
            If ScrollBar.Value <> ScrollBar.Value + Offset Then
                ScrollBar.Value = ScrollBar.Value + Offset
            End If
        ElseIf ScrollBar.Value <> ScrollBar.Min Then
            ScrollBar.Value = ScrollBar.Min
        End If
    End If
    
    'ScrollBar.AutoRedraw = True
End Sub

Private Function WinProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim lDispatch As Boolean
    lDispatch = True
    
    If SubClassed.Count > 0 Then
        Dim frm2 As ScrollBar
        Dim frm3 As TextBox
        Dim PT As POINTAPI
        Dim st As SetTextEx
        Dim txt As TextRange
        Dim str As String
                            
        Dim cnt As Long

        
        
        For cnt = 1 To SubClassed.Count
        

                If TypeName(SubClassed(cnt)) = "ScrollBar" Then
                    Set frm2 = SubClassed(cnt)
                    If frm2.hWnd = hWnd Then

                        Select Case wMsg

                                
                            Case WM_MOUSEWHEEL
    
                                AdjustScollbarValue frm2, ((frm2.LargeChange - frm2.SmallChange) * 2) * IIf(wParam > 0, -1, 1)
    
                            Case WM_VSCROLL

                            Case WM_HSCROLL

                            Case EM_GETSCROLLPOS

                            Case EM_SETSCROLLPOS
                            

                            Case WM_PAINT
    
                                If frm2.AutoRedraw Then frm2.PaintBuffer
                               
                                
                                WinProc = 1
                                lDispatch = False
                                DefWindowProc hWnd, wMsg, wParam, lParam
                                
                            Case Else
                                lDispatch = True
                        End Select
            
                        If (lDispatch And (WinProc = 0)) Then
                            WinProc = CallWindowProc(frm2.hProc, hWnd, wMsg, wParam, lParam)
                        End If
        
                    End If
                    Set frm2 = Nothing
                    
                ElseIf TypeName(SubClassed(cnt)) = "TextBox" Then
                    
                    Set frm3 = SubClassed(cnt)
                    If frm3.hWnd = hWnd Then
                    
                        Select Case wMsg
                            Case WM_MOUSEWHEEL
                                'Debug.Print "WM_MOUSEWHEEL"; GetAsyncKeyState(VK_LSHIFT); GetAsyncKeyState(VK_RSHIFT)
                                If frm3.Enabled Then
                                    If GetAsyncKeyState(VK_LSHIFT) < 0 Or GetAsyncKeyState(VK_RSHIFT) < 0 Then
                                        AdjustScollbarValue frm3.HScroll, ((frm3.HScroll.LargeChange - frm3.HScroll.SmallChange) * 2) * IIf(wParam > 0, -1, 1)
                                    Else
                                        AdjustScollbarValue frm3.VScroll, ((frm3.VScroll.LargeChange - frm3.VScroll.SmallChange) * 2) * IIf(wParam > 0, -1, 1)
                                    End If
                                End If
                                lDispatch = False
                            Case WM_VSCROLL
                                'Debug.Print "WM_VSCROLL"; wParam; lParam
                                frm3.SetScrollBars
                                lDispatch = False
                            Case WM_HSCROLL
                                'Debug.Print "WM_HSCROLL"; wParam; lParam
                                frm3.SetScrollBars
                                lDispatch = False
                            Case WM_COPY
                                frm3.Copy
                                lDispatch = False
                            Case WM_PASTE
                                frm3.Paste
                                lDispatch = False
                            Case WM_CUT
                                frm3.Cut
                                lDispatch = False
                            Case WM_CLEAR
                                frm3.Delete
                                lDispatch = False
                            Case EM_CANUNDO
                                WinProc = -CInt(frm3.CanUndo) And frm3.Enabled
                                lDispatch = False
                            Case EM_CANREDO
                                WinProc = -CInt(frm3.CanUndo) And frm3.Enabled
                                lDispatch = False
                            Case EM_GETSEL
                                If frm3.Enabled Then
                                    PT.X = frm3.SelStart
                                    PT.Y = frm3.SelStart + frm3.SelLength
                                    CopyMemory ByVal wParam, ByVal VarPtr(PT.X), 4&
                                    CopyMemory ByVal lParam, ByVal VarPtr(PT.Y), 4&
                                End If
                                lDispatch = False
                            Case EM_SETSEL
                                If frm3.Enabled Then
                                    CopyMemory ByVal VarPtr(PT.X), ByVal wParam, 4&
                                    CopyMemory ByVal VarPtr(PT.Y), ByVal lParam, 4&
                                    frm3.SelStart = PT.X
                                    frm3.SelLength = PT.Y - PT.X
                                End If
                                lDispatch = False
                            Case EM_EXGETSEL
                                If frm3.Enabled Then
                                    CopyMemory ByVal VarPtr(PT), ByVal lParam, LenB(PT)
                                    PT.X = frm3.SelStart
                                    PT.Y = frm3.SelStart + frm3.SelLength
                                    CopyMemory ByVal lParam, ByVal VarPtr(PT), LenB(PT)
                                    WinProc = 1
                                End If
                                lDispatch = False
                            Case EM_EXSETSEL
                                If frm3.Enabled Then
                                    CopyMemory ByVal VarPtr(PT), ByVal lParam, LenB(PT)
                                    frm3.SelStart = PT.X
                                    frm3.SelLength = PT.Y - PT.X
    
                                    WinProc = 1
                                End If
                                lDispatch = False
                            Case EM_GETSCROLLPOS
                                If frm3.Enabled Then
                                    CopyMemory ByVal VarPtr(PT), ByVal lParam, LenB(PT)
                                    PT.X = -frm3.OffsetX
                                    PT.Y = -frm3.OffsetY
                                    CopyMemory ByVal lParam, ByVal VarPtr(PT), LenB(PT)
                                    WinProc = 1
                                End If
                                lDispatch = False
                            Case EM_SETSCROLLPOS
                                If frm3.Enabled Then
                                    CopyMemory ByVal VarPtr(PT), ByVal lParam, LenB(PT)
                                    frm3.OffsetX = -PT.X
                                    frm3.OffsetY = -PT.Y
                                    WinProc = 1
                                End If
                                lDispatch = False
                            Case EM_GETLINECOUNT
    
                                WinProc = frm3.LineCount
    
                                lDispatch = False
                            Case EM_CHARFROMPOS
    
                                CopyMemory ByVal VarPtr(PT), ByVal lParam, LenB(PT)
                                WinProc = frm3.CaretFromPoint(PT.X, PT.Y)
    
                                lDispatch = False
                                
                            Case EM_EXLINEFROMCHAR
                                
                                WinProc = frm3.LineIndex(lParam)
    
                                lDispatch = False
                                
                                
                            Case EM_GETFIRSTVISIBLELINE
                                lDispatch = False
                                WinProc = frm3.LineFirstVisible
                            Case EM_LINEINDEX
                                lDispatch = False
                                
                                WinProc = frm3.LineOffset(wParam)
    
                            Case EM_LINELENGTH
                                lDispatch = False
    
                                WinProc = frm3.LineLength(wParam)
    
                            Case EM_GETTEXTLENGTHEX
                                lDispatch = False
                                WinProc = frm3.Length
                            Case EM_SETTEXTMODE
                                If ((wParam And TM_MULTICODEPAGE) = TM_MULTICODEPAGE) Then
                                    frm3.ClearSeperators
                                ElseIf ((wParam And TM_SINGLECODEPAGE) = TM_SINGLECODEPAGE) Then
                                    frm3.Reset
                                End If
                                
                                If ((wParam And TM_SINGLELEVELUNDO) = TM_SINGLELEVELUNDO) Then
                                    frm3.UndoLimit = 1
                                ElseIf ((wParam And TM_MULTILEVELUNDO) = TM_MULTILEVELUNDO) Then
                                    frm3.UndoLimit = -1
                                End If

                                If ((wParam And TM_PLAINTEXT) = TM_PLAINTEXT) Then
                                ElseIf ((wParam And TM_RICHTEXT) = TM_RICHTEXT) Then
                                End If
                                
                                
                            Case EM_SETTEXTEX
                                   
                                CopyMemory ByVal VarPtr(st), ByVal wParam, LenB(st)
                                str = StringANSI(lParam)
                                If ((st.flags And ST_UNICODE) = ST_UNICODE) Then
                                    str = StrConv(str, vbFromUnicode)
                                End If
                                
                                txt.Range.StartPos = frm3.SelStart
                                txt.Range.StopPos = frm3.SelLength
                                
                                frm3.Cancel = True
                                frm3.CodePage = (st.CodePage + 1)
                                
                                If ((st.flags And ST_SELECTION) = ST_SELECTION) Then
                                    txt.lpStr = -1
                                Else
                                    txt.lpStr = 0
                                    If (Not ((st.flags And ST_NEWCHARS) = ST_NEWCHARS)) Then
                                        frm3.SelLength = Len(str)
                                    Else
                                        frm3.SelLength = 0
                                    End If
                                End If
    
                                frm3.SelText = str
    
                                If ((st.flags And ST_SELECTION) = ST_SELECTION) Then
                                    frm3.SelStart = txt.Range.StartPos
                                    frm3.SelLength = Len(str)
                                Else
                                    frm3.SelStart = txt.Range.StartPos
                                    frm3.SelLength = txt.Range.StopPos
                                End If
                                
                                frm3.Cancel = False

                                frm3.RaiseEventChange ((st.flags And ST_KEEPUNDO) = ST_KEEPUNDO)
                                
                                lDispatch = False
                                
                            Case EM_SETUNDOLIMIT
                                frm3.UndoLimit = wParam
                                lDispatch = False
    
                            Case EM_GETTEXTRANGE
    
                                CopyMemory ByVal VarPtr(txt), ByVal lParam, LenB(txt)
                                If txt.Range.StartPos >= 0 And txt.Range.StartPos < frm3.Length And txt.Range.StopPos >= txt.Range.StopPos Then
                                    str = frm3.Text.Partial(txt.Range.StartPos, txt.Range.StopPos - txt.Range.StartPos)
                                    txt.lpStr = StrPtr(str)
                                    CopyMemory ByVal lParam, ByVal VarPtr(txt), LenB(txt)
                                End If
                                WinProc = Len(str)
    
                                lDispatch = False
                                
                            Case WM_PAINT
                                
                                If frm3.AutoRedraw Then
                                   ' frm3.Paint
                                    frm3.PaintBuffer
                                End If
    
                                WinProc = 0
                                lDispatch = False
                                DefWindowProc hWnd, wMsg, wParam, lParam
    
                            Case Else
                                lDispatch = True
                        End Select
            
                        If (lDispatch And (WinProc = 0)) Then
                            WinProc = CallWindowProc(frm3.hProc, hWnd, wMsg, wParam, lParam)
                        End If
                    End If
                    Set frm3 = Nothing
                End If

        Next
    Else
        SetWindowLong hWnd, GWL_WNDPROC, ByVal 0&
    End If
    
End Function

Public Function GetLineText(ByVal hWnd As Long, ByVal Index As Long) As String
    Dim xLine As RangeType
    Dim st As SetTextEx
    st.flags = ST_SELECTION And ST_NEWCHARS
    xLine.StartPos = SendMessageLngPtr(hWnd, EM_LINEINDEX, Index, 0&)
    If xLine.StartPos = -1 Then Exit Function
    xLine.StopPos = xLine.StartPos + SendMessageLngPtr(hWnd, EM_LINELENGTH, xLine.StartPos, 0&)
    GetLineText = Replace(Replace(GetTextRange(hWnd, xLine.StartPos, xLine.StopPos), vbCr, ""), vbLf, "")

End Function

Public Sub SetLineText(ByVal hWnd As Long, ByVal Index As Long, ByVal Text As String)
    Dim xLine As RangeType
    Dim st As SetTextEx
    
    st.flags = ST_SELECTION Or ST_NEWCHARS

    xLine.StartPos = SendMessageLngPtr(hWnd, EM_LINEINDEX, Index, 0&)
    
    If xLine.StartPos = -1 Then Exit Sub
    xLine.StopPos = xLine.StartPos + SendMessageLngPtr(hWnd, EM_LINELENGTH, xLine.StartPos, 0&)
    
    SendMessageStruct hWnd, EM_EXSETSEL, 0, xLine
    SendMessageString hWnd, EM_SETTEXTEX, ByVal VarPtr(st), Text
    
End Sub
