VERSION 5.00
Begin VB.UserControl TextBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3780
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   MousePointer    =   3  'I-Beam
   ScaleHeight     =   1800
   ScaleWidth      =   3780
   ToolboxBitmap   =   "TextBox.ctx":0000
   Begin NTControls30.ScrollBar ScrollBar2 
      Height          =   285
      Left            =   195
      Top             =   1185
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   503
   End
   Begin NTControls30.ScrollBar ScrollBar1 
      Height          =   1290
      Left            =   3105
      Top             =   75
      Visible         =   0   'False
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   2275
      Orientation     =   0
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   75
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "De&lete"
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select &All"
      End
   End
End
Attribute VB_Name = "TextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum vbScrollBars
    Auto = -1
    None = 0
    Horizontal = 1
    Vertical = 2
    Both = 3
End Enum

Public Event Change()
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLECompleteDrag(Effect As Long)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Public Event Paint()
Public Event Resize()
Public Event SelChange()

Private dragStart As Integer
Private keySpeed As Long
Private hasFocus As Boolean
Private insertMode As Boolean

Private pOffsetX As Long
Private pOffsetY As Long

Private pHideSelection As Boolean
Private pScrollToCaret As Boolean
Private pMultiLine As Boolean
Private pForeColor As OLE_COLOR
Private pBackColor As OLE_COLOR
Private pSel As RangeType
Private pScrollBars As vbScrollBars

Private pOldProc As Long

Private pText As String
'Private pText As NTNodes10.Stream

Friend Property Get VScroll() As ScrollBar
    Set VScroll = ScrollBar1
End Property
Friend Property Get HScroll() As ScrollBar
    Set HScroll = ScrollBar2
End Property

Friend Property Get OldProc() As Long
    OldProc = pOldProc
End Property
Friend Property Let OldProc(ByVal RHS As Long)
    pOldProc = RHS
End Property

Private Function DrawableRect() As Rect
    DrawableRect = Rect(0, 0, UsercontrolWidth, UsercontrolHeight)
End Function
Private Property Get UsercontrolWidth() As Long
    UsercontrolWidth = IIf(ScrollBar1.Visible, (UserControl.Width - ScrollBar1.Width), UserControl.Width)
    If UsercontrolWidth < 0 Then UsercontrolWidth = 0
End Property
Private Property Get UsercontrolHeight() As Long
    UsercontrolHeight = IIf(ScrollBar2.Visible, (UserControl.Height - ScrollBar2.Height), UserControl.Height)
    If UsercontrolHeight < 0 Then UsercontrolHeight = 0
End Property

Public Property Get hWnd() As Long ' _
Returns the standard windows handle of the control.
    hWnd = UserControl.hWnd
End Property

Public Property Get ScrollBars() As vbScrollBars ' _
Gets the value that determines how scroll bars are used and displayed for the control.
    ScrollBars = pScrollBars
End Property
Public Property Let ScrollBars(ByVal RHS As vbScrollBars) ' _
Sets the behavior of and whether or not scroll bars are used for the control.
    pScrollBars = RHS
    SetScrollBars
    UserControl_Paint
End Property

Public Property Get TextHeight(Optional ByVal StrText As String = "Iy") As Long ' _
Returns the twip measurement height using the current font size and line spacing vertically of SrrText.
    TextHeight = UserControl.TextHeight(StrText)
End Property
Public Property Get TextWidth(Optional ByVal StrText As String = "W") As Long ' _
Returns the twip measurement width using the current font size and letter spacing horizontally of SrrText.
    TextWidth = UserControl.TextWidth(StrText)
End Property
Public Property Get MultipleLines() As Boolean ' _
Returns whether or not this text control allows multiple lines in Text, delimited by line feeds.
    MultipleLines = pMultiLine
End Property
Public Property Let MultipleLines(ByVal RHS As Boolean) ' _
Sets whehter or not this text control allows multiple lines in Text, delimited by line feeds.
    pMultiLine = RHS
    If Not pMultiLine Then
        'Dim tText As New Stream
        'tText.Concat Convert(Replace(Convert(pText.Partial), vbCr, ""))
        pText = Replace(pText, vbLf, "")
        'Set pText = tText
        pOffsetY = 0
    End If
End Property
Public Property Get HideSelection() As Boolean ' _
Gets whether or not the selection highlight will be hidden when the control is not in focus.
    HideSelection = pHideSelection
End Property
Public Property Let HideSelection(ByVal RHS As Boolean) ' _
Sets whether or not the selection highlight will be hidden when the control is not in focus.
    pHideSelection = RHS
End Property

Public Property Get ScrollToCaret() As Boolean ' _
Gets whether or not the caret forces the scrolling to keep it with in visibility.
    ScrollToCaret = pScrollToCaret
End Property
Public Property Let ScrollToCaret(ByVal RHS As Boolean) ' _
Sets whether or not the caret forces the scrolling to keep it with in visibility.
    pScrollToCaret = RHS
End Property

Public Property Get CanvasWidth() As Long ' _
Gets the width of the textual space needed and defined by the Text and Width property.
    CanvasWidth = UserControl.TextWidth(pText) + (UsercontrolWidth / 2)
End Property
Public Property Get CanvasHeight() As Long ' _
Gets the height of the textual space needed and defined by the Text and Height property.
    CanvasHeight = UserControl.TextHeight(pText) + (UsercontrolHeight / 2)
End Property

Public Property Get OffsetX() As Long ' _
Gets the current scroll bar offset of the horizontal canvas width drawn in visibility.
    OffsetX = pOffsetX
End Property
Public Property Let OffsetX(ByVal RHS As Long) ' _
Sets the current scroll bar offset of the horizontal canvas width drawn in visibility.
    pOffsetX = RHS
    UserControl_Paint
End Property

Public Property Get OffsetY() As Long ' _
Gets the current scroll bar offset of the vertical canvas height drawn in visibility.
    OffsetY = pOffsetY
End Property
Public Property Let OffsetY(ByVal RHS As Long) ' _
Sets the current scroll bar offset of the vertical canvas height drawn in visibility.
    pOffsetY = RHS
    UserControl_Paint
End Property

Private Sub CanvasValidate()
    If pOffsetX > 0 Or UsercontrolWidth > CanvasWidth Then pOffsetX = 0
    If pOffsetY > 0 Or UsercontrolHeight > CanvasHeight Then pOffsetY = 0
   ' If pOffsetX < -(CanvasWidth - UsercontrolWidth) - 1 Then pOffsetX = -(CanvasWidth - UsercontrolWidth) - 1
   ' If pOffsetY < -(CanvasHeight - UsercontrolHeight) - 1 Then pOffsetY = -(CanvasHeight - UsercontrolHeight) - 1
End Sub

Public Property Get Forecolor() As OLE_COLOR ' _
Gets the default forecolor of the text display when a specific color table coloring is not used.
    Forecolor = pForeColor
End Property
Public Property Let Forecolor(ByVal RHS As OLE_COLOR) ' _
Sets the default forecolor of the text display when a specific color table coloring is not used.
    pForeColor = RHS
    UserControl_Paint
End Property

Public Property Get Backcolor() As OLE_COLOR ' _
Gets the default background color of the text display when a specific color table coloring is not used.
    Backcolor = pBackColor
End Property
Public Property Let Backcolor(ByVal RHS As OLE_COLOR) ' _
Sets the default background color of the text display when a specific color table coloring is not used.
    pBackColor = RHS
    UserControl.Backcolor = RHS
    UserControl_Paint
End Property

Public Property Get SelText() As String ' _
Gets the selected text, the portion of text that is highlighted.
    If pSel.StopPos < pSel.StartPos Then
        SelText = Mid(pText, pSel.StopPos + 1, (pSel.StartPos - pSel.StopPos))
    Else
        SelText = Mid(pText, pSel.StartPos + 1, (pSel.StopPos - pSel.StartPos))
    End If
    SelText = Replace(SelText, vbLf, vbCrLf)
End Property

Public Property Let SelText(ByVal RHS As String) ' _
Sets the selected text, the portion of text that is highlighted, changing the selection if applicable.
    If pSel.StartPos <> pSel.StopPos Then
        If pSel.StartPos > pSel.StopPos Then
            Swap pSel.StartPos, pSel.StopPos
        End If
        pText = Left(pText, pSel.StartPos) & RHS & Mid(pText, pSel.StopPos + 1)
        pSel.StopPos = pSel.StartPos + Len(RHS)
        Swap pSel.StartPos, pSel.StopPos
        InvalidateCursor
    End If
End Property

Public Property Get SelStart() As Long ' _
Gets the selection start of the highlighted portion of text.
    If pSel.StopPos < pSel.StartPos Then
        SelStart = pSel.StopPos
    Else
        SelStart = pSel.StartPos
    End If
End Property
Public Property Let SelStart(ByVal RHS As Long) ' _
Sets the selection start of the highlighted portion of text.
    If pSel.StopPos < pSel.StartPos Then
        pSel.StopPos = RHS
    Else
        pSel.StartPos = RHS
    End If
    InvalidateCursor
End Property
Public Property Get SelLength() As Long ' _
Gets the selection length of the highlighted portion of text.
    If pSel.StopPos < pSel.StartPos Then
        SelLength = pSel.StartPos - pSel.StopPos
    Else
        SelLength = pSel.StopPos - pSel.StartPos
    End If
End Property
Public Property Let SelLength(ByVal RHS As Long) ' _
Sets the selection length of the highlighted portion of text.
    If pSel.StopPos < pSel.StartPos Then
        pSel.StartPos = -(pSel.StartPos - pSel.StopPos) + RHS
    Else
        pSel.StopPos = -(pSel.StopPos - pSel.StartPos) + RHS
    End If
    InvalidateCursor
End Property

Public Property Get Text() As String ' _
Gets the text contents of the control in the string data type.
    Text = Replace(Replace(Replace(pText, vbCrLf, vbLf), vbLf, IIf(pMultiLine, vbLf, "")), vbLf, vbCrLf)
End Property
Public Property Let Text(ByVal RHS As String) ' _
Sets the text contents of the control by String data type.
    If pText <> RHS Then
        pText = Replace(Replace(RHS, vbCrLf, vbLf), vbLf, IIf(pMultiLine, vbLf, ""))
        RaiseEvent Change
        UserControl_Paint
    End If
End Property

Public Property Get Font() As StdFont ' _
Gets the font that the text is displayed in.
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByRef newVal As StdFont) ' _
Sets the font that the text is displayed in.
    Set UserControl.Font = newVal
    UserControl.FontBold = UserControl.Font.Bold
    UserControl.FontItalic = UserControl.Font.Italic
    UserControl.FontName = UserControl.Font.Name
    UserControl.FontSize = UserControl.Font.size
    UserControl.FontStrikethru = UserControl.Font.Strikethrough
    UserControl.FontUnderline = UserControl.Font.Underline
End Property




Private Sub ScrollBar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If pScrollToCaret Then
            dragStart = 3
            pScrollToCaret = False
        End If
    End If
End Sub

Private Sub ScrollBar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If dragStart = 3 Then
        dragStart = 0
        pScrollToCaret = True
    End If
End Sub

Private Sub ScrollBar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If pScrollToCaret Then
            dragStart = 3
            pScrollToCaret = False
        End If
    End If
End Sub

Private Sub ScrollBar2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If dragStart = 3 Then
        dragStart = 0
        pScrollToCaret = True
    End If
End Sub

Private Sub Timer1_Timer()
    Static cursorBlink As Boolean
    Static lastLoc As POINTAPI
    
    Dim newloc As POINTAPI
    newloc = CaretLocation
    
    If ((Not cursorBlink) Or (Not hasFocus)) Or ((newloc.X <> lastLoc.X) Or (newloc.Y <> lastLoc.Y)) Then
        If insertMode Then
            
            'ClipPrintText lastLoc.X, lastLoc.Y, IIf(Convert(pText.Partial(pSel.StartPos, 1)) = vbLf, " ", Convert(pText.Partial(pSel.StartPos, 1))), SystemColorConstants.vbWindowText, False
            ClipPrintText lastLoc.X, lastLoc.Y, IIf(Mid(pText, pSel.StartPos + 1, 1) = vbLf, " ", Mid(pText, pSel.StartPos + 1, 1)), SystemColorConstants.vbWindowText, False
        Else
            ClipLineDraw lastLoc.X, lastLoc.Y, lastLoc.X, (lastLoc.Y + TextHeight), SystemColorConstants.vbWindowBackground
        End If
    End If

    Static lastSel As RangeType
    If lastSel.StartPos <> pSel.StartPos Or lastSel.StopPos <> pSel.StopPos Then
        MakeCaretVisible newloc
    End If
    lastSel.StartPos = pSel.StartPos
    lastSel.StopPos = pSel.StopPos
    
    UserControl_Paint
    
    If ((cursorBlink Or ((newloc.X <> lastLoc.X) Or (newloc.Y <> lastLoc.Y))) And hasFocus) Then
        If insertMode Then
            'ClipPrintText newloc.X, newloc.Y, IIf(Convert(pText.Partial(pSel.StartPos, 1)) = vbLf, " ", Convert(pText.Partial(pSel.StartPos, 1))), SystemColorConstants.vbWindowText, True
            ClipPrintText newloc.X, newloc.Y, IIf(Mid(pText, pSel.StartPos + 1, 1) = vbLf, " ", Mid(pText, pSel.StartPos + 1, 1)), SystemColorConstants.vbWindowText, True
        Else
            ClipLineDraw newloc.X, newloc.Y, newloc.X, (newloc.Y + TextHeight), SystemColorConstants.vbWindowText
        End If
        
    ElseIf ((Not cursorBlink) Or (Not hasFocus)) Or ((newloc.X <> lastLoc.X) Or (newloc.Y <> lastLoc.Y)) Then
        If insertMode Then
            'ClipPrintText lastLoc.X, lastLoc.Y, IIf(Convert(pText.Partial(pSel.StartPos, 1)) = vbLf, " ", Convert(pText.Partial(pSel.StartPos, 1))), SystemColorConstants.vbWindowText, False
            ClipPrintText lastLoc.X, lastLoc.Y, IIf(Mid(pText, pSel.StartPos + 1, 1) = vbLf, " ", Mid(pText, pSel.StartPos + 1, 1)), SystemColorConstants.vbWindowText, False
        Else
            ClipLineDraw lastLoc.X, lastLoc.Y, lastLoc.X, (lastLoc.Y + TextHeight), SystemColorConstants.vbWindowBackground
        End If
    End If
    
    lastLoc.X = newloc.X
    lastLoc.Y = newloc.Y
    
    cursorBlink = Not cursorBlink
    
    If Not Timer1.Enabled Then
        Timer1.Enabled = cursorBlink
        If Not Timer1.Enabled Then
            Timer1_Timer
        End If
    End If
End Sub
Private Function MakeCaretVisible(ByRef loc As POINTAPI) As Boolean
    If pScrollToCaret And (Not ClippingWouldDraw(DrawableRect, Rect(loc.X, loc.Y, loc.X + TextWidth, loc.Y + TextHeight), True)) Then
        If loc.X < 0 Then
            pOffsetX = pOffsetX + ((0 - loc.X) + (UsercontrolWidth / 2))
            If ScrollBar2.Visible Then ScrollBar2.Value = -pOffsetX
            MakeCaretVisible = True
        ElseIf loc.X >= UsercontrolWidth Then
            pOffsetX = pOffsetX - ((loc.X - UsercontrolWidth) + (UsercontrolWidth / 2))
            If ScrollBar2.Visible Then ScrollBar2.Value = -pOffsetX
            MakeCaretVisible = True
        End If
        If loc.Y < 0 Then
            pOffsetY = pOffsetY + (((((0 - loc.Y) + (UsercontrolHeight / 2)) \ TextHeight)) * TextHeight)
            If ScrollBar1.Visible Then ScrollBar1.Value = -pOffsetY
            MakeCaretVisible = True
        ElseIf loc.Y >= (UsercontrolHeight - TextHeight) Then
            pOffsetY = pOffsetY - (((((loc.Y - UsercontrolHeight) + (UsercontrolHeight / 2)) \ TextHeight)) * TextHeight)
            If ScrollBar1.Visible Then ScrollBar1.Value = -pOffsetY
            MakeCaretVisible = True
        End If
        
        If MakeCaretVisible = True Then
            UserControl_Paint
        End If
    End If
End Function
Private Function ClippingWouldDraw(ByRef rct1 As Rect, ByRef rct2 As Rect, Optional ByVal FullDrawOnly As Boolean = False) As Boolean
    Dim rct As Rect
    Dim rct3 As Rect
    rct3 = Rect(rct1.Left, rct1.Top, rct1.Right, rct1.Bottom)
    ClippingWouldDraw = (IntersectRect(rct, rct3, rct2) <> 0)
    If Not ClippingWouldDraw Then
        If Not FullDrawOnly Then
            ClippingWouldDraw = (PtInRect(rct3, rct2.Left, rct2.Top) <> 0 Or PtInRect(rct3, rct2.Right, rct2.Top) <> 0 Or _
                        PtInRect(rct3, rct2.Left, rct2.Bottom) <> 0 Or PtInRect(rct3, rct2.Right, rct2.Bottom) <> 0)
            If Not ClippingWouldDraw Then
                ClippingWouldDraw = (rct3.Left = rct2.Left) Or (rct3.Top = rct2.Top) Or (rct3.Right = rct2.Right) Or (rct3.Bottom = rct2.Bottom)
            End If
        Else
            ClippingWouldDraw = (PtInRect(rct3, rct2.Left, rct2.Top) <> 0 And PtInRect(rct3, rct2.Right, rct2.Top) <> 0 And _
                        PtInRect(rct3, rct2.Left, rct2.Bottom) <> 0 And PtInRect(rct3, rct2.Right, rct2.Bottom) <> 0)
        End If
    ElseIf FullDrawOnly Then
        ClippingWouldDraw = False
    End If
End Function
Private Function ClipPrintText(ByVal X1 As Single, ByVal Y1 As Single, ByVal StrText As String, ByVal Color As Long, Optional ByVal BoxFill As Boolean = False) As Boolean

    If BoxFill Then
        If ClipLineDraw(X1, Y1, (UserControl.TextWidth(StrText) + X1), (UserControl.TextHeight(StrText) + Y1), Color, True) Then
            UserControl.Forecolor = pBackColor
            UserControl.CurrentX = X1
            UserControl.CurrentY = Y1
            UserControl.Print StrText
            ClipPrintText = True
        End If
    ElseIf ClippingWouldDraw(DrawableRect, Rect(X1, Y1, (UserControl.TextWidth(StrText) + X1), (UserControl.TextHeight(StrText) + Y1))) Then
        UserControl.Forecolor = Color
        UserControl.CurrentX = X1
        UserControl.CurrentY = Y1
        UserControl.Print StrText
        ClipPrintText = True
    End If
    
End Function

Private Function ClipPrintTextBlock(ByVal X1 As Single, ByVal Y1 As Single, ByVal StrText As String, ByVal Color As Long, Optional ByVal BoxFill As Boolean = False) As POINTAPI
    With ClipPrintTextBlock
        .X = X1
        .Y = Y1
        Dim outLine As String
        If StrText <> "" Then
            UserControl.CurrentX = .X
            UserControl.CurrentY = .Y
            Do While InStr(StrText, vbLf) > 0
                outLine = RemoveNextArg(StrText, vbLf)
                ClipPrintText .X, .Y, outLine, Color, BoxFill
                .X = pOffsetX
                .Y = .Y + UserControl.TextHeight(outLine)
            Loop
            If StrText <> "" Then
                ClipPrintText .X, .Y, StrText, Color, BoxFill
                .X = .X + UserControl.TextWidth(StrText)
            End If
        End If
    End With
End Function

Private Function ClipLineDraw(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ByVal Color As Long, Optional ByVal BoxFill As Boolean = False) As Boolean
    If ClippingWouldDraw(DrawableRect, Rect(X1, Y1, X2, Y2)) Then
        If BoxFill Then
            UserControl.Line (X1, Y1)-(X2, Y2), Color, BF
        Else
            UserControl.Line (X1, Y1)-(X2, Y2), Color
        End If
        ClipLineDraw = True
    End If
End Function

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_GotFocus()
    hasFocus = True
End Sub

Private Sub UserControl_Initialize()
    
    SystemParametersInfo SPI_GETKEYBOARDSPEED, 0, keySpeed, 0
    Timer1.Interval = keySpeed * 10

    Hook
End Sub

Private Sub Hook()
    If Not IsRunningMode Then
        SubClassed2.Add Me, "H" & UserControl.hWnd
        OldProc = GetWindowLong(UserControl.hWnd, GWL_WNDPROC)
        SetWindowLong UserControl.hWnd, GWL_WNDPROC, AddressOf TextBoxProc
    End If
End Sub

Private Sub Unhook()
    If OldProc <> 0 Then
        SetWindowLong UserControl.hWnd, GWL_WNDPROC, OldProc
        SubClassed2.Remove "H" & UserControl.hWnd
        OldProc = 0
    End If
End Sub

Private Sub UserControl_InitProperties()
    pForeColor = SystemColorConstants.vbWindowText
    pBackColor = SystemColorConstants.vbWindowBackground
    UserControl.Backcolor = SystemColorConstants.vbWindowBackground
    pScrollToCaret = True
    pHideSelection = True
    UserControl.Font.Name = "Lucida Console"
    pMultiLine = False
    pScrollBars = vbScrollBars.None
End Sub

Public Function LineOffset(ByVal LineIndex As Long) As Long ' _
Returns the offset amount of characters upto a line, specified by the zero based LineIndex. Example, LineOffset(0)=0.
Attribute LineOffset.VB_Description = "Returns the offset amount of characters upto a line, specified by the zero based LineIndex. Example, LineOffset(0)=0."
    Dim lines() As String
    lines = Split(pText, vbLf)
    If LineIndex >= LBound(lines) And LineIndex <= UBound(lines) And LineIndex - 1 >= 0 Then
        Dim cnt As Long
        For cnt = 0 To LineIndex - 1
            LineOffset = LineOffset + Len(lines(cnt)) + 1
        Next
    End If
End Function

Public Function LineLength(ByVal LineIndex As Long) As Long ' _
Returns the length of characters with-in a line, specifiied by the zero based LineIndex.
Attribute LineLength.VB_Description = "Returns the length of characters with-in a line, specifiied by the zero based LineIndex."
    Dim lines() As String
    lines = Split(pText, vbLf)
    If LineIndex >= LBound(lines) And LineIndex <= UBound(lines) Then
        LineLength = (Len(lines(LineIndex)) + 1)
    End If
End Function

Public Function LineText(ByVal LineIndex As Long) As String ' _
Returns the text with-in a line, specified by the zero based LineIndex.
Attribute LineText.VB_Description = "Returns the text with-in a line, specified by the zero based LineIndex."
    Dim lPos As Long
    lPos = LineLength(LineIndex)
    If lPos >= 0 Then
        LineIndex = LineOffset(LineIndex)
        If LineIndex > 0 Then
            LineText = Mid(pText, LineIndex, lPos)
        ElseIf lPos > 0 Then
            LineText = Mid(pText, 1, lPos)
        End If
    End If
End Function

Public Function LineIndex(Optional ByVal CharIndex As Long = -1) As Long ' _
Returns the zero based index line number of which the given zero based character index falls upon with-in Text.
Attribute LineIndex.VB_Description = "Returns the zero based index line number of which the given zero based character index falls upon with-in Text."
    If CharIndex = -1 Then CharIndex = pSel.StartPos
    If CharIndex > 0 Then
        LineIndex = CountWord(Left(pText, CharIndex), vbLf)
    Else
        LineIndex = 0
    End If
End Function

Public Function LineCount() As Long ' _
Returns the numerical count of how many lines, delimited by line feeds, that exists with-in Text.
Attribute LineCount.VB_Description = "Returns the numerical count of how many lines, delimited by line feeds, that exists with-in Text."
    Dim lines() As String
    lines = Split(pText, vbLf)
    LineCount = ((UBound(lines) - LBound(lines)) + 1)
End Function

Private Function CaretLocation(Optional ByVal AtCharPos As Long = -1) As POINTAPI
    Dim part As String
    If pSel.StartPos <= 0 Then pSel.StartPos = 0
    If AtCharPos = -1 Then AtCharPos = pSel.StartPos
    part = Left(pText, AtCharPos)
    If InStr(part, vbLf) > 0 Then
        CaretLocation.Y = (TextHeight * CountWord(part, vbLf)) + pOffsetY
        part = StrReverse(RemoveNextArg(StrReverse(part), vbLf))
    Else
        CaretLocation.Y = pOffsetY
    End If
    CaretLocation.X = UserControl.TextWidth(part) + pOffsetX
End Function

Public Function CaretFromPoint(ByVal X As Single, ByVal Y As Single) As Long ' _
Retruns the zero based character index position with-in Text based upon given pixel corrdinates, X and Y.
Attribute CaretFromPoint.VB_Description = "Retruns the zero based character index position with-in Text based upon given pixel corrdinates, X and Y."
    Dim lText As String
    X = (X - pOffsetX)
    Y = ((Y - pOffsetY) \ TextHeight)
    If Y < LineCount() Then
        If X >= TextWidth(LineText(Y)) Then
            CaretFromPoint = LineOffset(Y) + Len(Replace(LineText(Y), vbLf, ""))
        Else
            lText = Replace(LineText(Y), vbLf, "")
            Do While TextWidth(lText) - (TextWidth(Right(lText, 1)) / 2) >= X
                lText = Left(lText, Len(lText) - 1)
                If lText = "" Then Exit Do
            Loop
            CaretFromPoint = LineOffset(Y) + Len(lText)
        End If
    Else
        CaretFromPoint = Len(pText)
    End If
End Function

Public Function LineFirstVisible() As Long ' _
Returns the zero based line index number of the first visible line on the screen.
Attribute LineFirstVisible.VB_Description = "Returns the zero based line index number of the first visible line on the screen."
    LineFirstVisible = (-pOffsetY \ TextHeight)
End Function

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    If KeyCode <> 0 Then
        Select Case KeyCode
            Case 13 'enter
                If pSel.StartPos = pSel.StopPos Then
                    pText = Left(pText, pSel.StartPos) & vbLf & Mid(pText, pSel.StartPos + 1)
                    pSel.StartPos = pSel.StartPos + 1
                Else
                    If pSel.StartPos > pSel.StopPos Then
                        Swap pSel.StartPos, pSel.StopPos
                    End If
                    pText = Left(pText, pSel.StartPos) & vbLf & Mid(pText, pSel.StopPos + 1)

                    pSel.StartPos = pSel.StartPos + 1
                End If
                pSel.StopPos = pSel.StartPos
                InvalidateCursor
                RaiseChangeEvent
            Case 45 'insert
                If Shift = 0 Then insertMode = Not insertMode
            Case 46, 8 'delete 'backspace
                If pSel.StartPos = pSel.StopPos Then
                    If KeyCode = 8 And pSel.StartPos > 0 Then
                        pText = Left(pText, pSel.StartPos - 1) & Mid(pText, pSel.StartPos + 1)

                        pSel.StartPos = pSel.StartPos - 1
                    ElseIf KeyCode = 46 Then
                        pText = Left(pText, pSel.StartPos) & Mid(pText, pSel.StartPos + 2)
                        
                        pSel.StartPos = pSel.StartPos
                    End If
                Else
                    If pSel.StartPos > pSel.StopPos Then
                        Swap pSel.StartPos, pSel.StopPos
                    End If
                        
                    pText = Left(pText, pSel.StartPos) & Mid(pText, pSel.StopPos + 1)
                    
                End If
                pSel.StopPos = pSel.StartPos
                InvalidateCursor
                RaiseChangeEvent
            Case 36 'home
                pSel.StartPos = LineOffset(LineIndex(pSel.StartPos))
                If Shift = 0 Then pSel.StopPos = pSel.StartPos
                'InvalidateCursor
            Case 35 'end
                pSel.StartPos = LineOffset(LineIndex(pSel.StartPos)) + LineLength(LineIndex(pSel.StartPos)) - 1
                If Shift = 0 Then pSel.StopPos = pSel.StartPos
                'InvalidateCursor
            Case 38 'up
                KeyArrowUp Shift
                InvalidateCursor
            Case 40 'down
                KeyArrowDown Shift
                InvalidateCursor
            Case 39 'right
                KeyArrowRight Shift
                InvalidateCursor
            Case 37 'left
                KeyArrowLeft Shift
                InvalidateCursor
            Case 33 'pgup
                KeyPageUp Shift
                InvalidateCursor
            Case 34 'pgdn
                KeyPageDown Shift
                InvalidateCursor
        End Select
        If Shift = 2 Then
            Select Case KeyCode
                Case 90 'ctrl+z
                Case 88 'ctrl+x
                Case 67 'ctrl+c
                Case 86 'ctrl+v
                Case 65 'ctrl+a
                
            End Select
        End If
    End If
End Sub
Friend Sub KeyPageUp(ByRef Shift As Integer)
    pSel.StartPos = LineCharPos(LineIndex - (UsercontrolHeight \ TextHeight))
    If Shift = 0 Then pSel.StopPos = pSel.StartPos
End Sub
Friend Sub KeyPageDown(ByRef Shift As Integer)
    pSel.StartPos = LineOffset(LineIndex + (UsercontrolHeight \ TextHeight))
    If Shift = 0 Then pSel.StopPos = pSel.StartPos
End Sub
Friend Sub KeyArrowUp(ByRef Shift As Integer)
    Dim caret As POINTAPI
    caret = CaretLocation
    caret.Y = caret.Y - TextHeight
    pSel.StartPos = CaretFromPoint(caret.X, caret.Y)
    If Shift = 0 Then pSel.StopPos = pSel.StartPos
End Sub
Friend Sub KeyArrowDown(ByRef Shift As Integer)
    Dim caret As POINTAPI
    caret = CaretLocation
    caret.Y = caret.Y + TextHeight
    pSel.StartPos = CaretFromPoint(caret.X, caret.Y)
    If Shift = 0 Then pSel.StopPos = pSel.StartPos
End Sub
Friend Sub KeyArrowLeft(ByRef Shift As Integer)
    If pSel.StartPos > 0 Then pSel.StartPos = pSel.StartPos - 1
    If Shift = 0 Then pSel.StopPos = pSel.StartPos
End Sub
Friend Sub KeyArrowRight(ByRef Shift As Integer)
    If pSel.StartPos < Len(pText) Then pSel.StartPos = pSel.StartPos + 1
    If Shift = 0 Then pSel.StopPos = pSel.StartPos
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    If KeyAscii <> 0 Then
        If ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or IsNumeric(Chr(KeyAscii))) Or _
            (Chr(KeyAscii) = "`" Or Chr(KeyAscii) = "~" Or Chr(KeyAscii) = "!" Or Chr(KeyAscii) = "@" Or Chr(KeyAscii) = "#" Or Chr(KeyAscii) = "$" Or _
            Chr(KeyAscii) = "%" Or Chr(KeyAscii) = "^" Or Chr(KeyAscii) = "&" Or Chr(KeyAscii) = "*" Or Chr(KeyAscii) = "(" Or Chr(KeyAscii) = ")" Or _
            Chr(KeyAscii) = "_" Or Chr(KeyAscii) = "-" Or Chr(KeyAscii) = "+" Or Chr(KeyAscii) = "=" Or Chr(KeyAscii) = "\" Or Chr(KeyAscii) = "|" Or _
            Chr(KeyAscii) = "[" Or Chr(KeyAscii) = "{" Or Chr(KeyAscii) = "]" Or Chr(KeyAscii) = "}" Or Chr(KeyAscii) = ":" Or Chr(KeyAscii) = ";" Or _
            Chr(KeyAscii) = """" Or Chr(KeyAscii) = "'" Or Chr(KeyAscii) = "<" Or Chr(KeyAscii) = "," Or Chr(KeyAscii) = ">" Or Chr(KeyAscii) = "." Or _
            Chr(KeyAscii) = "?" Or Chr(KeyAscii) = "/") Then

            If insertMode Then
                pText = Left(pText, pSel.StartPos) & Chr(KeyAscii) & Mid(pText, pSel.StartPos + 2)
            ElseIf pSel.StartPos < Len(pText) Then
                pText = Left(pText, pSel.StartPos) & Chr(KeyAscii) & Mid(pText, pSel.StartPos + 1)
            Else
                pText = pText & Chr(KeyAscii)
            End If
            
            pSel.StartPos = pSel.StartPos + 1
            pSel.StopPos = pSel.StartPos
            InvalidateCursor
            RaiseChangeEvent

        End If
    End If
End Sub

Private Sub RaiseChangeEvent()
    SetScrollBars
    RaiseEvent Change
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_LostFocus()
    hasFocus = False
    If pHideSelection Then UserControl_Paint
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button = 1 Then
    
        Dim lPos As Long
        pSel.StartPos = CaretFromPoint(X, Y)
        If Shift = 0 Then pSel.StopPos = pSel.StartPos
        
        InvalidateCursor
    End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)

    If Button = 1 Then
        Dim lPos As Long
        lPos = CaretFromPoint(X, Y)

        If lPos < pSel.StopPos And (dragStart = 1 Or dragStart = 0) Then
            pSel.StopPos = lPos
            dragStart = 1
        ElseIf (dragStart = 2 Or dragStart = 0) Then
            pSel.StartPos = lPos
            dragStart = 2
        End If
        
        Dim pt As POINTAPI
        Dim rct As Rect
        GetCursorPos pt
        GetWindowRect UserControl.hWnd, rct
        
        Dim loc As POINTAPI
        loc = CaretLocation

        Dim newloc As POINTAPI
        newloc.X = loc.X
        newloc.Y = loc.Y

        If pt.X < rct.Left Then
            If ((rct.Left - pt.X) < (UsercontrolWidth / 2)) Then 'slow
                newloc.X = newloc.X - TextWidth
            Else
                newloc.X = newloc.X - (TextWidth * 4)
            End If
        
        ElseIf pt.X > rct.Right Then
            If ((pt.X - rct.Right) < (UsercontrolWidth / 2)) Then 'slow
                newloc.X = newloc.X + TextWidth
            Else
                newloc.X = newloc.X + (TextWidth * 4)
            End If
        End If
        
        If pt.Y < rct.Top Then
            If ((rct.Top - pt.Y) < (UsercontrolWidth / 2)) Then 'slow
                newloc.Y = newloc.Y - TextHeight
            Else
                newloc.Y = newloc.Y - (TextHeight * 4)
            End If
        ElseIf pt.Y > rct.Bottom Then
            If ((pt.Y - rct.Bottom) < (UsercontrolWidth / 2)) Then 'slow
                newloc.Y = newloc.Y + TextHeight
            Else
                newloc.Y = newloc.Y + (TextHeight * 4)
            End If
        End If

        If loc.X <> newloc.X Or loc.Y <> newloc.Y Then
            MakeCaretVisible newloc
        End If

    Else
        dragStart = 0
    End If
    
End Sub

Friend Sub InvalidateCursor()
    Timer1.Enabled = False
    Timer1_Timer
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    dragStart = 0
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_Paint()
    CanvasValidate
    
    Dim tmpStr As String
    Dim cur As POINTAPI
    Dim tmpSel As RangeType
    cur.X = pOffsetX
    cur.Y = pOffsetY
    
    tmpStr = pText

    If pSel.StartPos <> pSel.StopPos Then
        If pSel.StartPos > pSel.StopPos Then
            tmpSel.StartPos = pSel.StopPos
            tmpSel.StopPos = pSel.StartPos
        Else
            tmpSel.StartPos = pSel.StartPos
            tmpSel.StopPos = pSel.StopPos
        End If
    End If

    UserControl.Cls
    UserControl.Line (0, 0)-(UsercontrolWidth, UsercontrolHeight), SystemColorConstants.vbWindowBackground, BF
    
    UserControl.CurrentX = cur.X
    UserControl.CurrentY = cur.Y
    
    If ((pSel.StartPos <> pSel.StopPos) And (hasFocus Xor ((Not hasFocus) And (Not pHideSelection)))) Then
        cur = ClipPrintTextBlock(cur.X, cur.Y, Left(tmpStr, tmpSel.StartPos), pForeColor, False)
        cur = ClipPrintTextBlock(cur.X, cur.Y, Mid(tmpStr, tmpSel.StartPos + 1, (tmpSel.StopPos - tmpSel.StartPos)), SystemColorConstants.vbHighlight, True)
        tmpStr = Mid(tmpStr, tmpSel.StopPos + 1)
        If Len(tmpStr) <> 0 Then
            cur = ClipPrintTextBlock(cur.X, cur.Y, tmpStr, pForeColor, False)
        End If

    Else
        ClipPrintTextBlock cur.X, cur.Y, tmpStr, pForeColor, False
    End If

    If ScrollBar1.Visible And ScrollBar2.Visible Then
        DrawFrameControl hdc, Rect(((UsercontrolWidth / Screen.TwipsPerPixelX)), ((UsercontrolHeight / Screen.TwipsPerPixelY)), UserControl.Width / Screen.TwipsPerPixelX, UserControl.Height / Screen.TwipsPerPixelY), DFC_SCROLL, DFCS_SCROLLSIZEGRIP
    End If
    
    Static lastSel As RangeType
    If lastSel.StartPos <> pSel.StartPos Or lastSel.StopPos <> pSel.StopPos Then
        RaiseEvent SelChange
    End If
    lastSel.StartPos = pSel.StartPos
    lastSel.StopPos = pSel.StopPos
    
    RaiseEvent Paint


End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Backcolor = PropBag.ReadProperty("Backcolor", SystemColorConstants.vbWindowBackground)
    Forecolor = PropBag.ReadProperty("Forecolor", SystemColorConstants.vbWindowText)
    ScrollToCaret = PropBag.ReadProperty("ScrollToCaret", True)
    HideSelection = PropBag.ReadProperty("HideSelection", True)
    UserControl.Font.Name = PropBag.ReadProperty("Fontname", "Lucida Console")
    Text = PropBag.ReadProperty("Text", "")
    MultipleLines = PropBag.ReadProperty("MultipleLines", False)
    ScrollBars = PropBag.ReadProperty("ScrollBars", vbScrollBars.None)
End Sub

Private Sub UserControl_Resize()
    CanvasValidate
    SetScrollBars
    RaiseEvent Resize
End Sub

Private Sub ScrollBar1_Change()
    If ScrollBar1.Visible Then OffsetY = -ScrollBar1.Value
End Sub

Private Sub ScrollBar1_Scroll()
    If ScrollBar1.Visible Then OffsetY = -ScrollBar1.Value
End Sub

Private Sub ScrollBar2_Change()
     If ScrollBar2.Visible Then OffsetX = -ScrollBar2.Value
End Sub

Private Sub ScrollBar2_Scroll()
    If ScrollBar2.Visible Then OffsetX = -ScrollBar2.Value
End Sub

Private Sub SetScrollBars()

    Dim i As Integer
    For i = 0 To 1
        If ((CanvasHeight > UsercontrolHeight And UsercontrolHeight > ScrollBar2.Height) And (pScrollBars = vbScrollBars.Auto)) Or ((pScrollBars = vbScrollBars.Both) Or (pScrollBars = vbScrollBars.Vertical)) Then
            If Not ScrollBar1.Visible Then ScrollBar1.Visible = True
        ElseIf ScrollBar1.Visible Then
            ScrollBar1.Visible = False
        End If
        If ScrollBar1.Visible And CanvasHeight < UsercontrolHeight Then
            If ScrollBar1.Enabled Then ScrollBar1.Enabled = False
        ElseIf ScrollBar1.Visible And CanvasHeight >= UsercontrolHeight Then
            If Not ScrollBar1.Enabled Then ScrollBar1.Enabled = True
        End If
        
        If ((CanvasWidth > UsercontrolWidth And UsercontrolWidth > ScrollBar1.Width) And (pScrollBars = vbScrollBars.Auto)) Or ((pScrollBars = vbScrollBars.Both) Or (pScrollBars = vbScrollBars.Horizontal)) Then
            If Not ScrollBar2.Visible Then ScrollBar2.Visible = True
        ElseIf ScrollBar2.Visible Then
            ScrollBar2.Visible = False
        End If
        If ScrollBar2.Visible And CanvasWidth < UsercontrolWidth Then
            If ScrollBar2.Enabled Then ScrollBar2.Enabled = False
        ElseIf ScrollBar2.Visible And CanvasWidth >= UsercontrolWidth Then
            If Not ScrollBar2.Enabled Then ScrollBar2.Enabled = True
        End If
    Next
    
    If ScrollBar1.Visible Then
        ScrollBar1.Max = (CanvasHeight - UsercontrolHeight)
        ScrollBar1.SmallChange = TextHeight
        ScrollBar1.LargeChange = ScrollBar1.SmallChange * 4
        If ScrollBar1.Top <> 0 Then ScrollBar1.Top = 0
        If ScrollBar1.Width <> Screen.TwipsPerPixelX ^ 2 Then ScrollBar1.Width = Screen.TwipsPerPixelX ^ 2
        If ScrollBar1.Left <> UsercontrolWidth Then ScrollBar1.Left = UsercontrolWidth
        If ScrollBar1.Height <> UsercontrolHeight Then ScrollBar1.Height = UsercontrolHeight
    End If
    If ScrollBar2.Visible Then
        ScrollBar2.Max = (CanvasWidth - UsercontrolWidth)
        ScrollBar2.SmallChange = TextWidth
        ScrollBar2.LargeChange = ScrollBar2.SmallChange * 4
        If ScrollBar2.Left <> 0 Then ScrollBar2.Left = 0
        If ScrollBar2.Height <> Screen.TwipsPerPixelY ^ 2 Then ScrollBar2.Height = Screen.TwipsPerPixelY ^ 2
        If ScrollBar2.Top <> UsercontrolHeight Then ScrollBar2.Top = UsercontrolHeight
        If ScrollBar2.Width <> UsercontrolWidth Then ScrollBar2.Width = UsercontrolWidth
    End If

End Sub

Private Sub UserControl_Show()
    CanvasValidate
    SetScrollBars
    UserControl_Paint
End Sub

Private Sub UserControl_Terminate()
    Unhook

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Backcolor", pBackColor, SystemColorConstants.vbWindowBackground
    PropBag.WriteProperty "Forecolor", pForeColor, SystemColorConstants.vbWindowText
    PropBag.WriteProperty "ScrollToCaret", pScrollToCaret, True
    PropBag.WriteProperty "HideSelection", pHideSelection, True
    PropBag.WriteProperty "MultipleLines", pMultiLine, False
    PropBag.WriteProperty "Fontname", UserControl.Font.Name, "Lucida Console"
    PropBag.WriteProperty "Text", pText, ""
    PropBag.WriteProperty "ScrollBars", pScrollBars, vbScrollBars.None
End Sub

Private Function RemoveNextArg(ByRef TheParams As Variant, ByVal TheSeperator As String) As String
    If InStr(1, TheParams, TheSeperator) > 0 Then
        RemoveNextArg = Left(TheParams, InStr(1, TheParams, TheSeperator) - 1)
        TheParams = Mid(TheParams, InStr(1, TheParams, TheSeperator) + Len(TheSeperator))
    Else
        RemoveNextArg = TheParams
        TheParams = ""
    End If
End Function

Public Function CountWord(ByVal Text As String, ByVal Word As String, Optional ByVal Exact As Boolean = True) As Long ' _
Counts how many times Word appears in Text, optionally specifying the Exact parameter to false for case insensitive matching.
    Dim cnt As Long
    cnt = UBound(Split(Text, Word, , IIf(Exact, vbBinaryCompare, vbTextCompare)))
    If cnt > 0 Then CountWord = cnt
End Function

