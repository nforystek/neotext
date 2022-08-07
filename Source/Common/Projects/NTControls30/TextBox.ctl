VERSION 5.00
Begin VB.UserControl TextBox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3780
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
      Height          =   315
      Left            =   165
      Top             =   1185
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   450
      AutoRedraw      =   0   'False
   End
   Begin NTControls30.ScrollBar ScrollBar1 
      Height          =   1290
      Left            =   3090
      Top             =   75
      Visible         =   0   'False
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   2275
      Orientation     =   0
      AutoRedraw      =   0   'False
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   75
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "De&lete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuClear 
         Caption         =   "C&lear All"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuDash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIndent 
         Caption         =   "Indent &Tab"
      End
      Begin VB.Menu mnuUnindent 
         Caption         =   "Unindent Ta&b"
      End
   End
End
Attribute VB_Name = "TextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IControl


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

Private pEnabled As Boolean
Private pLocked As Boolean
Private pHideSelection As Boolean
Private pScrollToCaret As Boolean
Private pMultiLine As Boolean
Private pLineNumbers As Boolean
'Private pForecolors As Collection
'Private pBackcolors As Collection
Private pForecolor As OLE_COLOR
Private pBackcolor As OLE_COLOR
Private pScrollBars As vbScrollBars
Private pTabSpace As String

Private pOldProc As Long

Private pText As NTNodes10.Strands
Private pScroll1Buffer As Backbuffer
Private pScroll2Buffer As Backbuffer
Private pBackBuffer As Backbuffer

Private xCancel As Long 'hold the cancel stack, for every set true, must also occur the set false

Private xUndoStack As Long
Private xUndoDirty As Boolean
Private xLatency As Single
Private xUndoDelay As Single

Private xUndoText() As NTNodes10.Strands
Private xUndoStart() As Long
Private xUndoLength() As Long
Private xUndoStage As Long
Private xUndoBuffer As Long

Private pLastSel As RangeType
Private pSel As RangeType 'where the current selection is held at all states or set

Public Property Get UndoDelay() As Long
    UndoDelay = xUndoDelay
End Property
Public Property Let UndoDelay(ByVal RHS As Long)
    If RHS >= 0 And RHS <> xUndoDelay Then
        xUndoDelay = RHS
    End If
End Property

Public Property Get UndoStack() As Long
    UndoStack = xUndoStack
End Property
Public Property Let UndoStack(ByVal RHS As Long)
    If RHS > -2 And xUndoStack <> RHS Then
        xUndoStack = RHS
        ResetUndoRedo
    End If
End Property

Public Property Get UndoDirty() As Boolean
    UndoDirty = xUndoDirty
End Property
Public Property Let UndoDirty(ByVal RHS As Boolean)
    If xUndoDirty <> RHS Then
        xUndoDirty = RHS
        ResetUndoRedo
    End If
End Property

Private Sub ResetUndoRedo()
    If ArraySize(xUndoText) > 0 Then
        Dim cnt As Long
        For cnt = LBound(xUndoText) To UBound(xUndoText)
            Set xUndoText(cnt) = Nothing
        Next
    End If
        
    ReDim xUndoText(0) As NTNodes10.Strands
    ReDim xUndoStart(0) As Long
    ReDim xUndoLength(0) As Long
    xUndoStage = 0
    xUndoDirty = False
    xUndoBuffer = xUndoStack
    Set xUndoText(0) = New NTNodes10.Strands
    If Me.Text.Length > 0 Then
        xUndoText(0).Concat pText.Partial
    End If
    xUndoStart(0) = Me.SelStart
    xUndoLength(0) = Me.SelLength
    
End Sub

Private Sub AddUndo()

    If xUndoDirty Then
    
       If (xUndoStack <> 0) Then
        
            Dim cnt As Long
            
            If ((UBound(xUndoText) < xUndoBuffer) Or (xUndoStack = -1)) Or (Not (xUndoStage = UBound(xUndoText))) Then
            
                If ArraySize(xUndoText) >= xUndoStage + 1 Then
                    For cnt = (xUndoStage + 1) To UBound(xUndoText)
                        Set xUndoText(cnt) = Nothing
                    Next
                End If
                
                ReDim Preserve xUndoText(xUndoStage + 1) As NTNodes10.Strands
                Set xUndoText(xUndoStage + 1) = New NTNodes10.Strands
                ReDim Preserve xUndoStart(xUndoStage + 1) As Long
                ReDim Preserve xUndoLength(xUndoStage + 1) As Long
                xUndoStage = xUndoStage + 1
            ElseIf (UBound(xUndoText) = xUndoBuffer) Or (xUndoStage = UBound(xUndoText)) Then
                
                Set xUndoText(LBound(xUndoText)) = Nothing
                For cnt = (LBound(xUndoText) + 1) To UBound(xUndoText)
                    xUndoText(cnt - 1) = xUndoText(cnt)
                    xUndoStart(cnt - 1) = xUndoStart(cnt)
                    xUndoLength(cnt - 1) = xUndoLength(cnt)
                Next
                Set xUndoText(xUndoStage) = New NTNodes10.Strands
        
            End If
    
            If Me.Text.Length > 0 Then
                xUndoText(UBound(xUndoText)).Concat pText.Partial
            End If
            xUndoStart(UBound(xUndoStart)) = Me.SelStart
            xUndoLength(UBound(xUndoLength)) = Me.SelLength
            
        ElseIf xUndoDirty Then
            ResetUndoRedo
        End If
        
        xUndoDirty = False
        
    End If
End Sub

Public Function CanUndo() As Boolean
    CanUndo = ((UBound(xUndoText) > 0) And (xUndoStage > 0)) And (Not Locked) And (xUndoDelay > -1)
End Function
Public Function CanRedo() As Boolean
    CanRedo = ((xUndoStage < UBound(xUndoText)) And (UBound(xUndoText) > 0)) And (Not Locked) And (xUndoDelay > -1)
End Function

Public Sub Undo()
    If xUndoDelay > 0 And xUndoDirty Then AddUndo
    
    If CanUndo Then
                
        Cancel = True
        xUndoStage = xUndoStage - 1
        Set Me.Text = xUndoText(xUndoStage)
        pSel.StartPos = xUndoStart(xUndoStage)
        pSel.StopPos = pSel.StartPos + xUndoLength(xUndoStage)
        
        Cancel = False
        RaiseEvent Change
        InvalidateCursor
        
    End If
End Sub

Public Sub Redo()
    If xUndoDelay > 0 And xUndoDirty Then AddUndo
    
    If CanRedo Then
    
        Cancel = True
        xUndoStage = xUndoStage + 1
        Set Me.Text = xUndoText(xUndoStage)
        pSel.StartPos = xUndoStart(xUndoStage)
        pSel.StopPos = pSel.StartPos + xUndoLength(xUndoStage)
        
        Cancel = False
        RaiseEvent Change
        InvalidateCursor
    End If
End Sub

Friend Property Get Cancel() As Boolean
    Cancel = CBool(xCancel > 0)
End Property
Friend Property Let Cancel(ByVal newVal As Boolean)
    If newVal Then
        xCancel = xCancel + 1
    Else
        xCancel = xCancel - 1
    End If
End Property

Public Property Get TabSpace() As String
    TabSpace = pTabSpace
End Property
Public Property Let TabSpace(ByVal RHS As String)
    If Replace(RHS, " ", "") = "" And Len(RHS) > 0 Then
        pTabSpace = RHS
    End If
End Property
Public Sub SelectAll()
    If Enabled Then
        pSel.StartPos = 0
        pSel.StopPos = Length
        InvalidateCursor
    End If
End Sub
Public Sub Delete()
    If Not Locked And Enabled Then
        UserControl_KeyDown 46, 0
    End If
End Sub

Private Sub mnuClear_Click()
    ClearAll
End Sub

Private Sub mnuCopy_Click()
    Copy
End Sub

Private Sub mnuCut_Click()
    Cut
End Sub

Private Sub mnuDelete_Click()
    Delete
End Sub

Private Sub mnuPaste_Click()
    Paste
End Sub

Private Sub mnuSelectAll_Click()
    SelectAll
End Sub


Private Sub mnuIndent_Click()
    Indenting SelStart, SelLength, Chr(9), True
End Sub

Private Sub mnuUnindent_Click()
    Indenting SelStart, SelLength, Chr(9) & Chr(8), True
End Sub


Private Sub mnuEdit_Click()
    On Error Resume Next
    RefreshEditMenu
    If Err Then Err.Clear
    On Error GoTo 0
End Sub

Public Sub Cut()

    If Not Locked And Enabled Then

        If SelLength > 0 Then
            Cancel = True
            
            Clipboard.SetText SelText
            SelText = ""
            
            Cancel = False
            RaiseEventChange
        End If

    End If

End Sub
Public Sub Copy()
    If Enabled Then
        If SelLength > 0 Then
            Clipboard.SetText SelText
        End If
    End If
End Sub
Public Sub Paste()
    If Not Locked And Enabled Then

        If Clipboard.GetFormat(ClipBoardConstants.vbCFText) Then
            Cancel = True
            
            SelText = Clipboard.GetText(ClipBoardConstants.vbCFText)
            If pSel.StartPos < pSel.StopPos Then
                pSel.StartPos = pSel.StopPos
            Else
                pSel.StopPos = pSel.StartPos
            End If
        
            Cancel = False
            RaiseEventChange
        End If

    End If
End Sub
Public Sub ClearAll()
    If Not Locked And Enabled Then
        Cancel = True
        
        pSel.StartPos = 0
        pSel.StopPos = 0
        pText.Reset
        
        Cancel = False
        RaiseEventChange
              
    End If
End Sub

Public Sub Test()
'    Dim pt As POINTAPI
'    GetCursorPos pt
'    pt.X = 480
'    pt.Y = 780
'    Debug.Print SendMessageStruct(hWnd, EM_SETSCROLLPOS, 0&, pt)
'    Debug.Print pt.X; pt.Y
    
    'SelLength = 4
    SetLineText UserControl.hWnd, 0, "Option Compare Binary"
    
    'Debug.Print GetTextRange(hWnd, textbox1.SelStart, textbox1.SelStart + textbox1.SelLength)
    
    
    
End Sub

Public Property Get LineNumbers() As Boolean
    LineNumbers = pLineNumbers
End Property
Public Property Let LineNumbers(ByVal RHS As Boolean)
    If pLineNumbers <> RHS Then
        pLineNumbers = RHS
        UserControl_Paint
        PaintBuffer
    End If
End Property
Public Property Get Enabled() As Boolean ' _
Gets whether or not the text control accepts any interaction at all.
Attribute Enabled.VB_Description = "Gets whether or not the text control accepts any interaction at all."
    Enabled = pEnabled
End Property
Public Property Let Enabled(ByVal RHS As Boolean) ' _
Sets whether or not the text control accepts any interaction at all
Attribute Enabled.VB_Description = "Sets whether or not the text control accepts any interaction at all"
    pEnabled = RHS
    SetScrollBars
End Property

Public Property Get Locked() As Boolean ' _
Gets whether or not the text contents may be altered by user input, or is read-only locked.
Attribute Locked.VB_Description = "Gets whether or not the text contents may be altered by user input, or is read-only locked."
    Locked = pLocked
End Property
Public Property Let Locked(ByVal RHS As Boolean) ' _
Sets whether or not the text contents may be altered by user input, or is read-only locked.
Attribute Locked.VB_Description = "Sets whether or not the text contents may be altered by user input, or is read-only locked."
    pLocked = RHS
End Property
Public Function Length() As Long ' _
Returns the count of the number of characters in pText
Attribute Length.VB_Description = "Returns the count of the number of characters in pText"
    Length = pText.Length
End Function
Public Property Get AutoRedraw() As Boolean ' _
Gets whether or not the scroll bar automatically redraws itself.
Attribute AutoRedraw.VB_Description = "Gets whether or not the scroll bar automatically redraws itself."
    AutoRedraw = UserControl.AutoRedraw
End Property
Public Property Let AutoRedraw(ByVal RHS As Boolean) ' _
Sets whether or not the scroll bar automaticall redraws itself.
Attribute AutoRedraw.VB_Description = "Sets whether or not the scroll bar automaticall redraws itself."
    If UserControl.AutoRedraw <> RHS Then
        UserControl.AutoRedraw = RHS
        UserControl_Paint
        PaintBuffer
    End If
End Property

Public Property Get Backbuffer() As Backbuffer
    Set Backbuffer = pBackBuffer
End Property
Public Property Set Backbuffer(ByRef RHS As Backbuffer)
    Set pBackBuffer = RHS
End Property

Friend Property Get VScroll() As ScrollBar
    Set VScroll = ScrollBar1
End Property
Friend Property Get HScroll() As ScrollBar
    Set HScroll = ScrollBar2
End Property

Friend Property Get hProc() As Long
    hProc = pOldProc
End Property
Friend Property Let hProc(ByVal RHS As Long)
    pOldProc = RHS
End Property

Private Function DrawableRect() As RECT
    DrawableRect = RECT(LineColumnWidth, 0, UsercontrolWidth, UsercontrolHeight)
End Function
Private Property Get UsercontrolWidth() As Long
    UsercontrolWidth = IIf(ScrollBar1.Visible, (UserControl.Width - ScrollBar1.Width), UserControl.Width) - LineColumnWidth
    If UsercontrolWidth < 0 Then UsercontrolWidth = 0
End Property
Private Property Get UsercontrolHeight() As Long
    UsercontrolHeight = IIf(ScrollBar2.Visible, (UserControl.Height - ScrollBar2.Height), UserControl.Height)
    If UsercontrolHeight < 0 Then UsercontrolHeight = 0
End Property

Private Property Get LineColumnWidth() As Long
    If pLineNumbers Then
        If ((UsercontrolHeight \ TextHeight) + 1) > LineCount - LineFirstVisible Then
            LineColumnWidth = TextWidth("." & ((UsercontrolHeight \ TextHeight) + 1) & ".")
        Else
            LineColumnWidth = TextWidth("." & LineCount & ".")
        End If
    Else
        LineColumnWidth = 0
    End If
End Property
'Private Property Get LineColumnText() As String
'    Dim cnt As Long
'    For cnt = LineFirstVisible To ((UsercontrolHeight \ TextHeight) + 1)
'        LineColumnText=LineColumnText " " & (cnt + 1) & ". ", GetSysColor(COL
'    Next
'End Property


Public Property Get hWnd() As Long ' _
Returns the standard windows handle of the control.
Attribute hWnd.VB_Description = "Returns the standard windows handle of the control."
    hWnd = UserControl.hWnd
End Property

Public Property Get ScrollBars() As vbScrollBars ' _
Gets the value that determines how scroll bars are used and displayed for the control.
Attribute ScrollBars.VB_Description = "Gets the value that determines how scroll bars are used and displayed for the control."
    ScrollBars = pScrollBars
End Property
Public Property Let ScrollBars(ByVal RHS As vbScrollBars) ' _
Sets the behavior of and whether or not scroll bars are used for the control.
Attribute ScrollBars.VB_Description = "Sets the behavior of and whether or not scroll bars are used for the control."
    If pScrollBars <> RHS Then
        pScrollBars = RHS
        SetScrollBars
        Paint
        PaintBuffer
    End If
End Property

Public Property Get TextHeight(Optional ByVal StrText As String = "Iy") As Long ' _
Returns the twip measurement height using the current font size and line spacing vertically of SrrText.
Attribute TextHeight.VB_Description = "Returns the twip measurement height using the current font size and line spacing vertically of SrrText."
    TextHeight = UserControl.TextHeight(StrText)
End Property
Public Property Get TextWidth(Optional ByVal StrText As String = "W") As Long ' _
Returns the twip measurement width using the current font size and letter spacing horizontally of SrrText.
Attribute TextWidth.VB_Description = "Returns the twip measurement width using the current font size and letter spacing horizontally of SrrText."
    TextWidth = UserControl.TextWidth(Replace(StrText, Chr(9), TabSpace))
End Property
Public Property Get MultipleLines() As Boolean ' _
Returns whether or not this text control allows multiple lines in Text, delimited by line feeds.
Attribute MultipleLines.VB_Description = "Returns whether or not this text control allows multiple lines in Text, delimited by line feeds."
    MultipleLines = pMultiLine
End Property
Public Property Let MultipleLines(ByVal RHS As Boolean) ' _
Sets whehter or not this text control allows multiple lines in Text, delimited by line feeds.
Attribute MultipleLines.VB_Description = "Sets whehter or not this text control allows multiple lines in Text, delimited by line feeds."
    If pMultiLine <> RHS Then
        pMultiLine = RHS
        If Not pMultiLine Then
    '        pText = Replace(pText, vbLf, "")
            Dim tText As New Strands
           
            tText.Concat Convert(Replace(Replace(Convert(pText.Partial), vbCrLf, vbLf), vbLf, ""))
    
            If tText.Length > 0 Then
                pText.Clone tText
            Else
                pText.Reset
            End If
            Set tText = Nothing
            
            pOffsetY = 0
        End If
        UserControl_Paint
        PaintBuffer
    End If
End Property
Public Property Get HideSelection() As Boolean ' _
Gets whether or not the selection highlight will be hidden when the control is not in focus.
Attribute HideSelection.VB_Description = "Gets whether or not the selection highlight will be hidden when the control is not in focus."
    HideSelection = pHideSelection
End Property
Public Property Let HideSelection(ByVal RHS As Boolean) ' _
Sets whether or not the selection highlight will be hidden when the control is not in focus.
Attribute HideSelection.VB_Description = "Sets whether or not the selection highlight will be hidden when the control is not in focus."
    pHideSelection = RHS
End Property

Public Property Get ScrollToCaret() As Boolean ' _
Gets whether or not the caret forces the scrolling to keep it with in visibility.
Attribute ScrollToCaret.VB_Description = "Gets whether or not the caret forces the scrolling to keep it with in visibility."
    ScrollToCaret = pScrollToCaret
End Property
Public Property Let ScrollToCaret(ByVal RHS As Boolean) ' _
Sets whether or not the caret forces the scrolling to keep it with in visibility.
Attribute ScrollToCaret.VB_Description = "Sets whether or not the caret forces the scrolling to keep it with in visibility."
    pScrollToCaret = RHS
End Property
Private Function VisibleText() As String
    Dim tmp As Long
    Dim tmp2 As Long
    If pText.Length > 0 Then
        tmp2 = LineFirstVisible
        If tmp2 > 0 Then
            tmp = pText.Poll(Asc(vbLf), tmp2 + 1)
        End If
        tmp2 = pText.Poll(Asc(vbLf), tmp2 + UsercontrolHeight \ TextHeight + 1)
        
        If tmp2 - tmp > 0 Then
            VisibleText = Convert(pText.Partial(tmp, tmp2 - tmp))
        End If
    End If
End Function

Friend Property Get GetCanvasWidth(Optional ByVal Changed As Boolean = False) As Long
    Static pCanvasWidth As Long
    If Changed Or pCanvasWidth = 0 Then
        If pText.Length > 0 Then
            Dim cnt As Long
            Dim max As Long
            Dim tmp As Long
            For cnt = 0 To LineCount - 1
                tmp = LineLength(cnt)
                If tmp > max Then
                    'pCanvasWidth = me.TextWidth(LineText(cnt))
                    pCanvasWidth = Me.TextWidth(String(tmp, "W"))
                    max = tmp
                End If
            Next
            pCanvasWidth = pCanvasWidth + (UsercontrolWidth / 2)
            
            'pCanvasWidth = me.TextWidth(VisibleText) + (UsercontrolWidth / 2)
            
        Else
            pCanvasWidth = (UsercontrolWidth / 2)
        End If
    End If
    GetCanvasWidth = pCanvasWidth
End Property

Friend Property Get GetCanvasHeight(Optional ByVal Changed As Boolean = False) As Long
    Static pCanvasHeight As Long
    If Changed Or pCanvasHeight = 0 Then
        If pText.Length > 0 Then
            pCanvasHeight = (Me.TextHeight("Iy") * LineCount) + (UsercontrolHeight / 2)
        Else
            pCanvasHeight = (UsercontrolHeight / 2)
        End If
    End If
    GetCanvasHeight = pCanvasHeight
End Property

Public Property Get CanvasWidth() As Long ' _
Gets the width of the textual space needed and defined by the Text and Width property.
Attribute CanvasWidth.VB_Description = "Gets the width of the textual space needed and defined by the Text and Width property."
    CanvasWidth = GetCanvasWidth
End Property
Public Property Get CanvasHeight() As Long ' _
Gets the height of the textual space needed and defined by the Text and Height property.
Attribute CanvasHeight.VB_Description = "Gets the height of the textual space needed and defined by the Text and Height property."
    CanvasHeight = GetCanvasHeight
End Property

Public Property Get OffsetX() As Long ' _
Gets the current scroll bar offset of the horizontal canvas width drawn in visibility.
Attribute OffsetX.VB_Description = "Gets the current scroll bar offset of the horizontal canvas width drawn in visibility."
    OffsetX = pOffsetX
End Property
Public Property Let OffsetX(ByVal RHS As Long) ' _
Sets the current scroll bar offset of the horizontal canvas width drawn in visibility.
Attribute OffsetX.VB_Description = "Sets the current scroll bar offset of the horizontal canvas width drawn in visibility."
    If pOffsetX <> RHS Then
        pOffsetX = RHS
        UserControl_Paint
        PaintBuffer
    End If
End Property

Public Property Get OffsetY() As Long ' _
Gets the current scroll bar offset of the vertical canvas height drawn in visibility.
Attribute OffsetY.VB_Description = "Gets the current scroll bar offset of the vertical canvas height drawn in visibility."
    OffsetY = pOffsetY
End Property
Public Property Let OffsetY(ByVal RHS As Long) ' _
Sets the current scroll bar offset of the vertical canvas height drawn in visibility.
Attribute OffsetY.VB_Description = "Sets the current scroll bar offset of the vertical canvas height drawn in visibility."
    If pOffsetY <> RHS Then
        pOffsetY = RHS
        UserControl_Paint
        PaintBuffer
    End If
End Property

Private Sub CanvasValidate(Optional ByVal RecalcSizeOf As Boolean = True)

    If pOffsetX > 0 Or UsercontrolWidth > GetCanvasWidth(RecalcSizeOf) Then pOffsetX = 0
    If pOffsetY > 0 Or UsercontrolHeight > GetCanvasHeight(RecalcSizeOf) Then pOffsetY = 0

   ' If pOffsetX < -(CanvasWidth - UsercontrolWidth)  Then pOffsetX = -(CanvasWidth - UsercontrolWidth)
   ' If pOffsetY < -(CanvasHeight - UsercontrolHeight) Then pOffsetY = -(CanvasHeight - UsercontrolHeight)
End Sub

Public Property Get Forecolor() As OLE_COLOR ' _
Gets the default forecolor of the text display when a specific color table coloring is not used.
Attribute Forecolor.VB_Description = "Gets the default forecolor of the text display when a specific color table coloring is not used."
    Forecolor = pForecolor
End Property
Public Property Let Forecolor(ByVal RHS As OLE_COLOR) ' _
Sets the default forecolor of the text display when a specific color table coloring is not used.
Attribute Forecolor.VB_Description = "Sets the default forecolor of the text display when a specific color table coloring is not used."
    If pForecolor <> RHS Then
        pForecolor = RHS
        UserControl_Paint
        PaintBuffer
    End If
End Property

Public Property Get Backcolor() As OLE_COLOR ' _
Gets the default background color of the text display when a specific color table coloring is not used.
Attribute Backcolor.VB_Description = "Gets the default background color of the text display when a specific color table coloring is not used."
    Backcolor = pBackcolor
End Property
Public Property Let Backcolor(ByVal RHS As OLE_COLOR) ' _
Sets the default background color of the text display when a specific color table coloring is not used.
Attribute Backcolor.VB_Description = "Sets the default background color of the text display when a specific color table coloring is not used."
    If pBackcolor <> RHS Then
        pBackcolor = RHS
        UserControl.Backcolor = RHS
        UserControl_Paint
        PaintBuffer
    End If
End Property

Public Property Get SelText() As String ' _
Gets the selected text, the portion of text that is highlighted.
Attribute SelText.VB_Description = "Gets the selected text, the portion of text that is highlighted."
    If pText.Length > 0 Then
        If pSel.StopPos < pSel.StartPos And (pSel.StartPos - pSel.StopPos) > 0 Then
            SelText = Convert(pText.Partial(pSel.StopPos, (pSel.StartPos - pSel.StopPos)))
        ElseIf pSel.StopPos >= pSel.StartPos And (pSel.StopPos - pSel.StartPos) > 0 Then
            SelText = Convert(pText.Partial(pSel.StartPos, (pSel.StopPos - pSel.StartPos)))
        End If
    End If
    SelText = Replace(SelText, vbLf, vbCrLf)
End Property

Public Property Let SelText(ByVal RHS As String) ' _
Sets the selected text, the portion of text that is highlighted, changing the selection if applicable.
Attribute SelText.VB_Description = "Sets the selected text, the portion of text that is highlighted, changing the selection if applicable."

    Dim tText As New Strands
    
    If pSel.StartPos > pSel.StopPos Then
        Swap pSel.StartPos, pSel.StopPos
    End If
    'pText = Left(pText, pSel.StartPos) & RHS & Mid(pText, pSel.StopPos + 1)
    
    If pSel.StartPos > 0 And pText.Length > pSel.StartPos Then
        tText.Concat pText.Partial(0, pSel.StartPos)
    End If
    
    tText.Concat Convert(Replace(RHS, vbCrLf, vbLf))
    
    If pText.Length > pSel.StopPos Then
        tText.Concat pText.Partial(pSel.StopPos)
    End If

    If tText.Length > 0 Then
        pText.Clone tText
    Else
        pText.Reset
    End If
    Set tText = Nothing
    
    pSel.StopPos = pSel.StartPos + Len(Replace(RHS, vbCrLf, vbLf))
    Swap pSel.StartPos, pSel.StopPos
    InvalidateCursor

End Property

Public Property Get SelStart() As Long ' _
Gets the selection start of the highlighted portion of text.
Attribute SelStart.VB_Description = "Gets the selection start of the highlighted portion of text."
    If pSel.StopPos < pSel.StartPos Then
        SelStart = pSel.StopPos
    Else
        SelStart = pSel.StartPos
    End If
End Property
Public Property Let SelStart(ByVal RHS As Long) ' _
Sets the selection start of the highlighted portion of text.
Attribute SelStart.VB_Description = "Sets the selection start of the highlighted portion of text."
    If pSel.StopPos < pSel.StartPos Then
        pSel.StopPos = RHS
    Else
        pSel.StartPos = RHS
    End If
    InvalidateCursor
End Property
Public Property Get SelLength() As Long ' _
Gets the selection length of the highlighted portion of text.
Attribute SelLength.VB_Description = "Gets the selection length of the highlighted portion of text."
    If pSel.StopPos < pSel.StartPos Then
        SelLength = pSel.StartPos - pSel.StopPos
    Else
        SelLength = pSel.StopPos - pSel.StartPos
    End If
End Property
Public Property Let SelLength(ByVal RHS As Long) ' _
Sets the selection length of the highlighted portion of text.
Attribute SelLength.VB_Description = "Sets the selection length of the highlighted portion of text."
    If pSel.StopPos < pSel.StartPos Then
        pSel.StartPos = pSel.StopPos + RHS
    Else
        pSel.StopPos = pSel.StartPos + RHS
    End If
    InvalidateCursor
End Property

Public Property Get Text() ' _
Gets the text contents of the control in the NTNodes10.Staands object type.
Attribute Text.VB_Description = "Gets the text contents of the control in the NTNodes10.Staands object type."
    Set Text = pText
End Property
Public Property Let Text(ByVal RHS) ' _
Sets the text contents of the control by String data type.
Attribute Text.VB_Description = "Sets the text contents of the control by String data type."
    pText.Reset
    If Len(RHS) > 0 Then
        pText.Concat Convert(Replace(Replace(RHS, vbCrLf, vbLf), vbLf, IIf(pMultiLine, vbLf, "")))
        RaiseEventChange True

    End If
    
End Property
Public Property Set Text(ByRef RHS As Strands) ' _
Sets the text contents of the control in the NTNodes10.Staands object type.
Attribute Text.VB_Description = "Sets the text contents of the control in the NTNodes10.Staands object type."
    Select Case TypeName(RHS)
        Case "Strands", "NTNodes10.Strands", "NTControls30.Strands"
            pText.Reset
            If RHS.Length > 0 Then
                pText.Concat Convert(Replace(Replace(Convert(RHS.Partial), vbCrLf, vbLf), vbLf, IIf(pMultiLine, vbLf, "")))
                RaiseEventChange True
            End If
    End Select
    
End Property

Public Property Get Font() As StdFont ' _
Gets the font that the text is displayed in.
Attribute Font.VB_Description = "Gets the font that the text is displayed in."
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByRef newVal As StdFont) ' _
Sets the font that the text is displayed in.
Attribute Font.VB_Description = "Sets the font that the text is displayed in."
    Set UserControl.Font = newVal
    UserControl.FontBold = UserControl.Font.Bold
    UserControl.FontItalic = UserControl.Font.Italic
    UserControl.FontName = UserControl.Font.name
    UserControl.FontSize = UserControl.Font.Size
    UserControl.FontStrikethru = UserControl.Font.Strikethrough
    UserControl.FontUnderline = UserControl.Font.Underline
    Set pBackBuffer.Font = UserControl.Font
    UserControl_Paint
    PaintBuffer
End Property


Private Property Let IControl_hProc(ByVal RHS As Long)
    Me.hProc = RHS
End Property

Private Property Get IControl_hProc() As Long
    IControl_hProc = Me.hProc
End Property

Private Property Get IControl_hWnd() As Long
    IControl_hWnd = Me.hWnd
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
Private Function InsertCharacter(ByVal StrText As String) As String
    InsertCharacter = IIf(StrText = vbLf, " ", StrText)
End Function
Private Sub Timer1_Timer()
    Static cursorBlink As Boolean
    Static lastLoc As POINTAPI
    
    Dim newloc As POINTAPI
    newloc = CaretLocation
    
    If ((Not cursorBlink) Or (Not hasFocus)) Or ((newloc.X <> lastLoc.X) Or (newloc.Y <> lastLoc.Y)) Then
        If insertMode Then
            
            ClipPrintText lastLoc.X, lastLoc.Y, InsertCharacter(Convert(pText.Partial(pSel.StartPos, 1))), pForecolor, False
            'ClipPrintText lastLoc.X, lastLoc.Y, IIf(Mid(pText, pSel.StartPos + 1, 1) = vbLf, " ", Mid(pText, pSel.StartPos + 1, 1)), GetSysColor(COLOR_WINDOWTEXT), False
        Else
            ClipLineDraw lastLoc.X, lastLoc.Y, lastLoc.X, (lastLoc.Y + TextHeight), pBackcolor
        End If
    End If

    Static lastSel As RangeType
    If lastSel.StartPos <> pSel.StartPos Or lastSel.StopPos <> pSel.StopPos Then
        MakeCaretVisible newloc, True
        cursorBlink = False
    End If
    lastSel.StartPos = pSel.StartPos
    lastSel.StopPos = pSel.StopPos
    
    Paint

    If (((cursorBlink Or ((newloc.X <> lastLoc.X) Or (newloc.Y <> lastLoc.Y))) And hasFocus)) And Enabled Then
        If insertMode Then
            ClipPrintText newloc.X, newloc.Y, InsertCharacter(Convert(pText.Partial(pSel.StartPos, 1))), pForecolor, True
            'ClipPrintText newloc.X, newloc.Y, IIf(Mid(pText, pSel.StartPos + 1, 1) = vbLf, " ", Mid(pText, pSel.StartPos + 1, 1)), GetSysColor(COLOR_WINDOWTEXT), True
        Else
            ClipLineDraw newloc.X, newloc.Y, newloc.X, (newloc.Y + TextHeight), pForecolor
        End If
    ElseIf ((Not cursorBlink) Or (Not hasFocus)) Or ((newloc.X <> lastLoc.X) Or (newloc.Y <> lastLoc.Y)) Then
        If insertMode Then
            ClipPrintText lastLoc.X, lastLoc.Y, InsertCharacter(Convert(pText.Partial(pSel.StartPos, 1))), pForecolor, False
            'ClipPrintText lastLoc.X, lastLoc.Y, IIf(Mid(pText, pSel.StartPos + 1, 1) = vbLf, " ", Mid(pText, pSel.StartPos + 1, 1)), GetSysColor(COLOR_WINDOWTEXT), False
        Else
            ClipLineDraw lastLoc.X, lastLoc.Y, lastLoc.X, (lastLoc.Y + TextHeight), pBackcolor
        End If
    End If
    
    lastLoc.X = newloc.X
    lastLoc.Y = newloc.Y
    
    cursorBlink = Not cursorBlink

    If xUndoDelay > 0 Then

        If CSng(Timer - xLatency) * 1000 >= xUndoDelay Then
            AddUndo
            xLatency = Timer
        End If
    End If
    
    If Not Timer1.Enabled Then
        Timer1.Enabled = cursorBlink
        If Not Timer1.Enabled Then
            Timer1_Timer
        End If
    End If
End Sub
Private Function MakeCaretVisible(ByRef loc As POINTAPI, ByVal LargeJump As Boolean) As Boolean
    If Enabled Then
        If pScrollToCaret And (Not ClippingWouldDraw(DrawableRect, RECT(loc.X, loc.Y, loc.X + TextWidth, loc.Y + TextHeight), True)) Then
            If loc.X < 1 Then
                If LargeJump Then
                    pOffsetX = pOffsetX + ((1 - loc.X) + (UsercontrolWidth / 2))
                Else
                    pOffsetX = pOffsetX + ((1 - loc.X) + ScrollBar2.SmallChange)
                End If
                If ScrollBar2.Visible And ScrollBar2.Value <> -pOffsetX Then ScrollBar2.Value = -pOffsetX
                MakeCaretVisible = True
            ElseIf loc.X > UsercontrolWidth Then
                If LargeJump Then
                    pOffsetX = pOffsetX - ((loc.X - UsercontrolWidth) + (UsercontrolWidth / 2))
                Else
                    pOffsetX = pOffsetX - ((loc.X - UsercontrolWidth) + ScrollBar2.SmallChange)
                End If
                If ScrollBar2.Visible And ScrollBar2.Value <> -pOffsetX Then ScrollBar2.Value = -pOffsetX
                MakeCaretVisible = True
            End If
            If loc.Y < 1 Then
                If LargeJump Then
                    pOffsetY = pOffsetY + (((((1 - loc.Y) + (UsercontrolHeight / 2)) \ TextHeight)) * TextHeight)
                Else
                    pOffsetY = pOffsetY + (((((1 - loc.Y) + ScrollBar1.SmallChange) \ TextHeight)) * TextHeight)
                End If
                If ScrollBar1.Visible And ScrollBar1.Value <> -pOffsetY Then ScrollBar1.Value = -pOffsetY
                MakeCaretVisible = True
            ElseIf loc.Y > (UsercontrolHeight - TextHeight) Then
                If LargeJump Then
                    pOffsetY = pOffsetY - (((((loc.Y - UsercontrolHeight) + (UsercontrolHeight / 2)) \ TextHeight)) * TextHeight)
                Else
                    pOffsetY = pOffsetY - (((((loc.Y - UsercontrolHeight) + ScrollBar1.SmallChange) \ TextHeight)) * TextHeight)
                End If
                If ScrollBar1.Visible And ScrollBar1.Value <> -pOffsetY Then ScrollBar1.Value = -pOffsetY
                MakeCaretVisible = True
            End If
            
            If MakeCaretVisible = True Then
                UserControl_Paint
                PaintBuffer
            End If
        End If
    End If
End Function
Private Function ClippingWouldDraw(ByRef rct1 As RECT, ByRef rct2 As RECT, Optional ByVal FullDrawOnly As Boolean = False) As Boolean
    Dim rct As RECT
    ClippingWouldDraw = (IntersectRect(rct, rct1, rct2) <> 0)
    If Not ClippingWouldDraw Then
        If Not FullDrawOnly Then
            ClippingWouldDraw = (PtInRect(rct1, rct2.Left, rct2.Top) <> 0 Or PtInRect(rct1, rct2.Right, rct2.Top) <> 0 Or _
                        PtInRect(rct1, rct2.Left, rct2.Bottom) <> 0 Or PtInRect(rct1, rct2.Right, rct2.Bottom) <> 0)
            If Not ClippingWouldDraw Then
                ClippingWouldDraw = (rct1.Left = rct2.Left) Or (rct1.Top = rct2.Top) Or (rct1.Right = rct2.Right) Or (rct1.Bottom = rct2.Bottom)
            End If
        Else
            ClippingWouldDraw = (PtInRect(rct1, rct2.Left, rct2.Top) <> 0 And PtInRect(rct1, rct2.Right, rct2.Top) <> 0 And _
                        PtInRect(rct1, rct2.Left, rct2.Bottom) <> 0 And PtInRect(rct1, rct2.Right, rct2.Bottom) <> 0)
        End If
    ElseIf FullDrawOnly Then
        ClippingWouldDraw = False
    End If
End Function
Private Function ClipPrintText(ByVal X1 As Single, ByVal Y1 As Single, ByVal StrText As String, Optional Color As Variant, Optional ByVal BoxFill As Boolean = False) As Boolean

    If BoxFill Then
        If ClipLineDraw(X1, Y1, (Me.TextWidth(StrText) + X1), (Me.TextHeight(StrText) + Y1), Color, True) Then
            pBackBuffer.DrawText X1 / Screen.TwipsPerPixelX + 1, Y1 / Screen.TwipsPerPixelY, Replace(StrText, Chr(9), TabSpace), pBackcolor
            'UserControl.CurrentX = X1
            'UserControl.CurrentY = Y1
            'UserControl.Print StrText
            ClipPrintText = True
        End If
    ElseIf ClippingWouldDraw(DrawableRect, RECT(X1, Y1, (Me.TextWidth(StrText) + X1), (Me.TextHeight(StrText) + Y1))) Then
        pBackBuffer.DrawText X1 / Screen.TwipsPerPixelX + 1, Y1 / Screen.TwipsPerPixelY, Replace(StrText, Chr(9), TabSpace), Color
        'UserControl.CurrentX = X1
        'UserControl.CurrentY = Y1
        'UserControl.Print StrText
        ClipPrintText = True
    End If
    
End Function

Private Function ClipPrintTextBlock(ByVal X1 As Single, ByVal Y1 As Single, ByVal StrText As String, Optional Color As Variant, Optional ByVal BoxFill As Boolean = False) As POINTAPI
    With ClipPrintTextBlock
        .X = X1
        .Y = Y1
        Dim outLine As String
        If StrText <> "" Then
            Do While InStr(StrText, vbLf) > 0
                outLine = RemoveNextArg(StrText, vbLf)
                ClipPrintText .X, .Y, outLine, Color, BoxFill
                .X = pOffsetX + LineColumnWidth
                .Y = .Y + Me.TextHeight(outLine)
            Loop
            If StrText <> "" Then
                ClipPrintText .X, .Y, StrText, Color, BoxFill
                .X = .X + Me.TextWidth(StrText)
            End If
        End If
    End With
End Function

Private Function ClipLineDraw(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, Optional Color As Variant, Optional ByVal BoxFill As Boolean = False) As Boolean
    If ClippingWouldDraw(DrawableRect, RECT(X1, Y1, X2, Y2)) Then
        If BoxFill Then
            pBackBuffer.DrawLine X1 / Screen.TwipsPerPixelX, Y1 / Screen.TwipsPerPixelY, X2 / Screen.TwipsPerPixelX, Y2 / Screen.TwipsPerPixelY, Color, BF
            'UserControl.Line (X1, Y1)-(X2, Y2), Color, BF
        Else
            pBackBuffer.DrawLine X1 / Screen.TwipsPerPixelX, Y1 / Screen.TwipsPerPixelY, X2 / Screen.TwipsPerPixelX, Y2 / Screen.TwipsPerPixelY, Color
            'UserControl.Line (X1, Y1)-(X2, Y2), Color
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

Private Sub RefreshEditMenu()
    mnuCut.Enabled = Enabled And Not Locked
    mnuCopy.Enabled = (SelLength > 0)
    mnuPaste.Enabled = Enabled And Not Locked
    mnuDelete.Enabled = Enabled And Not Locked
    mnuSelectAll.Enabled = Enabled
    mnuIndent.Enabled = (Enabled And Not Locked And (CountWord(SelText, vbCrLf) > 1))
    mnuUnindent.Enabled = (Enabled And Not Locked And (CountWord(SelText, vbCrLf) > 1))
End Sub

Private Sub UserControl_Initialize()
    Cancel = True
        

    Set pText = New Strands
    
    xUndoStack = 150
    xUndoDelay = 1000
    
    ResetUndoRedo
        
    SystemParametersInfo SPI_GETKEYBOARDSPEED, 0, keySpeed, 0
    Timer1.Interval = keySpeed * 10

    Set pBackBuffer = New Backbuffer
    pBackBuffer.hWnd = UserControl.hWnd
    pBackBuffer.Forecolor = ConvertColor(SystemColorConstants.vbWindowText)
    pBackBuffer.Backcolor = ConvertColor(SystemColorConstants.vbWindowBackground)
    Set pBackBuffer.Font = UserControl.Font
    Set pScroll1Buffer = ScrollBar1.Backbuffer
    Set pScroll2Buffer = ScrollBar2.Backbuffer
    ScrollBar1.AutoRedraw = False
    ScrollBar2.AutoRedraw = False
    pScroll1Buffer.hdc = pBackBuffer.hdc
    pScroll2Buffer.hdc = pBackBuffer.hdc
    pTabSpace = "    "
    pLineNumbers = True
'    Set pForecolors = New Collection
'    Set pBackcolors = New Collection

    Cancel = False

    Hook Me
End Sub

Private Sub UserControl_InitProperties()
    pForecolor = GetSysColor(COLOR_WINDOWTEXT)
    pBackcolor = GetSysColor(COLOR_WINDOW)
    UserControl.Backcolor = GetSysColor(COLOR_WINDOW)
    pScrollToCaret = True
    pHideSelection = True
    UserControl.Font.name = "Lucida Console"
    Set pBackBuffer.Font = UserControl.Font
    pMultiLine = True
    pScrollBars = vbScrollBars.Both
    pEnabled = True
    pTabSpace = "    "
    xUndoStack = 150
    xUndoDelay = 1000
    xUndoDirty = False
    pLineNumbers = True
    ResetUndoRedo
    
End Sub

Public Function LineOffset(ByVal LineIndex As Long) As Long ' _
Returns the offset amount of characters upto a line, specified by the zero based LineIndex. Example, LineOffset(0)=0.
Attribute LineOffset.VB_Description = "Returns the offset amount of characters upto a line, specified by the zero based LineIndex. Example, LineOffset(0)=0."
    If pText.Length > 0 Then
        LineOffset = pText.Poll(Asc(vbLf), LineIndex)
        If LineIndex > 0 Then LineOffset = LineOffset + 1
    End If
End Function

Public Function LineLength(ByVal LineIndex As Long) As Long ' _
Returns the length of characters with-in a line, specifiied by the zero based LineIndex.
Attribute LineLength.VB_Description = "Returns the length of characters with-in a line, specifiied by the zero based LineIndex."
    If pText.Length > 0 Then
        LineLength = (pText.Poll(Asc(vbLf), LineIndex + 1) - pText.Poll(Asc(vbLf), LineIndex))
        If LineIndex > 0 Then LineLength = LineLength - 1
    End If
End Function

Public Function LineText(ByVal LineIndex As Long) As String ' _
Returns the text with-in a line, specified by the zero based LineIndex.
Attribute LineText.VB_Description = "Returns the text with-in a line, specified by the zero based LineIndex."
    Dim lPos As Long
    lPos = LineLength(LineIndex)
    If lPos > 0 Then
        LineIndex = LineOffset(LineIndex)
        If LineIndex > 0 Then
            'LineText = Mid(pText, LineIndex, lPos)
            LineText = Convert(pText.Partial(LineIndex, lPos))
        Else
            'LineText = Mid(pText, 1, lPos)
            LineText = Convert(pText.Partial(0, lPos))
        End If
    End If
End Function

Public Function LineIndex(Optional ByVal CharIndex As Long = -1) As Long ' _
Returns the zero based index line number of which the given zero based character index falls upon with-in Text.
Attribute LineIndex.VB_Description = "Returns the zero based index line number of which the given zero based character index falls upon with-in Text."
    If CharIndex = -1 Then CharIndex = pSel.StartPos
    If CharIndex > 0 Then
        'LineIndex = CountWord(Left(pText, CharIndex), vbLf)
        LineIndex = pText.Pass(Asc(vbLf), 0, CharIndex)
    End If
End Function

Public Function LineCount() As Long ' _
Returns the numerical count of how many lines, delimited by line feeds, that exists with-in Text.
Attribute LineCount.VB_Description = "Returns the numerical count of how many lines, delimited by line feeds, that exists with-in Text."
    If pText.Length > 0 Then
        LineCount = pText.Pass(Asc(vbLf)) + 1
    End If
End Function

Private Function CaretLocation(Optional ByVal AtCharPos As Long = -1) As POINTAPI
    Dim part As String
    If pSel.StartPos <= 0 Then pSel.StartPos = 0
    If AtCharPos = -1 Then AtCharPos = pSel.StartPos
    If AtCharPos > 0 And AtCharPos <= pText.Length Then
        Dim cnt As Long
        cnt = pText.Pass(Asc(vbLf), 0, AtCharPos)
        If cnt >= 0 Then
            CaretLocation.Y = (TextHeight * cnt) + pOffsetY
            part = Left(LineText(cnt), (pText.Length - LineOffset(cnt)) - (pText.Length - AtCharPos))
        Else
            CaretLocation.Y = pOffsetY
        End If
    End If
    CaretLocation.X = Me.TextWidth(part) + pOffsetX + LineColumnWidth
End Function

Public Function CaretFromPoint(ByVal X As Single, ByVal Y As Single) As Long ' _
Retruns the zero based character index position with-in Text based upon given pixel corrdinates, X and Y.
Attribute CaretFromPoint.VB_Description = "Retruns the zero based character index position with-in Text based upon given pixel corrdinates, X and Y."
    Dim lText As String
    X = (X - pOffsetX) - LineColumnWidth
    Y = ((Y - pOffsetY) \ TextHeight)
    If Y < LineCount() Then
        If X >= TextWidth(LineText(Y)) Then
            CaretFromPoint = LineOffset(Y) + Len(Replace(LineText(Y), vbLf, ""))
        Else
            lText = Replace(LineText(Y), vbLf, "")
            Do While lText <> ""
                If TextWidth(lText) - (TextWidth(Right(lText, 1)) / 2) < X Then Exit Do
                lText = Left(lText, Len(lText) - 1)
            Loop
            CaretFromPoint = LineOffset(Y) + Len(lText)
        End If
    Else
        CaretFromPoint = pText.Length
    End If
End Function

Public Function LineFirstVisible() As Long ' _
Returns the zero based line index number of the first visible line on the screen.
Attribute LineFirstVisible.VB_Description = "Returns the zero based line index number of the first visible line on the screen."
    LineFirstVisible = (-pOffsetY \ TextHeight)
End Function

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  
    'Debug.Print "KeyDown "; Convert(Me.Text.Partial); Me.SelStart; Me.SelLength
    RaiseEvent KeyDown(KeyCode, Shift)
    If KeyCode <> 0 Then
        Dim tText As Strands
        Dim lIndex As Long
        Dim txt As String
        Dim temp As Long
        
        Select Case KeyCode
            Case 93 'menu
                RefreshEditMenu
                PopupMenu mnuEdit, , 0, 0
            Case 13 'enter
                If pLocked Then Exit Sub
                
                
                If pSel.StartPos = pSel.StopPos Then
                    'pText = Left(pText, pSel.StartPos) & vbLf & Mid(pText, pSel.StartPos + 1)
                    
                    Set tText = New Strands
                    If pSel.StartPos > 0 Then tText.Concat pText.Partial(0, pSel.StartPos)
                    tText.Concat Convert(vbLf)
                    If pSel.StartPos < pText.Length Then tText.Concat pText.Partial(pSel.StartPos)

                    If tText.Length > 0 Then
                        pText.Clone tText
                    Else
                        pText.Reset
                    End If
                    Set tText = Nothing
                    
                Else
                    If pSel.StartPos > pSel.StopPos Then
                        Swap pSel.StartPos, pSel.StopPos
                    End If
                    'pText = Left(pText, pSel.StartPos) & vbLf & Mid(pText, pSel.StopPos + 1)

                    Set tText = New Strands
                    If pSel.StartPos > 0 Then tText.Concat pText.Partial(0, pSel.StartPos)
                    tText.Concat Convert(vbLf)
                    If pSel.StopPos < pText.Length Then tText.Concat pText.Partial(pSel.StopPos)

                    If tText.Length > 0 Then
                        pText.Clone tText
                    Else
                        pText.Reset
                    End If
                    Set tText = Nothing
                    
                    
                End If

                pSel.StartPos = pSel.StartPos + 1
                pSel.StopPos = pSel.StartPos
                RaiseEventChange

            Case 45 'insert
                If Shift = 0 Then
                    insertMode = Not insertMode
                ElseIf Shift = 1 Then 'shift
                    KeyCode = 0
                    Paste
                ElseIf Shift = 2 Then 'ctrl
                    KeyCode = 0
                    Copy
                End If
                
            Case 46, 8 'delete 'backspace
                If pLocked Then Exit Sub
                
                If pSel.StartPos = pSel.StopPos Then
                    If KeyCode = 8 And pSel.StartPos > 0 Then 'backspace
                        'pText = Left(pText, pSel.StartPos - 1) & Mid(pText, pSel.StartPos + 1)
    
                        Set tText = New Strands
                        If pSel.StartPos - 1 > 0 Then tText.Concat pText.Partial(0, pSel.StartPos - 1)
                        If pSel.StartPos < pText.Length Then tText.Concat pText.Partial(pSel.StartPos)

                        If tText.Length > 0 Then
                            pText.Clone tText
                        Else
                            pText.Reset
                        End If
                        Set tText = Nothing
                        
                        pSel.StartPos = pSel.StartPos - 1
                    ElseIf KeyCode = 46 Then 'delete
                        'pText = Left(pText, pSel.StartPos) & Mid(pText, pSel.StartPos + 2)

                        Set tText = New Strands
                        If pSel.StartPos > 0 Then tText.Concat pText.Partial(0, pSel.StartPos)
                        If pSel.StartPos + 1 < pText.Length Then tText.Concat pText.Partial(pSel.StartPos + 1)

                        If tText.Length > 0 Then
                            pText.Clone tText
                        Else
                            pText.Reset
                        End If
                        Set tText = Nothing
                        
                        pSel.StartPos = pSel.StartPos
                    End If
                Else 'delete or backspace
                    If pSel.StartPos > pSel.StopPos Then
                        Swap pSel.StartPos, pSel.StopPos
                    End If
                    'pText = Left(pText, pSel.StartPos) & Mid(pText, pSel.StopPos + 1)

                    Set tText = New Strands
                    If pSel.StartPos > 0 Then tText.Concat pText.Partial(0, pSel.StartPos)
                    If pSel.StopPos < pText.Length Then tText.Concat pText.Partial(pSel.StopPos)

                    If tText.Length > 0 Then
                        pText.Clone tText
                    Else
                        pText.Reset
                    End If
                    Set tText = Nothing
                    
                End If
                pSel.StopPos = pSel.StartPos
                
                RaiseEventChange
                
                
            Case 9 'tab
                If pLocked Then Exit Sub


                If SelLength > 0 Then
                
'                    lIndex = LineOffset(LineIndex(pSel.StartPos))
'                    temp = LineOffset(LineIndex(pSel.StopPos))
'                    If temp > lIndex Then
    
                        If Shift = 0 Then
                            Indenting SelStart, SelLength, Chr(9), True
                        ElseIf Shift = 1 Then
                            Indenting SelStart, SelLength, Chr(9) & Chr(8), True
                        End If
        
                        KeyCode = 0
                        
                        RaiseEventChange
'                    End Ife

                
                End If
                
            Case 36 'home
                temp = LineIndex(pSel.StartPos)
                lIndex = LineLength(temp) - Len(LTrimStrip(LTrimStrip(LineText(temp), vbTab), " "))
                temp = LineOffset(temp)
                If Not pSel.StartPos = temp + lIndex Then
                    pSel.StartPos = temp + lIndex
                Else
                    pSel.StartPos = temp
                End If
                If Shift = 0 Then pSel.StopPos = pSel.StartPos
                InvalidateCursor
            Case 35 'end
                pSel.StartPos = LineIndex(pSel.StartPos)
                pSel.StartPos = LineOffset(pSel.StartPos) + LineLength(pSel.StartPos) '- 1
                If Shift = 0 Then pSel.StopPos = pSel.StartPos
                InvalidateCursor
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
        
        If Shift = 2 And KeyCode <> 0 Then
            Select Case KeyCode
                Case 90 'ctrl+z
                    KeyCode = 0
                    Undo
                Case 88 'ctrl+x
                    KeyCode = 0
                    Cut
                Case 67 'ctrl+c
                    KeyCode = 0
                    Copy
                Case 86 'ctrl+v
                    KeyCode = 0
                    Paste
                Case 82, 89 'ctrl+r,ctrl+y
                    KeyCode = 0
                    Redo
                Case 65 'ctrl+a
                    KeyCode = 0
                    SelectAll
                Case 76 'ctrl+l
                    KeyCode = 0
                    ClearAll
            End Select
        End If
    End If
End Sub

Public Sub Indenting(ByVal SelStart As Long, ByVal SelLength As Long, Optional ByVal CharStr As String = "", Optional ByVal SelectAfter As Boolean = False)
    'i.e. selecting text and using tab to indent, or commenting and uncommenting selected text, default is to remove
    'any tab by setting CharSet to Chr(9) & Chr(8), SelectAfter forces the range edited to be selected when done

    Dim txt As String
    Dim temp As Long
    Dim newtxt As String

    Dim tmpSel As RangeType
    
    If CharStr = "" Then CharStr = Chr(9) & Chr(8)
    
    If pSel.StartPos > pSel.StopPos Then
        tmpSel.StopPos = pSel.StartPos
        tmpSel.StartPos = pSel.StopPos
    Else
        tmpSel.StartPos = pSel.StartPos
        tmpSel.StopPos = pSel.StopPos
    End If
    
    If tmpSel.StartPos > LineOffset(LineIndex(tmpSel.StartPos)) Then
        tmpSel.StartPos = LineOffset(LineIndex(tmpSel.StartPos))
    End If

    If tmpSel.StopPos - LineOffset(LineIndex(tmpSel.StopPos)) <= 1 Then tmpSel.StopPos = tmpSel.StopPos - 1
    If tmpSel.StopPos < LineOffset(LineIndex(tmpSel.StopPos)) + LineLength(LineIndex(tmpSel.StopPos)) Then
        tmpSel.StopPos = LineOffset(LineIndex(tmpSel.StopPos)) + LineLength(LineIndex(tmpSel.StopPos))
    End If
    pSel.StartPos = tmpSel.StartPos
    pSel.StopPos = tmpSel.StopPos
    
    For temp = LineIndex(tmpSel.StartPos) To LineIndex(tmpSel.StopPos)

        txt = LineText(temp)

        If Len(txt) > 0 Then
            If InStr(CharStr, Chr(8)) = 0 Then
                txt = CharStr & txt
                tmpSel.StopPos = tmpSel.StopPos + 1
            Else
                If Left(txt, Len(Replace(CharStr, Chr(8), ""))) = Replace(CharStr, Chr(8), "") Then
                    txt = Mid(txt, Len(Replace(CharStr, Chr(8), "")) + 1)
                    tmpSel.StopPos = tmpSel.StopPos - Len(Replace(CharStr, Chr(8), ""))
                End If
            End If
        End If

        newtxt = newtxt & txt & vbLf

    Next

    SelText = RTrimStrip(newtxt, vbLf)
    
    If pSel.StartPos > pSel.StopPos Then modCommon.Swap pSel.StartPos, pSel.StopPos

End Sub
Friend Sub KeyPageUp(ByRef Shift As Integer)
    pSel.StartPos = LineOffset(LineIndex - (UsercontrolHeight \ TextHeight))
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
    If pSel.StartPos >= 0 Then pSel.StartPos = pSel.StartPos - 1
    If Shift = 0 Then pSel.StopPos = pSel.StartPos
End Sub
Friend Sub KeyArrowRight(ByRef Shift As Integer)
    If pSel.StartPos < pText.Length Then pSel.StartPos = pSel.StartPos + 1
    If Shift = 0 Then pSel.StopPos = pSel.StartPos
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

    RaiseEvent KeyPress(KeyAscii)
    If pLocked Then KeyAscii = 0
    If KeyAscii = 9 And SelLength > 0 Then KeyAscii = 0

    If KeyAscii > 0 Then
        If KeyAscii = 22 Then
            KeyAscii = 0
        End If

        If KeyAscii = 5 Then
            KeyAscii = 0
        End If

        If KeyAscii = 3 Then
            KeyAscii = 0
        End If

        If KeyAscii = 24 Then
            KeyAscii = 0
        End If

        If KeyAscii = 26 Then
            KeyAscii = 0
        End If

        If KeyAscii = 18 Then
            KeyAscii = 0
        End If

    End If


    If KeyAscii <> 0 Then
        
        If ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or IsNumeric(Chr(KeyAscii))) Or _
            (Chr(KeyAscii) = "`" Or Chr(KeyAscii) = "~" Or Chr(KeyAscii) = "!" Or Chr(KeyAscii) = "@" Or Chr(KeyAscii) = "#" Or Chr(KeyAscii) = "$" Or _
            Chr(KeyAscii) = "%" Or Chr(KeyAscii) = "^" Or Chr(KeyAscii) = "&" Or Chr(KeyAscii) = "*" Or Chr(KeyAscii) = "(" Or Chr(KeyAscii) = ")" Or _
            Chr(KeyAscii) = "_" Or Chr(KeyAscii) = "-" Or Chr(KeyAscii) = "+" Or Chr(KeyAscii) = "=" Or Chr(KeyAscii) = "\" Or Chr(KeyAscii) = "|" Or _
            Chr(KeyAscii) = "[" Or Chr(KeyAscii) = "{" Or Chr(KeyAscii) = "]" Or Chr(KeyAscii) = "}" Or Chr(KeyAscii) = ":" Or Chr(KeyAscii) = ";" Or _
            Chr(KeyAscii) = """" Or Chr(KeyAscii) = "'" Or Chr(KeyAscii) = "<" Or Chr(KeyAscii) = "," Or Chr(KeyAscii) = ">" Or Chr(KeyAscii) = "." Or _
            Chr(KeyAscii) = "?" Or Chr(KeyAscii) = "/" Or Chr(KeyAscii) = Chr(9) Or Chr(KeyAscii) = " ") Then
                        
            Dim tText As Strands
            If insertMode Then

                Set tText = New Strands
                If pSel.StartPos > 0 Then tText.Concat pText.Partial(0, pSel.StartPos)
                tText.Concat Convert(Chr(KeyAscii))
                If pSel.StartPos + 1 < pText.Length Then tText.Concat pText.Partial(pSel.StartPos + 1)

                If tText.Length > 0 Then
                    pText.Clone tText
                Else
                    pText.Reset
                End If
                Set tText = Nothing
        
            ElseIf pSel.StartPos < pText.Length Then

                Set tText = New Strands
                If pSel.StartPos > pSel.StopPos Then
                    If pSel.StopPos > 0 Then tText.Concat pText.Partial(0, pSel.StopPos)
                    tText.Concat Convert(Chr(KeyAscii))
                    If pSel.StartPos < pText.Length Then tText.Concat pText.Partial(pSel.StartPos)
                    pSel.StartPos = pSel.StopPos
                Else
                    If pSel.StartPos > 0 Then tText.Concat pText.Partial(0, pSel.StartPos)
                    tText.Concat Convert(Chr(KeyAscii))
                    If pSel.StopPos < pText.Length Then tText.Concat pText.Partial(pSel.StopPos)
                End If
                
                If tText.Length > 0 Then
                    pText.Clone tText
                Else
                    pText.Reset
                End If
                Set tText = Nothing
            Else
                pText.Concat Convert(Chr(KeyAscii))
            End If
            
            pSel.StartPos = pSel.StartPos + 1
            pSel.StopPos = pSel.StartPos
            
            RaiseEventChange
            
            

        End If


            
    End If
    
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyUp(KeyCode, Shift)
    
    If KeyCode > 0 Then
        If KeyCode = 45 And Shift = 1 Then
            KeyCode = 0
        End If

        If KeyCode = 45 And Shift = 2 And (SelLength > 0) Then
            KeyCode = 0
        End If

        If KeyCode = 46 And (SelLength > 0) Then
            KeyCode = 0
        End If
    End If
End Sub

Private Sub UserControl_LostFocus()
    hasFocus = False
    If pHideSelection Then UserControl_Paint
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button = 1 And Shift = 0 Then
        'If hasFocus Or ((Not hasFocus) And (Not pHideSelection)) Then
        
        Dim lPos As Long
        pSel.StartPos = CaretFromPoint(X, Y)
        If Shift = 0 Then pSel.StopPos = pSel.StartPos
        
        InvalidateCursor

    End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)

    If Button = 1 And hasFocus Then

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
        Dim rct As RECT
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
            MakeCaretVisible newloc, False
        End If

    Else
        dragStart = 0
    End If
    
    If pLineNumbers And X < LineColumnWidth Then
        UserControl.MousePointer = 1
    Else
        UserControl.MousePointer = IIf(Enabled, 3, 1)
    End If
    
End Sub

Friend Sub InvalidateCursor()
    Timer1.Enabled = False
    If Not Cancel Then Timer1_Timer
End Sub

Private Sub RaiseEventChange(Optional ByVal PaintEvent As Boolean = False)
  
    RaiseEventSelChange
    
    If Not Cancel Then
        Cancel = True
        
        xUndoDirty = True
        
        If xUndoDelay = 0 Then
            AddUndo
        Else
            xLatency = Timer
        End If
        
        RaiseEvent Change
        
        Cancel = False
        
        If Not PaintEvent Then
            InvalidateCursor
        Else
            UserControl_Paint
            PaintBuffer
        End If
                
    End If
    
End Sub
Private Sub RaiseEventSelChange()
    
    If pLastSel.StartPos <> pSel.StartPos Or pLastSel.StopPos <> pSel.StopPos Then
        
        CanvasValidate True
        SetScrollBars

        RaiseEvent SelChange
        
        pLastSel.StartPos = pSel.StartPos
        pLastSel.StopPos = pSel.StopPos
    
    End If
    
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    dragStart = 0
    If Button = 2 And hasFocus Then
        RefreshEditMenu
        PopupMenu mnuEdit, , X, Y
    End If
    
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Unhook
'
'    Dim pt As POINTAPI
'
'    pt.X = (X / Screen.TwipsPerPixelX)
'    pt.Y = (Y / Screen.TwipsPerPixelY)
'    SelStart = SendMessageStruct(hWnd, EM_CHARFROMPOS, 0&, pt)
'    SelLength = 0
'
'    SelText = Data.GetData(ClipBoardConstants.vbCFText)
'
'    Hook
    
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
'    Dim pt As POINTAPI
'
'    pt.X = (X / Screen.TwipsPerPixelX)
'    pt.Y = (Y / Screen.TwipsPerPixelY)
'    SelStart = SendMessageStruct(hWnd, EM_CHARFROMPOS, 0&, pt)
    
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
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
    If AutoRedraw Then
        Paint
    End If

    RaiseEvent Paint
End Sub
Friend Sub PaintBuffer()

    If ScrollBar1.Visible And ScrollBar2.Visible Then
        'pBackBuffer.DrawFrame (UsercontrolWidth / Screen.TwipsPerPixelX), (UsercontrolHeight / Screen.TwipsPerPixelY), (UserControl.Width / Screen.TwipsPerPixelX), (UserControl.Height / Screen.TwipsPerPixelY), DFC_SCROLL, DFCS_SCROLLSIZEGRIP
        DrawFrameControl pBackBuffer.hdc, RECT((UsercontrolWidth / Screen.TwipsPerPixelX) - 1, (UsercontrolHeight / Screen.TwipsPerPixelY) - 1, (UserControl.Width / Screen.TwipsPerPixelX), (UserControl.Height / Screen.TwipsPerPixelY)), DFC_SCROLL, DFCS_SCROLLSIZEGRIP
    End If

    pScroll1Buffer.Paint (ScrollBar1.Left / Screen.TwipsPerPixelX), (ScrollBar1.Top / Screen.TwipsPerPixelY), ((ScrollBar1.Left + ScrollBar1.Width) / Screen.TwipsPerPixelX), ((ScrollBar1.Top + ScrollBar1.Height) / Screen.TwipsPerPixelY)
    pScroll2Buffer.Paint (ScrollBar2.Left / Screen.TwipsPerPixelX), (ScrollBar2.Top / Screen.TwipsPerPixelY), ((ScrollBar2.Left + ScrollBar2.Width) / Screen.TwipsPerPixelX), ((ScrollBar2.Top + ScrollBar2.Height) / Screen.TwipsPerPixelY)
    pBackBuffer.Paint 0, 0, ((UsercontrolWidth + LineColumnWidth) / Screen.TwipsPerPixelX) + 1, (UsercontrolHeight / Screen.TwipsPerPixelY) + 1
    'DrawFrameControl hdc, RECT((UsercontrolWidth / Screen.TwipsPerPixelX), (UsercontrolHeight / Screen.TwipsPerPixelY), (UserControl.Width / Screen.TwipsPerPixelX), (UserControl.Height / Screen.TwipsPerPixelY)), DFC_SCROLL, DFCS_SCROLLSIZEGRIP
End Sub

Public Sub Paint()

    CanvasValidate False
    
    'Dim tmpStr As String
    Dim cur As POINTAPI
    Dim tmpSel As RangeType
    cur.X = pOffsetX + LineColumnWidth
    cur.Y = pOffsetY
    
    'Debug.Print VisibleText
    
    'tmpStr = pText
    If pSel.StartPos <> pSel.StopPos Then
        If pSel.StartPos > pSel.StopPos Then
            tmpSel.StartPos = pSel.StopPos
            tmpSel.StopPos = pSel.StartPos
        Else
            tmpSel.StartPos = pSel.StartPos
            tmpSel.StopPos = pSel.StopPos
        End If
    End If

    'UserControl.Cls
    pBackBuffer.DrawCls UsercontrolWidth + LineColumnWidth, UsercontrolHeight
    'UserControl.Line (1, 1)-(UsercontrolWidth, UsercontrolHeight), pBackcolor, BF
    
    If Enabled Then
    
        If ((pSel.StartPos <> pSel.StopPos) And (hasFocus Xor ((Not hasFocus) And (Not pHideSelection)))) Then
            If tmpSel.StartPos > 0 Then
                cur = ClipPrintTextBlock(cur.X, cur.Y, Convert(pText.Partial(0, tmpSel.StartPos)), , False)
            End If
            If (tmpSel.StopPos - tmpSel.StartPos) > 0 Then
                cur = ClipPrintTextBlock(cur.X, cur.Y, Convert(pText.Partial(tmpSel.StartPos, (tmpSel.StopPos - tmpSel.StartPos))), GetSysColor(COLOR_HIGHLIGHT), True)
            End If
            If pText.Length - tmpSel.StopPos > 0 Then
                cur = ClipPrintTextBlock(cur.X, cur.Y, Convert(pText.Partial(tmpSel.StopPos)), , False)
            End If
    
        ElseIf pText.Length > 0 Then
            ClipPrintTextBlock cur.X, cur.Y, Convert(pText.Partial), , False
        End If
    ElseIf pText.Length > 0 Then
        ClipPrintTextBlock cur.X, cur.Y, Convert(pText.Partial), GetSysColor(COLOR_GRAYTEXT), False
    End If
    
    If pLineNumbers Then
        pBackBuffer.DrawLine 0, 0, LineColumnWidth / Screen.TwipsPerPixelX, UsercontrolHeight, GetSysColor(COLOR_SCROLLBAR), BF

        Dim cnt As Long
        For cnt = LineFirstVisible To ((UsercontrolHeight \ TextHeight) + 1) + LineFirstVisible
            pBackBuffer.DrawText (LineColumnWidth - (TextWidth((cnt + 1) & "."))) / Screen.TwipsPerPixelX, ((((cnt - LineFirstVisible) * TextHeight) / Screen.TwipsPerPixelY) + 1), (cnt + 1) & ".", GetSysColor(COLOR_GRAYTEXT)
        Next
        
    End If
    
    RaiseEventSelChange

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Cancel = True
    
    Backcolor = PropBag.ReadProperty("Backcolor", GetSysColor(COLOR_WINDOW))
    Forecolor = PropBag.ReadProperty("Forecolor", GetSysColor(COLOR_WINDOWTEXT))
    ScrollToCaret = PropBag.ReadProperty("ScrollToCaret", True)
    HideSelection = PropBag.ReadProperty("HideSelection", True)
    UserControl.Font.name = PropBag.ReadProperty("Fontname", "Lucida Console")
    UserControl.Font.Size = PropBag.ReadProperty("Fontsize", 9)
    Set pBackBuffer.Font = UserControl.Font
    Text = PropBag.ReadProperty("Text", "")
    MultipleLines = PropBag.ReadProperty("MultipleLines", True)
    ScrollBars = PropBag.ReadProperty("ScrollBars", vbScrollBars.Both)
    AutoRedraw = PropBag.ReadProperty("AutoRedraw", True)
    Locked = PropBag.ReadProperty("Locked", False)
    Enabled = PropBag.ReadProperty("Enabled", True)
    TabSpace = PropBag.ReadProperty("TabSpace", "    ")
    UndoStack = PropBag.ReadProperty("UndoStack", 150)
    UndoDelay = PropBag.ReadProperty("UndoDelay", 1000)
    LineNumbers = PropBag.ReadProperty("LineNumbers", True)
    
    Cancel = False
End Sub

Private Sub UserControl_Resize()
    SetScrollBars
    CanvasValidate
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

Friend Sub SetScrollBars()

    Dim I As Integer
    For I = 0 To 1
        If ((CanvasHeight > UsercontrolHeight And UsercontrolHeight > ScrollBar2.Height) And (pScrollBars = vbScrollBars.Auto)) Or ((pScrollBars = vbScrollBars.Both) Or (pScrollBars = vbScrollBars.Vertical)) Then
            If Not ScrollBar1.Visible Then ScrollBar1.Visible = True
        ElseIf ScrollBar1.Visible Then
            ScrollBar1.Visible = False
        End If
        If ScrollBar1.Visible And CanvasHeight < UsercontrolHeight Then
            If ScrollBar1.Enabled Then ScrollBar1.Enabled = False
        ElseIf ScrollBar1.Visible And CanvasHeight >= UsercontrolHeight Then
            If Not ScrollBar1.Enabled Then ScrollBar1.Enabled = Enabled
        Else
            ScrollBar1.Enabled = Enabled
        End If

        If ((CanvasWidth > UsercontrolWidth And UsercontrolWidth > ScrollBar1.Width) And (pScrollBars = vbScrollBars.Auto)) Or ((pScrollBars = vbScrollBars.Both) Or (pScrollBars = vbScrollBars.Horizontal)) Then
            If Not ScrollBar2.Visible Then ScrollBar2.Visible = True
        ElseIf ScrollBar2.Visible Then
            ScrollBar2.Visible = False
        End If
        If ScrollBar2.Visible And CanvasWidth < UsercontrolWidth Then
            If ScrollBar2.Enabled Then ScrollBar2.Enabled = False
        ElseIf ScrollBar2.Visible And CanvasWidth >= UsercontrolWidth Then
            If Not ScrollBar2.Enabled Then ScrollBar2.Enabled = Enabled
        Else
            ScrollBar2.Enabled = Enabled
        End If

        If ScrollBar1.Visible Then
            ScrollBar1.max = (CanvasHeight - UsercontrolHeight)
            ScrollBar1.SmallChange = TextHeight
            ScrollBar1.LargeChange = ScrollBar1.SmallChange * 4
            If ScrollBar1.Top <> 0 Then ScrollBar1.Top = 0
            If ScrollBar1.Width <> Screen.TwipsPerPixelX ^ 2 Then ScrollBar1.Width = Screen.TwipsPerPixelX ^ 2
            If ScrollBar1.Left <> UsercontrolWidth + LineColumnWidth Then ScrollBar1.Left = UsercontrolWidth + LineColumnWidth
            If ScrollBar1.Height <> UsercontrolHeight Then ScrollBar1.Height = UsercontrolHeight
        End If
        If ScrollBar2.Visible Then
            ScrollBar2.max = (CanvasWidth - UsercontrolWidth)
            ScrollBar2.SmallChange = TextWidth
            ScrollBar2.LargeChange = ScrollBar2.SmallChange * 4
            If ScrollBar2.Left <> 0 Then ScrollBar2.Left = 0
            If ScrollBar2.Height <> Screen.TwipsPerPixelY ^ 2 Then ScrollBar2.Height = Screen.TwipsPerPixelY ^ 2
            If ScrollBar2.Top <> UsercontrolHeight Then ScrollBar2.Top = UsercontrolHeight
            If ScrollBar2.Width <> UsercontrolWidth + LineColumnWidth Then ScrollBar2.Width = UsercontrolWidth + LineColumnWidth
        End If
    Next
    
End Sub

Private Sub UserControl_Show()
    CanvasValidate
    SetScrollBars
    UserControl_Paint
End Sub

Private Sub UserControl_Terminate()
    Unhook Me
'    Set pForecolors = Nothing
'    Set pBackcolors = Nothing
    ResetUndoRedo
    Set xUndoText(0) = Nothing
    Set pText = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Backcolor", pBackcolor, GetSysColor(COLOR_WINDOW)
    PropBag.WriteProperty "Forecolor", pForecolor, GetSysColor(COLOR_WINDOWTEXT)
    PropBag.WriteProperty "ScrollToCaret", pScrollToCaret, True
    PropBag.WriteProperty "HideSelection", pHideSelection, True
    PropBag.WriteProperty "MultipleLines", pMultiLine, True
    PropBag.WriteProperty "Fontname", UserControl.Font.name, "Lucida Console"
    PropBag.WriteProperty "Fontsize", UserControl.Font.Size, 9
    If pText.Length > 0 Then
        PropBag.WriteProperty "Text", Convert(pText.Partial), ""
    Else
        PropBag.WriteProperty "Text", "", ""
    End If
    PropBag.WriteProperty "ScrollBars", pScrollBars, vbScrollBars.Both
    PropBag.WriteProperty "AutoRedraw", AutoRedraw, True
    PropBag.WriteProperty "Locked", Locked, False
    PropBag.WriteProperty "Enabled", Enabled, True
    PropBag.WriteProperty "TabSpace", TabSpace, "    "
    PropBag.WriteProperty "UndoStack", UndoStack, 150
    PropBag.WriteProperty "UndoDelay", UndoDelay, 1000
    PropBag.WriteProperty "LineNumbers", LineNumbers, True
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
Attribute CountWord.VB_Description = "Counts how many times Word appears in Text, optionally specifying the Exact parameter to false for case insensitive matching."
    Dim cnt As Long
    cnt = UBound(Split(Text, Word, , IIf(Exact, vbBinaryCompare, vbTextCompare)))
    If cnt > 0 Then CountWord = cnt
End Function

