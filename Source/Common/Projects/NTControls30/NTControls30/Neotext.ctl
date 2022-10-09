VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.UserControl Neotext 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipControls    =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "Neotext.ctx":0000
   Begin NTControls30.ScrollBar HScroll1 
      Height          =   255
      Left            =   330
      Top             =   2805
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   450
   End
   Begin NTControls30.ScrollBar VScroll1 
      Height          =   1140
      Left            =   3780
      Top             =   360
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2011
      Orientation     =   0
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2175
      Left            =   570
      TabIndex        =   0
      Top             =   555
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   3836
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      Appearance      =   0
      RightMargin     =   2.00000e5
      TextRTF         =   $"Neotext.ctx":0312
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   1920
      Left            =   -60
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   3387
      _Version        =   393217
      BackColor       =   -2147483648
      BorderStyle     =   0
      Enabled         =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"Neotext.ctx":0395
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
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
         Caption         =   "C&ut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuDash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIndent 
         Caption         =   "Indent &Tab"
      End
      Begin VB.Menu mnuUnindent 
         Caption         =   "Un-Indent Ta&b"
      End
   End
End
Attribute VB_Name = "Neotext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private xRTFTable1 As String 'full text cache of concurrent to coloring table in rtf
Private xRTFTable2 As String

Private xVisRange As RangeType
Private xPreRange As RangeType
Private xPstRange As RangeType

Private xCancel As Integer

Private xUndoBuffer As Long 'this holds the current count of onhand undo/redo actions
'Private xUndoSels() As RangeType 'holds undo selections summed short to text edits
'Private xUndoText() As String 'this holds the actual active save info for undo/redo
Private xUndoActs() As UndoType
Private xUndoStage As Long 'this points to the current undo/redo act in the buffer info
'if a undo or redo is preformed this only moves when it happens again equal to last move
'stage-1 for undo, and stage+1 for redo, when if first occurnace, then subs is moved in
'the opposing direction, such that subs at -2 or 1 (also typing, erasing redo if undone)
'is the repreat tracking of only selection changed in that opposing direction lags by 1
Private xUndoSubs As Integer 'this sub projects -2,-1,0,1
'where as 0=stage of texts + subs = state of selelections
'1 = current most recent selection (equal to idle or stage+1)
'0 = to a recent undo, the stage, to a recent redo, stage+1
'-1 = to a recent redo, the state-1, to a recent undo, stage
Private xUndoTemp1 As String 'a currency of text information that is set to be ready for addundo
Private xUndoTemp2 As String 'a adaptive of text information that is set to be ready for addundo
'the follow are exposed properties for setting the undo ability
Private xUndoStack As Long ' -1 for infinite, 0 for no undo, or positive number to limit undo
Private xUndoDirty As Boolean 'value indicating the undo state it can be set false to clear it

Private xRedrawSel As RangeType 'where the current selection is held at all states or set
Private xRedrawVScroll As Long 'used to store the vertical scroll information in redraw pause
Private xRedrawHScroll As Long 'used to store the horizonal scroll information in redraw pause
Private xRedrawRange As RangeType
Private xRedrawSizes As RangeType
Private xCountLength As RangeType

Private xPriorSelHolds As RangeType
Private xNowSizeHolds As RangeType
Private xNowSelHolds As RangeType
Private xPriorSizeHolds As RangeType

Private xTextLines As Strands

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
Public Event OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLEDragOver(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLESetData(Data As RichTextLib.DataObject, DataFormat As Integer)
Public Event OLEStartDrag(Data As RichTextLib.DataObject, AllowedEffects As Long)
Public Event SelChange()

Private pLeftMargin As Boolean
Private pLineNumbers As Boolean

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

Friend Property Get PriorLength() As Long
    PriorLength = xPriorSizeHolds.StartPos
End Property
Friend Property Let PriorLength(ByVal RHS As Long)
    xPriorSizeHolds.StartPos = RHS
End Property
Friend Property Get PriorLineCount() As Long
    PriorLineCount = xPriorSizeHolds.StopPos
End Property
Friend Property Let PriorLineCount(ByVal RHS As Long)
    xPriorSizeHolds.StopPos = RHS
End Property
Friend Property Get PriorSelection() As RangeType
    PriorSelection = xPriorSelHolds
End Property
Friend Property Let PriorSelection(ByRef RHS As RangeType)
    xPriorSelHolds = RHS
End Property
Friend Property Get NowLength() As Long
    NowLength = xNowSizeHolds.StartPos
End Property
Friend Property Let NowLength(ByVal RHS As Long)
    xNowSizeHolds.StartPos = RHS
End Property
Friend Property Get NowLineCount() As Long
    NowLineCount = xNowSizeHolds.StopPos
End Property
Friend Property Let NowLineCount(ByVal RHS As Long)
    xNowSizeHolds.StopPos = RHS
End Property
Friend Property Get NowSelection() As RangeType
    NowSelection = xNowSelHolds
End Property
Friend Property Let NowSelection(ByRef RHS As RangeType)
    xNowSelHolds = RHS
End Property


Private Sub RTFCreateTables()
    
    
    xRTFTable1 = "{\rtf1\ansi\deff0{\fonttbl{\f0\fnil\fcharset0 " & RichTextBox1.Font.Name & ";}}"
        
    xRTFTable2 = xRTFTable1
   
    Dim Red As Long
    Dim Green As Long
    Dim Blue As Long
    
    xRTFTable1 = xRTFTable1 & "{\colortbl"
    ConvertColor UserControl.Forecolor, Red, Green, Blue
    xRTFTable1 = xRTFTable1 & "\red" & Trim(str(Red)) & "\green" & Trim(str(Green)) & "\blue" & Trim(str(Blue)) & ";}" & _
                     "\viewkind4\uc1\pard\sll-" & Font.Weight & "\slmult1" & GetLineHeader

    xRTFTable2 = xRTFTable2 & "{\colortbl"
    ConvertColor SystemColorConstants.vb3DShadow, Red, Green, Blue
    xRTFTable2 = xRTFTable2 & "\red" & Trim(str(Red)) & "\green" & Trim(str(Green)) & "\blue" & Trim(str(Blue)) & ";"
    ConvertColor SystemColorConstants.vbScrollBars, Red, Green, Blue
    xRTFTable2 = xRTFTable2 & "\red" & Trim(str(Red)) & "\green" & Trim(str(Green)) & "\blue" & Trim(str(Blue)) & ";"
    xRTFTable2 = xRTFTable2 & "}" & _
                     "\viewkind4\uc1\pard\sll-" & Font.Weight & "\slmult1" & GetLineHeader
                     
End Sub

Friend Function GetLineHeader() As String
    If UserControl.Font.Bold Then GetLineHeader = "\b"
    If UserControl.Font.Strikethrough Then GetLineHeader = GetLineHeader + "\strike"
    If UserControl.Font.Underline Then GetLineHeader = GetLineHeader + "\ul"
    GetLineHeader = "\plain\f0" + GetLineHeader + "\fs" & TextWeight & " "
End Function

Public Function ConvertColor(ByVal Color As Variant, Optional ByRef Red As Long, Optional ByRef Green As Long, Optional ByRef Blue As Long) As Long
On Error GoTo catcH
    Dim lngColor As Long
    If InStr(CStr(Color), "#") > 0 Then
        GoTo HTMLorHexColor
    ElseIf InStr(CStr(Color), "&H") > 0 Then
        GoTo SysOrLongColor
    ElseIf IsAlphaNumeric(Color) Then
        If (Not (Len(Color) = 6)) And (Not Left(Color, 1) = "0") Then
            GoTo SysOrLongColor
        Else
            GoTo HTMLorHexColor
        End If
    End If
SysOrLongColor:
    lngColor = CLng(Color)
    If Not (lngColor >= 0 And lngColor <= 16777215) Then 'if system colour
        lngColor = lngColor And Not &H80000000
        lngColor = GetSysColor(lngColor)
    End If
    Color = Right("000000" & Hex(lngColor), 6)
    Blue = CByte("&h" & Mid(Color, 5, 2))
    Green = CByte("&h" & Mid(Color, 3, 2))
    Red = CByte("&h" & Mid(Color, 1, 2))
    ConvertColor = RGB(Red, Green, Blue)
    Exit Function
HTMLorHexColor:
    Red = Val("&H" & Mid(Color, 2, 2))
    Green = Val("&H" & Mid(Color, 4, 2))
    Blue = Val("&H" & Right(Color, 2))
    ConvertColor = RGB(Red, Green, Blue)
    Exit Function
catcH:
    Err.Clear
    ConvertColor = 0
End Function



Public Property Get Backcolor() As OLE_COLOR
    Backcolor = RichTextBox1.Backcolor
End Property
Public Property Let Backcolor(ByVal RHS As OLE_COLOR)
    RichTextBox1.Backcolor = RHS
    RTFCreateTables
End Property

Public Property Get Forecolor() As OLE_COLOR
    Forecolor = UserControl.Forecolor
End Property
Public Property Let Forecolor(ByVal RHS As OLE_COLOR)
    UserControl.Forecolor = RHS
    RTFCreateTables
End Property

Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByRef newVal As StdFont)
    Cancel = True
    Set UserControl.Font = newVal
    Set RichTextBox1.Font = newVal
    Set RichTextBox2.Font = newVal
    UserControl.FontBold = UserControl.Font.Bold
    UserControl.FontItalic = UserControl.Font.Italic
    UserControl.FontName = UserControl.Font.Name
    UserControl.FontSize = UserControl.Font.size
    UserControl.FontStrikethru = UserControl.Font.Strikethrough
    UserControl.FontUnderline = UserControl.Font.Underline
        
    Cancel = False
    FindVisRange
    RTFCreateTables

End Property


Friend Property Get TextWeight() As Long ' _
Returns the height of a text string as it would be printed in the current font.
Attribute TextWeight.VB_Description = "Returns the height of a text string as it would be printed in the current font."
    Static setVal As Single
    If setVal = 0 Then setVal = PixelPerPoint
    TextWeight = (RichTextBox1.Font.size * setVal) - 1 'least spacing
End Property

Public Sub SelectAll()
    RichTextBox1.SelStart = 0
    RichTextBox1.SelLength = Length
End Sub

Private Sub HScroll1_KeyDown(KeyCode As Integer, Shift As Integer)
    RichTextBox1_KeyDown KeyCode, Shift
End Sub

Private Sub HScroll1_KeyPress(KeyAscii As Integer)
    RichTextBox1_KeyPress KeyAscii
End Sub

Private Sub HScroll1_KeyUp(KeyCode As Integer, Shift As Integer)
    RichTextBox1_KeyUp KeyCode, Shift
End Sub

Private Sub HScroll1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, HScroll1.Left + X, HScroll1.Top + Y)
End Sub

Private Sub HScroll1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, HScroll1.Left + X, HScroll1.Top + Y)
End Sub

Private Sub HScroll1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, HScroll1.Left + X, HScroll1.Top + Y)
End Sub

Private Sub mnuCopy_Click()
    Copy
End Sub

Private Sub mnuCut_Click()
    Cut
End Sub

Private Sub mnuDelete_Click()
    If Not Locked Then
        If RichTextBox1.SelLength = 0 Then RichTextBox1.SelLength = 1
        RichTextBox1.SelText = ""
    End If

End Sub

Private Sub mnuEdit_Click()
    On Error Resume Next
    RefreshEditMenu
    If Err Then Err.Clear
    On Error GoTo 0
End Sub

Private Sub mnuIndent_Click()
    Indenting RichTextBox1.SelStart, RichTextBox1.SelLength, Chr(9), True
End Sub

Private Sub mnuPaste_Click()
    Paste
End Sub

Private Sub mnuSelectAll_Click()
    SelectAll
End Sub

Private Sub mnuUnindent_Click()
    Indenting RichTextBox1.SelStart, RichTextBox1.SelLength, Chr(9) & Chr(8), True
End Sub

Private Sub RichTextBox1_Click()
    RaiseEvent Click
End Sub

Private Sub RichTextBox1_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub RichTextBox1_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim getChar As RangeType
    xUndoTemp1 = RichTextBox1.SelText
    If KeyCode = 8 Or KeyCode = 46 Then
        If KeyCode = 46 Then 'right directed single
            xUndoTemp2 = Chr(0) 'key form deleting select
            xRedrawSel.StartPos = RichTextBox1.SelStart
            xRedrawSel.StopPos = RichTextBox1.SelStart + 1
            'next set descriptions to inform the action: disapear the
            'text to the right of the character motion before it moves
        ElseIf KeyCode = 8 Then 'key form of movement
            xUndoTemp2 = Chr(8) 'to the left backspace
            xRedrawSel.StartPos = RichTextBox1.SelStart - 1
            xRedrawSel.StopPos = RichTextBox1.SelStart
            'next set descriptions to inform the action: disapear the
            'text to the right of the character IN motion so is swaps
        End If
    Else
        xUndoTemp2 = Chr(KeyCode)
        xRedrawSel.StartPos = RichTextBox1.SelStart
        xRedrawSel.StopPos = RichTextBox1.SelStart + 1
    End If
    If RichTextBox1.SelLength > 0 Then
        xRedrawRange.StartPos = RichTextBox1.SelStart
        xRedrawRange.StopPos = RichTextBox1.SelStart + RichTextBox1.SelLength
    End If
    
    RaiseEvent KeyDown(KeyCode, Shift)
    If KeyCode <> 0 Then
        
        Dim lIndex As Long
        Dim txt As String
        Dim temp As Long
        Select Case KeyCode

            Case 93 'menu
                Dim pt As POINTAPI
                GetCursorPos pt
                PopupMenu mnuEdit, , 0, 0
            Case 35 'end
                Cancel = True
                lIndex = SendMessageLngPtr(RichTextBox1.hwnd, EM_LINEINDEX, RichTextBox1.GetLineFromChar(RichTextBox1.SelStart), 0&)
                txt = GetLineText(RichTextBox1.GetLineFromChar(RichTextBox1.SelStart))
                temp = (lIndex + Len(txt))
                If Shift = 1 Then
                    If RichTextBox1.SelLength = 0 Then
                        temp = temp - RichTextBox1.SelStart
                        If temp < 0 Then temp = 0
                        RichTextBox1.SelLength = temp
                    Else
                        RichTextBox1.SelStart = RichTextBox1.SelStart + RichTextBox1.SelLength
                        RichTextBox1.SelLength = 0
                    End If
                ElseIf Shift = 0 Then
                    RichTextBox1.SelStart = temp
                End If
                Cancel = False
                SetScrollbarsByVisibility Me
                KeyCode = 0
            Case 36 'home
                Cancel = True
                lIndex = SendMessageLngPtr(RichTextBox1.hwnd, EM_LINEINDEX, RichTextBox1.GetLineFromChar(RichTextBox1.SelStart), 0&)
                txt = GetTextRange(RichTextBox1.hwnd, lIndex, RichTextBox1.SelStart)
                If Not ((Replace(Replace(txt, Chr(9), ""), " ", "") = "") And (Len(txt) > 0)) Then
                    txt = GetLineText(RichTextBox1.GetLineFromChar(RichTextBox1.SelStart))
                    Do While (Left(txt, 1) = Chr(9) Or Left(txt, 1) = " ")
                        txt = Mid(txt, 2)
                        lIndex = lIndex + 1
                    Loop
                End If
                If Shift = 1 Then
                    temp = RichTextBox1.SelLength + (RichTextBox1.SelStart - lIndex)
                    RichTextBox1.SelStart = lIndex
                    If temp < 0 Then temp = 0
                    RichTextBox1.SelLength = temp
                ElseIf Shift = 0 Then
                    RichTextBox1.SelStart = lIndex
                End If
                Cancel = False
                SetScrollbarsByVisibility Me, False
                KeyCode = 0
            Case 9 'tab
                Cancel = True
                If RichTextBox1.SelLength > 0 Then
                
                    lIndex = SendMessageLngPtr(RichTextBox1.hwnd, EM_LINEINDEX, RichTextBox1.GetLineFromChar(RichTextBox1.SelStart), 0&)
                    temp = SendMessageLngPtr(RichTextBox1.hwnd, EM_LINEINDEX, RichTextBox1.GetLineFromChar(RichTextBox1.SelStart + (RichTextBox1.SelLength - 1)), 0&)
                    If temp > lIndex Then
    
                        If Shift = 0 Then
                            Indenting RichTextBox1.SelStart, RichTextBox1.SelLength, Chr(9), True
                        ElseIf Shift = 1 Then
                            Indenting RichTextBox1.SelStart, RichTextBox1.SelLength, Chr(9) & Chr(8), True
                        End If
        
                        KeyCode = 0
                        SetScrollbarsByVisibility Me, False
                    End If
                End If
                Cancel = False
            Case Else
                If KeyCode <> 16 And KeyCode <> 17 And KeyCode <> 18 Then
                    SetScrollbarsByVisibility Me
                End If
        End Select
    
    
        If KeyCode > 0 Then
    
            If (Shift = 2) And (KeyCode = 69) Then
                KeyCode = 0
            End If
            
            If (Shift = 2) And (KeyCode = 82) Then
                KeyCode = 0
                Redo
            End If
        
            If (Shift = 2) And (KeyCode = 90) Then
                KeyCode = 0
                Undo
            End If
            
            If Shift = 2 And KeyCode = 88 Then
                KeyCode = 0
                Cut
            End If
            
            If KeyCode = 45 And Shift = 1 Then
                KeyCode = 0
                Paste
            End If
        
            If KeyCode = 45 And Shift = 2 And (SelLength > 0) Then
                KeyCode = 0
                Copy
            End If
        
            If KeyCode = 46 And (SelLength > 0) Then
                KeyCode = 0
                Clear
            End If
    
        End If
    End If
End Sub

Public Sub Indenting(ByVal SelStart As Long, ByVal SelLength As Long, Optional ByVal CharStr As String = "", Optional ByVal SelectAfter As Boolean = False)
    'i.e. selecting text and using tab to indent, or commenting and uncommenting selected text, default is to remove
    'any tab by setting CharSet to Chr(9) & Chr(8), SelectAfter forces the range edited to be selected when done
    Dim lIndex As Long
    Dim txt As String
    Dim temp As Long
    Dim eIndex As Long
    
    If CharStr = "" Then CharStr = Chr(9) & Chr(8)

    lIndex = SendMessageLngPtr(RichTextBox1.hwnd, EM_LINEINDEX, RichTextBox1.GetLineFromChar(SelStart), 0&)
    eIndex = SendMessageLngPtr(RichTextBox1.hwnd, EM_LINEINDEX, RichTextBox1.GetLineFromChar(SelStart + (SelLength - 1)), 0&)
    eIndex = (eIndex + SendMessageLngPtr(RichTextBox1.hwnd, EM_LINELENGTH, eIndex, 0&)) - lIndex

    For temp = RichTextBox1.GetLineFromChar(SelStart) To RichTextBox1.GetLineFromChar(SelStart + (SelLength - 1))

        txt = GetLineText(temp)
        If InStr(CharStr, Chr(8)) = 0 Then
            txt = CharStr & txt
            eIndex = eIndex + Len(CharStr)
        Else
            If Left(txt, Len(Replace(CharStr, Chr(8), ""))) = Replace(CharStr, Chr(8), "") Then
                txt = Mid(txt, Len(Replace(CharStr, Chr(8), "")) + 1)
                eIndex = eIndex - Len(Replace(CharStr, Chr(8), ""))
            End If
        End If
        SetLineText temp, txt

    Next

    If SelectAfter Then
        SetScrollbarsByVisibility Me, False
        RichTextBox1.SelStart = lIndex
        RichTextBox1.SelLength = eIndex
    End If

End Sub

Private Sub RichTextBox1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)

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
    
        If KeyAscii = 9 Then
            If (SelLength > 0) And (InStr(SelText, vbCrLf) > 0) Then
                KeyAscii = 0
            End If
        End If

        If Not KeyAscii = 0 Then
            If xUndoTemp2 = UCase(Chr(KeyAscii)) Then
                xUndoTemp2 = Chr(KeyAscii)
            End If
            If RichTextBox1.SelLength = 0 Then
                xUndoTemp1 = ""
            End If
        End If
        
    End If
    
End Sub

Private Sub RichTextBox1_KeyUp(KeyCode As Integer, Shift As Integer)
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

Private Sub RichTextBox1_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub RichTextBox1_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Unhook
    
    Cancel = True
    
    Dim pt As POINTAPI
    
    pt.X = (X / Screen.TwipsPerPixelX)
    pt.Y = (Y / Screen.TwipsPerPixelY)
    RichTextBox1.SelStart = SendMessageStruct(RichTextBox1.hwnd, EM_CHARFROMPOS, 0&, pt)
    RichTextBox1.SelLength = 0
    
    RichTextBox1.SelText = Data.GetData(ClipBoardConstants.vbCFText)
        
    Cancel = False
    
    FindVisRange
'    If xRefresh Then RefreshStateProc 1
    
    Hook
    
    
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub RichTextBox1_OLEDragOver(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Dim pt As POINTAPI
    
    pt.X = (X / Screen.TwipsPerPixelX)
    pt.Y = (Y / Screen.TwipsPerPixelY)
    RichTextBox1.SelStart = SendMessageStruct(RichTextBox1.hwnd, EM_CHARFROMPOS, 0&, pt)
    
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub RichTextBox1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub RichTextBox1_OLESetData(Data As RichTextLib.DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub RichTextBox1_OLEStartDrag(Data As RichTextLib.DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub RichTextBox2_KeyDown(KeyCode As Integer, Shift As Integer)
    RichTextBox1_KeyDown KeyCode, Shift
End Sub

Private Sub RichTextBox2_KeyPress(KeyAscii As Integer)
    RichTextBox1_KeyPress KeyAscii
End Sub

Private Sub RichTextBox2_KeyUp(KeyCode As Integer, Shift As Integer)
    RichTextBox1_KeyUp KeyCode, Shift
End Sub

Private Sub RichTextBox2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, RichTextBox2.Left + X, RichTextBox2.Top + Y)
End Sub

Private Sub RichTextBox2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, RichTextBox2.Left + X, RichTextBox2.Top + Y)
End Sub

Private Sub RichTextBox2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, RichTextBox2.Left + X, RichTextBox2.Top + Y)
End Sub

Private Sub UserControl_InitProperties()
    Cancel = True
    RichTextBox1.Font.Name = "Lucida Console"
    RichTextBox1.Font.Italic = False
    RichTextBox1.Font.Bold = False
    RichTextBox1.Font.Strikethrough = False
    RichTextBox1.Font.Underline = False
    RichTextBox1.Font.size = 8
    RichTextBox1.Font.Weight = 400
    RichTextBox1.Font.Charset = 0
    RichTextBox1.Backcolor = vbWhite
    Set Font = RichTextBox1.Font
    Set RichTextBox2.Font = RichTextBox1.Font
    UserControl.FontBold = RichTextBox1.Font.Bold
    UserControl.FontItalic = RichTextBox1.Font.Italic
    UserControl.FontName = RichTextBox1.Font.Name
    UserControl.FontSize = RichTextBox1.Font.size
    UserControl.FontStrikethru = RichTextBox1.Font.Strikethrough
    UserControl.FontUnderline = RichTextBox1.Font.Underline
    UserControl.Forecolor = vbBlack
    RichTextBox2.Backcolor = SystemColorConstants.vbScrollBars
    RichTextBox2.Visible = pLeftMargin Or pLineNumbers
    Cancel = False
    UserControl_Paint
End Sub
Public Sub SelectRow(ByVal CursorRow As Long, Optional ByVal FullRowSelect As Boolean = True)
    Cancel = True
    
    Dim lIndex As Long
    lIndex = SendMessageLngPtr(RichTextBox1.hwnd, EM_LINEINDEX, CursorRow, 0&)
    If LineNumber = CursorRow And Not FullRowSelect Then
        RichTextBox1.SelStart = RichTextBox1.SelStart
    Else
        Dim lLength As Long
        If FullRowSelect Then lLength = SendMessageLngPtr(RichTextBox1.hwnd, EM_LINELENGTH, lIndex, 0&)
        RichTextBox1.SelStart = lIndex
        RichTextBox1.SelLength = lLength
    End If
    
    Cancel = False
End Sub
Public Sub SelectRows(ByVal StartRow As Long, Optional ByVal LastRow As Long = 0)

    Dim lIndex As Long
    Dim lLength As Long

    Cancel = True
    VScroll1.AutoRedraw = False
    HScroll1.AutoRedraw = False
    
    Dim pV As POINTAPI
    Dim pH As POINTAPI
    SendMessageStruct RichTextBox1.hwnd, EM_GETSCROLLPOS, 0, pV
    SendMessageStruct RichTextBox1.hwnd, EM_GETSCROLLPOS, 0, pH
    
    If LastRow > StartRow Then
        lIndex = SendMessageLngPtr(RichTextBox1.hwnd, EM_LINEINDEX, StartRow, 0&)
        Dim tmp As Long
        For tmp = StartRow To LastRow
            lLength = lLength + SendMessageLngPtr(RichTextBox1.hwnd, EM_LINELENGTH, SendMessageLngPtr(RichTextBox1.hwnd, EM_LINEINDEX, tmp, 0&), 0&) + 2
        Next
        If lIndex < 0 Then lIndex = 0
        RichTextBox1.SelStart = lIndex
        RichTextBox1.SelLength = lLength
    Else
        lIndex = SendMessageLngPtr(RichTextBox1.hwnd, EM_LINEINDEX, StartRow, 0&)
        lLength = SendMessageLngPtr(RichTextBox1.hwnd, EM_LINELENGTH, lIndex, 0&)
        If lIndex < 0 Then lIndex = 0
        If RichTextBox1.SelStart = lIndex And RichTextBox1.SelLength = lLength Then
            RichTextBox1.SelStart = lIndex
            RichTextBox1.SelLength = 0
        Else
            RichTextBox1.SelStart = lIndex
            RichTextBox1.SelLength = lLength
        End If
    End If

    SendMessageStruct RichTextBox1.hwnd, EM_SETSCROLLPOS, 0, pV
    SendMessageStruct RichTextBox1.hwnd, EM_SETSCROLLPOS, 0, pH

    VScroll1.AutoRedraw = True
    HScroll1.AutoRedraw = True
    Cancel = False
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SelectionLines 0, Button, Shift, X, Y
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SelectionLines 1, Button, Shift, X, Y
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SelectionLines 2, Button, Shift, X, Y
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
    DrawFrameControl hdc, Rect(((UserControl.Width / Screen.TwipsPerPixelX) - 15), ((UserControl.Height / Screen.TwipsPerPixelY) - 15), (UserControl.Width / Screen.TwipsPerPixelX), (UserControl.Height / Screen.TwipsPerPixelY)), DFC_SCROLL, DFCS_SCROLLSIZEGRIP
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Cancel = True
    
    Enabled = PropBag.ReadProperty("Enabled", True)
    Locked = PropBag.ReadProperty("Locked", False)
    pLineNumbers = PropBag.ReadProperty("LineNumbers", False)
    pLeftMargin = PropBag.ReadProperty("LeftMargin", True)

    Font.Name = PropBag.ReadProperty("FontName", "Lucida Console")
    Font.Italic = PropBag.ReadProperty("FontItalic", False)
    Font.Bold = PropBag.ReadProperty("FontBold", False)
    Font.Strikethrough = PropBag.ReadProperty("FontStrikethrough", False)
    Font.Underline = PropBag.ReadProperty("FontUnderline", False)
    Font.size = PropBag.ReadProperty("FontSize", 8)
    Font.Weight = PropBag.ReadProperty("FontWeight", 400)
    Font.Charset = PropBag.ReadProperty("FontCharset", 0)
    
    UserControl.Forecolor = PropBag.ReadProperty("ForeColor", vbBlack)
    RichTextBox1.Backcolor = PropBag.ReadProperty("BackColor", vbWhite)
    
    Cancel = False
    UserControl_Paint
    
End Sub

Private Sub UserControl_Show()
    UserControl_Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Enabled", Enabled, True
    PropBag.WriteProperty "Locked", Locked, False
    PropBag.WriteProperty "LineNumbers", pLineNumbers, False
    PropBag.WriteProperty "LeftMargin", pLeftMargin, True

    PropBag.WriteProperty "FontName", Font.Name, "Lucida Console"
    PropBag.WriteProperty "FontItalic", Font.Italic, False
    PropBag.WriteProperty "FontBold", Font.Bold, False
    PropBag.WriteProperty "FontStrikethrough", Font.Strikethrough, False
    PropBag.WriteProperty "FontUnderline", Font.Underline, False
    PropBag.WriteProperty "FontSize", Font.size, 8
    PropBag.WriteProperty "FontWeight", Font.Weight, 400
    PropBag.WriteProperty "FontCharset", Font.Charset, 0
    
    PropBag.WriteProperty "ForeColor", UserControl.Forecolor, vbBlack
    PropBag.WriteProperty "BackColor", RichTextBox1.Backcolor, vbWhite
End Sub

Public Property Get LineNumbers() As Boolean
    LineNumbers = pLineNumbers
End Property
Public Property Let LineNumbers(ByVal RHS As Boolean)
    pLineNumbers = RHS
    If pLineNumbers Then LeftMargin = True
    UserControl_Resize
End Property

Public Property Get LineNumber() As Long
    LineNumber = RichTextBox1.GetLineFromChar(RichTextBox1.SelStart)
End Property

Public Property Get LeftMargin() As Boolean
    LeftMargin = pLeftMargin
End Property
Public Property Let LeftMargin(ByVal RHS As Boolean)
    pLeftMargin = RHS
    If pLeftMargin Then RichTextBox2.Visible = True
    UserControl_Resize
End Property
Public Property Get Height() As Long
    Height = UserControl.Height
End Property
Public Property Let Height(ByVal newVal As Long)
    UserControl.Height = newVal
End Property

Public Property Get Width() As Long
    Width = UserControl.Width
End Property
Public Property Let Width(ByVal newVal As Long)
    UserControl.Width = newVal
End Property

Public Property Get Locked() As Boolean
    Locked = RichTextBox1.Locked
End Property
Public Property Let Locked(ByVal RHS As Boolean)
    RichTextBox1.Locked = RHS
End Property
Public Property Get Enabled() As Boolean
    Enabled = RichTextBox1.Enabled
End Property
Public Property Let Enabled(ByVal RHS As Boolean)
    RichTextBox1.Enabled = RHS
    HScroll1.Enabled = RHS
    VScroll1.Enabled = RHS
End Property

Public Property Get TextWidth(ByVal Text As String) As Long
    TextWidth = UserControl.TextWidth(Text)
End Property

Public Property Get TextHeight(ByVal Text As String) As Long
    TextHeight = UserControl.TextHeight(Text)
End Property

Friend Property Get OldProc() As Long
    OldProc = CLng(RichTextBox1.Tag)
End Property


Friend Property Get RichTextBox() As RichTextBox
    Set RichTextBox = RichTextBox1
End Property

Friend Property Get ColumnTextBox() As RichTextBox
    Set ColumnTextBox = RichTextBox2
End Property

Friend Property Get HScroll() As ScrollBar
    Set HScroll = HScroll1
End Property

Friend Property Get VScroll() As ScrollBar
    Set VScroll = VScroll1
End Property

Public Property Get hwnd() As Long
    hwnd = RichTextBox1.hwnd
End Property


'Public Property Let Text(ByVal RHS As String)
'
'
'    RichTextBox1.TextRTF = xRTFTable1 & Replace(Replace(RHS, vbCrLf, vbLf), vbLf, "\line ") & "}"
'
'
'    UserControl_Resize
'End Property

Public Property Get Text()
    Text = RichTextBox1.Text
End Property
Public Property Let Text(ByRef RHS)
    Cancel = True
    Select Case TypeName(RHS)
        Case "String"

            RichTextBox1.TextRTF = xRTFTable1 & Replace(Replace(RHS, vbCrLf, vbLf), vbLf, "\line ") & "}"

    End Select

    If Not Locked Then ResetUndoRedo
    Cancel = False
    
    UserControl_Resize

End Property
Public Property Set Text(ByRef RHS)
    Cancel = True
    Select Case TypeName(RHS)
        Case "Strand", "NTControls30.Strand"
            'RichTextBox1.TextRTF = xRTFTable1 & Replace(Replace(RHS.GetString(), vbCrLf, vbLf), vbLf, "\line ") & "}"
            RichTextBox1.TextRTF = xRTFTable1 & Replace(Replace(RHS.GetString(), vbCrLf, vbLf), vbLf, "\line ") & "}"
            'xTextLines.Value = RHS.GetString()
        Case "Strands", "NTNodes10.Strands"
            RichTextBox1.TextRTF = xRTFTable1 & Replace(Replace(RHS.Value, vbCrLf, vbLf), vbLf, "\line ") & "}"
            'Set xTextLines = Nothing
            'Set xTextLines = RHS
    End Select

    If Locked Then ResetUndoRedo
    Cancel = False
    
    UserControl_Resize
End Property


Public Property Get SelText() As String
    SelText = RichTextBox1.SelText
End Property
Public Property Let SelText(ByVal RHS As String)
    Cancel = True
    RichTextBox1.SelText = RHS
    'xTextLines.Value = nArrayOfString(RHS)
    ResetUndoRedo
    Cancel = False
    UserControl_Resize
End Property

Public Property Get SelStart() As Long
    SelStart = RichTextBox1.SelStart
End Property
Public Property Let SelStart(ByVal newVal As Long)
    RichTextBox1.SelStart = newVal
End Property

Public Property Get SelLength() As Long
    SelLength = RichTextBox1.SelLength
End Property
Public Property Let SelLength(ByVal newVal As Long)
    RichTextBox1.SelLength = newVal
End Property

'Public Property Get TextRTF() As String
'    TextRTF = RichTextBox1.TextRTF
'End Property

Public Sub Reset()
    Cancel = True
    UserControl_InitProperties
    RichTextBox1.Text = ""
    'xTextLines.Behavior = ScopeNormal
    Cancel = False
    
    UserControl_Resize
    FindVisRange
    RTFCreateTables

End Sub

'Public Property Let ErrorLine(ByVal RHS As Long)
'
'    xErrorLine = RHS
'    FindVisRange
'    If xRefresh Then RefreshStateProc 1
'
'End Property
'
'Public Property Get ErrorLine() As Long
'    ErrorLine = xErrorLine
'End Property

Private Sub RichTextBox1_Change()

    Dim setScroll As Boolean
    
    If Not Cancel Then
        Cancel = True

        xCountLength.StartPos = SendMessageLngPtr(RichTextBox1.hwnd, EM_GETLINECOUNT, 0&, 0&)
        xCountLength.StopPos = Length
    
        Dim xRedrawLines As RangeType
        xRedrawLines.StopPos = xRedrawSizes.StopPos
        xRedrawLines.StartPos = xCountLength.StartPos
        xRedrawSizes.StopPos = xRedrawSizes.StartPos
        xRedrawSizes.StartPos = xCountLength.StopPos
        
        Dim xTmp As RangeType
        
       ' SendMessageStruct RichTextBox1.hWnd, EM_EXGETSEL, 0, xRedrawSel
        SendMessageStruct RichTextBox1.hwnd, EM_EXGETSEL, 0, xTmp
        
        PriorLength = xRedrawSizes.StopPos
        'Debug.Print "Prior Char Size: " & xRedrawSizes.StopPos;
       ' NowLength = xRedrawSizes.StartPos
        'Debug.Print " Now Char Size: " & xRedrawSizes.StartPos
        PriorLineCount = xRedrawLines.StopPos
        'Debug.Print " Prior Line Count: " & xRedrawLines.StopPos;
       ' NowLineCount = xRedrawLines.StartPos
        'Debug.Print "Now Line Count: " & xRedrawLines.StartPos
        
'        If xDirty = Matching Then
'            xDirty = TextChanged
'            xDirtyDate = Now
'        End If
        
        AddUndo
        
        'condition in stops here for values
        '###################################
        xRedrawRange.StartPos = xTmp.StartPos
        xRedrawRange.StopPos = xTmp.StopPos
        xRedrawSizes.StartPos = xCountLength.StopPos
        xRedrawSizes.StopPos = xCountLength.StartPos
        
        
        Cancel = False
        FindVisRange
        SetScrollBarsMax
        RaiseEvent Change
    End If
    
    
End Sub

Friend Sub SetScrollBarsMax(Optional ByVal LinesOnly As Boolean = False)
    Cancel = True
    
    
    VScroll1.AutoRedraw = False
    HScroll1.AutoRedraw = False
    
    Dim lCount As Long
    Dim temp As Long
    Dim numStr As String
    lCount = SendMessage(RichTextBox1.hwnd, EM_GETLINECOUNT, 0, ByVal 0&)
    If Not LinesOnly Then
        temp = (lCount - (RichTextBox1.Height \ (UserControl.TextHeight("A")))) + _
            ((RichTextBox1.Height \ (UserControl.TextHeight("A"))) \ 4)
        If VScroll1.Max <> temp Then VScroll1.Max = temp
    End If

    If (lCount > 0 And Not LinesOnly) Or LineNumbers Then
        Dim lLine As Long
        lLine = SendMessageLngPtr(RichTextBox.hwnd, EM_GETFIRSTVISIBLELINE, 0, 0&)
        lCount = (RichTextBox1.Height \ (UserControl.TextHeight("A")))
        Dim lLength As Long
        Dim lIndex As Long
        Dim vmax As Long
        If Not LinesOnly Then temp = (UserControl.RichTextBox1.RightMargin / Screen.TwipsPerPixelX) + RichTextBox1.Width
        For lCount = lLine + lCount To lLine Step -1
            If Not LinesOnly Then
                lIndex = SendMessageLngPtr(RichTextBox1.hwnd, EM_LINEINDEX, lCount, 0&)
                lLength = SendMessageLngPtr(RichTextBox1.hwnd, EM_LINELENGTH, lIndex, 0&)
                vmax = (UserControl.TextWidth(GetTextRange(RichTextBox1.hwnd, lIndex, lIndex + lLength)) - RichTextBox1.Width) - (RichTextBox1.Width \ 2)
                If temp < vmax Then temp = vmax
            End If
            If LineNumbers Then numStr = "\cf0 " & ((lLine + lCount + 1) - VScroll1.Value) & "\cf1 ." & vbCrLf & numStr
        Next
        If temp <> HScroll1.Max And Not LinesOnly Then HScroll1.Max = temp
    ElseIf Not LinesOnly Then
        HScroll1.Max = 0
    End If
    
    If LineNumbers Then
        RichTextBox2.TextRTF = xRTFTable2 & Replace(Replace(numStr, vbCrLf, vbLf), vbLf, "\qr\line " & vbCrLf) & "}"
       ' RichTextBox2.Width = UserControl.TextWidth(numStr) * Screen.TwipsPerPixelX
        
    ElseIf LeftMargin Then
        If Not RichTextBox2.Text = "" Then RichTextBox2.Text = ""
    End If

    VScroll1.AutoRedraw = True
    HScroll1.AutoRedraw = True
    
    Cancel = False
End Sub

Private Sub SelectionLines(ByVal MouseEvent As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Static SelectRange As POINTAPI
    Static SelectLineFlag As Long
    Static SelectJumpToStart As Long
    
    Select Case MouseEvent
        Case 0

            If Button = 1 Then
                SelectLineFlag = 1
                SelectJumpToStart = SendMessageLngPtr(RichTextBox.hwnd, EM_GETFIRSTVISIBLELINE, 0, 0&)
                SelectRange.X = (Y \ (UserControl.TextHeight("A")))
                
                SelectRows SelectJumpToStart + SelectRange.X
            ElseIf (X > 0 And Y > 0 And X < RichTextBox2.Width And Y < RichTextBox2.Height) Then
                SelectLineFlag = 0
                SelectRange.X = 0
                SelectRange.Y = 0
            End If
            
        Case 1
            
            If Button = 1 And (SelectLineFlag > 0) Then

                If (X > 0 And Y > 0 And X < RichTextBox2.Width And Y < RichTextBox2.Height) Then
                    SelectLineFlag = 2
                    SelectRange.Y = (Y \ (UserControl.TextHeight("A")))
                    If SelectRange.Y <> SelectRange.X Then
                        
                        If SelectRange.Y < SelectRange.X Then
                            SelectRows SendMessageLngPtr(RichTextBox.hwnd, EM_GETFIRSTVISIBLELINE, 0, 0&) + SelectRange.Y, SelectJumpToStart + SelectRange.X
                        Else
                            SelectRows SelectJumpToStart + SelectRange.X, SendMessageLngPtr(RichTextBox.hwnd, EM_GETFIRSTVISIBLELINE, 0, 0&) + SelectRange.Y
                        End If
                    End If
                    
                End If
                                
            ElseIf (X > 0 And Y > 0 And X < RichTextBox2.Width And Y < RichTextBox2.Height) Then
                SelectLineFlag = 0
                SelectRange.X = 0
                SelectRange.Y = 0
            End If
            
            SetScrollbarsByVisibility Me
                        
        Case 2

            SelectLineFlag = 0
            SelectRange.X = 0
            SelectRange.Y = 0
            SetScrollbarsByVisibility Me
    End Select
End Sub

Private Sub SelectionDrag(ByVal MouseEvent As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Static SelectRange As POINTAPI
    Static SelectDragFlag As Long
    
    Select Case MouseEvent
        Case 0
            If Button = 1 Then
                SelectDragFlag = 1
                SelectRange.X = RichTextBox1.SelStart
                SelectRange.Y = RichTextBox1.SelStart + RichTextBox1.SelLength
            Else
                SelectDragFlag = 0
            End If
        Case 1
            If Button = 1 And (SelectDragFlag = 1) Then
                SelectDragFlag = 2

                Dim rct As Rect
                Dim mse As POINTAPI
                Dim per As Single
                
                Do While GetAsyncKeyState%(VK_LBUTTON)
                    
                    GetWindowRect RichTextBox1.hwnd, rct
                    GetCursorPos mse
                    If mse.X < rct.Left Then
                        AdjustScollbarValue HScroll1, -IIf((rct.Left - mse.X) > (RichTextBox1.Width \ 2) \ Screen.TwipsPerPixelX, HScroll1.LargeChange, HScroll1.SmallChange)
                        SetScrollbarsByVisibility Me, (SelectRange.Y = RichTextBox1.SelStart + RichTextBox1.SelLength)
                    ElseIf mse.X > rct.Right Then
                        per = ((mse.X - rct.Right) / (RichTextBox1.Width \ 2)) * (HScroll1.Max - HScroll1.Value)
                        
                        AdjustScollbarValue HScroll1, -per + HScroll1.Value
                        
                        SetScrollbarsByVisibility Me, (SelectRange.Y = RichTextBox1.SelStart + RichTextBox1.SelLength)
                        
                    End If
                    
                    If LineNumbers Then SetScrollBarsMax True
        
                    DoEvents
                    Sleep 50
                Loop

                SendMessageStruct RichTextBox1.hwnd, EM_GETSCROLLPOS, 0, mse
        
                AdjustScollbarValue HScroll1, -HScroll1.Value + (mse.X * Screen.TwipsPerPixelX)
                
                If LineNumbers Then SetScrollBarsMax True
                
                SelectDragFlag = 0
                
            Else
                SelectDragFlag = 0
            End If
        Case 3
            SelectDragFlag = 0
            SetScrollbarsByVisibility Me
    End Select
End Sub

Private Sub RichTextBox1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        RefreshEditMenu
        PopupMenu mnuEdit
    End If
    SelectionDrag 0, Button, Shift, X, Y
    RaiseEvent MouseDown(Button, Shift, RichTextBox1.Left + X, RichTextBox1.Top + Y)
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

Private Sub RichTextBox1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SelectionDrag 1, Button, Shift, X, Y
    RaiseEvent MouseMove(Button, Shift, RichTextBox1.Left + X, RichTextBox1.Top + Y)
End Sub

Private Sub RichTextBox1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SelectionDrag 2, Button, Shift, X, Y
    RaiseEvent MouseUp(Button, Shift, RichTextBox1.Left + X, RichTextBox1.Top + Y)
End Sub

Private Sub RichTextBox1_SelChange()

    If Not Cancel Then
        
        Dim xLineSel As RangeType
        Dim xLineVis As RangeType
        xLineSel.StartPos = xRedrawSizes.StartPos
        xLineSel.StopPos = xRedrawSizes.StopPos
        xLineVis.StartPos = xRedrawRange.StartPos
        xLineVis.StopPos = xRedrawRange.StopPos
        'backed up we only change ranges here
    
        'proceed like chagne event with mods
        Dim xRedrawLines As RangeType
        xRedrawLines.StopPos = xRedrawSizes.StopPos
        xRedrawLines.StartPos = xCountLength.StartPos
        xRedrawSizes.StopPos = xRedrawSizes.StartPos
        xRedrawSizes.StartPos = xCountLength.StopPos
    
        Cancel = True

        xUndoActs(0).AfterSelRange.StartPos = xRedrawSel.StartPos
        xUndoActs(0).AfterSelRange.StopPos = xRedrawSel.StopPos
        
        SendMessageStruct RichTextBox1.hwnd, EM_EXGETSEL, 0, xRedrawSel
        
        'only what we want chaged by this event at this point is properly met
        PriorSelection = xRedrawRange
        'Debug.Print "Prior Selection: Start: " & xRedrawRange.StartPos & "  Stop: " & xRedrawRange.StopPos
        'NowSelection = xRedrawSel
        'Debug.Print "Now Selection: Start: " & xRedrawSel.StartPos & "  Stop: " & xRedrawSel.StopPos

        FindVisRange
        
        xUndoActs(0).PriorSelRange.StartPos = xRedrawRange.StartPos
        xUndoActs(0).PriorSelRange.StopPos = xRedrawRange.StopPos
        
        Cancel = False
        
        'condition in stops here for values
        '###################################
        xRedrawRange.StartPos = xRedrawSel.StartPos
        xRedrawRange.StopPos = xRedrawSel.StopPos
        xRedrawSizes.StartPos = xCountLength.StopPos
        xRedrawSizes.StopPos = xCountLength.StartPos
    
        'restore from the backed up what not change
        xRedrawSizes.StartPos = xLineSel.StartPos
        xRedrawSizes.StopPos = xLineSel.StopPos
        xRedrawSel.StartPos = xLineVis.StartPos
        xRedrawSel.StopPos = xLineVis.StopPos

    End If
        
    SetScrollbarsByVisibility Me
    RaiseEvent SelChange
End Sub

Private Sub UserControl_Resize()
    RichTextBox1.Top = 0

    If UserControl.Width - VScroll1.Width > 0 Then RichTextBox1.Width = UserControl.Width - VScroll1.Width
    If UserControl.Height - HScroll1.Height > 0 Then RichTextBox1.Height = UserControl.Height - HScroll1.Height
   
    VScroll1.Top = 0
    VScroll1.Left = RichTextBox1.Width
    VScroll1.Height = RichTextBox1.Height

    HScroll1.Left = 0
    HScroll1.Top = RichTextBox1.Height
    HScroll1.Width = RichTextBox1.Width
    
    If LeftMargin Or LineNumbers Then
        RichTextBox2.Visible = True
        RichTextBox2.Left = 0
        RichTextBox2.Top = 0
        RichTextBox2.Height = RichTextBox1.Height
        If LineNumbers Then
            If SendMessage(RichTextBox1.hwnd, EM_GETLINECOUNT, 0, ByVal 0&) > ((RichTextBox1.Height \ UserControl.TextHeight("A")) + 1) Then
                RichTextBox2.Width = UserControl.TextWidth("." & CStr(SendMessage(RichTextBox1.hwnd, EM_GETLINECOUNT, 0, ByVal 0&)) & ".")
            Else
                RichTextBox2.Width = UserControl.TextWidth("." & CStr(((RichTextBox1.Height \ UserControl.TextHeight("A")) + 1)) & ".")
            End If
        Else
            RichTextBox2.Width = UserControl.TextWidth("..") - (Screen.TwipsPerPixelX * 4)
        End If
        
        RichTextBox1.Left = RichTextBox2.Width
        RichTextBox1.Width = RichTextBox1.Width - RichTextBox2.Width
        
    Else
        RichTextBox1.Left = 0
    End If
    If (RichTextBox1.Width \ 2) > 0 Then HScroll1.LargeChange = (RichTextBox1.Width \ 2)
    SetScrollBarsMax
    
    FindVisRange
End Sub


Private Sub VScroll1_Change()
    If GetActiveWindow <> RichTextBox1.hwnd Then
        If RichTextBox1.Visible Then RichTextBox1.SetFocus
    End If

    VScroll1_Scroll
    SetScrollBarsMax
End Sub

Private Sub VScroll1_KeyDown(KeyCode As Integer, Shift As Integer)
    RichTextBox1_KeyDown KeyCode, Shift
End Sub

Private Sub VScroll1_KeyPress(KeyAscii As Integer)
    RichTextBox1_KeyPress KeyAscii
End Sub

Private Sub VScroll1_KeyUp(KeyCode As Integer, Shift As Integer)
    RichTextBox1_KeyUp KeyCode, Shift
End Sub

Private Sub VScroll1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, VScroll1.Left + X, VScroll1.Top + Y)
End Sub

Private Sub VScroll1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, VScroll1.Left + X, VScroll1.Top + Y)
End Sub

Private Sub VScroll1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, VScroll1.Left + X, VScroll1.Top + Y)
End Sub

Friend Sub VScroll1_Scroll()

    Dim p As POINTAPI
    SendMessageStruct RichTextBox1.hwnd, EM_GETSCROLLPOS, 0, p
    p.Y = VScroll1.Value * ((UserControl.TextHeight("A")) / Screen.TwipsPerPixelY)
    SendMessageStruct RichTextBox1.hwnd, EM_SETSCROLLPOS, 0, p
    If LineNumbers Then SetScrollBarsMax True

End Sub

Private Sub HScroll1_Change()
    If GetActiveWindow <> RichTextBox1.hwnd Then
        If RichTextBox1.Visible Then RichTextBox1.SetFocus
    End If
    HScroll1_Scroll
    SetScrollBarsMax
End Sub

Friend Sub HScroll1_Scroll()
    Dim p As POINTAPI
    SendMessageStruct RichTextBox1.hwnd, EM_GETSCROLLPOS, 0, p
    p.X = (HScroll1.Value / Screen.TwipsPerPixelX)
    SendMessageStruct RichTextBox1.hwnd, EM_SETSCROLLPOS, 0, p
    If LineNumbers Then SetScrollBarsMax True
    
End Sub

Private Sub UserControl_Initialize()

    Cancel = True
    
    Set xTextLines = New Strands
        
    xUndoStack = -1
    ResetUndoRedo
    
    Set Font = RichTextBox1.Font

    SendMessageLngPtr RichTextBox1.hwnd, EM_SETUNDOLIMIT, 0, ByVal 0&
    SendMessageLngPtr RichTextBox2.hwnd, EM_SETUNDOLIMIT, 0, ByVal 0&

    SendMessageLngPtr RichTextBox1.hwnd, EM_SETTEXTMODE, TM_SINGLECODEPAGE + TM_RICHTEXT, ByVal 0&
    SendMessageLngPtr RichTextBox2.hwnd, EM_SETTEXTMODE, TM_SINGLECODEPAGE + TM_RICHTEXT, ByVal 0&

    RTFCreateTables
    RichTextBox1.TextRTF = xRTFTable1 & "}"
    RichTextBox2.TextRTF = xRTFTable2 & "}"

'    RichTextBox1.SelTabCount = 4
'    RichTextBox2.SelTabCount = 4
    
    HScroll1.SmallChange = UserControl.TextWidth("A")
    
    Cancel = False

    SetScrollBarsMax
    
    Hook
End Sub
Private Sub Hook()
    If Not IsRunningMode Then
        SubClassed.Add Me, "H" & RichTextBox1.hwnd
        RichTextBox1.Tag = GetWindowLong(RichTextBox1.hwnd, GWL_WNDPROC)
        SetWindowLong RichTextBox1.hwnd, GWL_WNDPROC, AddressOf WinProc
        'SetTimer RichTextBox1.hwnd, ObjPtr(Me), xInterval, AddressOf FileProc
    End If
End Sub
Private Sub Unhook()
    If IsNumeric(RichTextBox1.Tag) Then
        SetWindowLong RichTextBox1.hwnd, GWL_WNDPROC, CLng(RichTextBox1.Tag)
        SubClassed.Remove "H" & RichTextBox1.hwnd
    End If
End Sub

Private Sub UserControl_Terminate()

    Unhook
    
    Set xTextLines = Nothing

End Sub
Friend Function FindVisRange(Optional ByVal SetProfiles As Boolean = True)
    Cancel = True
    xVisRange = VisibleRange
    xPreRange = Range(0, xVisRange.StartPos)
    xPstRange = Range(xVisRange.StopPos, Length)
    Cancel = False
End Function

Friend Function VisibleRange() As RangeType
    Dim pt As POINTAPI
    pt.X = 1
    pt.Y = 1
    VisibleRange.StartPos = SendMessageStruct(RichTextBox1.hwnd, EM_CHARFROMPOS, 0&, pt)
    pt.X = (RichTextBox1.Width \ Screen.TwipsPerPixelX)
    pt.Y = (RichTextBox1.Height \ Screen.TwipsPerPixelY)
    VisibleRange.StopPos = SendMessageStruct(RichTextBox1.hwnd, EM_CHARFROMPOS, 0&, pt)
End Function

Public Function Length(Optional LineNumber As Long = 0) As Long
    Dim gt As GetTextLengthEx
    If LineNumber = 0 Then
        gt.Flags = GTL_USECRLF Or GTL_NUMCHARS
        Length = SendMessageStruct(RichTextBox1.hwnd, EM_GETTEXTLENGTHEX, VarPtr(gt), ByVal 0&)
    Else
        Length = SendMessageLngPtr(RichTextBox1.hwnd, EM_LINEINDEX, LineNumber - 1, 0&)
        Length = SendMessageLngPtr(RichTextBox1.hwnd, EM_LINELENGTH, Length, 0&)
    End If
End Function

Public Function CountWord(ByVal Text As String, ByVal Word As String, Optional ByVal Exact As Boolean = True) As Long
    Dim cnt As Long
    cnt = UBound(Split(Text, Word, , IIf(Exact, vbBinaryCompare, vbTextCompare)))
    If cnt > 0 Then CountWord = cnt
End Function

Public Function WordCount(ByVal Text As String, Optional ByVal TheSeperator As String = " ", Optional ByVal Exact As Boolean = True) As Long
    Dim cnt As Long
    cnt = UBound(Split(Text, TheSeperator, , IIf(Exact, vbBinaryCompare, vbTextCompare)))
    If cnt >= 0 Then WordCount = cnt
End Function

'Friend Function SetTextRTF(ByVal newText As String, ByRef nRange As RangeType, Optional ByVal SkipDisable As Boolean = False)
'
'    If Not Cancel Then
'        Dim st As SetTextEx
'        st.Flags = ST_SELECTION
'
'        If Not SkipDisable Then DisableRedraw
'        SendMessageLngPtr RichTextBox1.hwnd, EM_SETTEXTMODE, TM_MULTICODEPAGE + TM_RICHTEXT, ByVal 0&
'        SendMessageStruct RichTextBox1.hwnd, EM_EXSETSEL, 0, nRange
'        SendMessageString RichTextBox1.hwnd, EM_SETTEXTEX, VarPtr(st), newText
'        SendMessageStruct RichTextBox1.hwnd, EM_EXSETSEL, 0, xRedrawSel
'        If Not SkipDisable Then EnableRedraw
'
'    End If
'
'End Function
'Friend Function SetText(ByVal newText As String, ByRef nRange As RangeType, Optional ByVal SkipDisable As Boolean = False)
'
'    If Not Cancel Then
'        Dim st As SetTextEx
'        st.Flags = ST_SELECTION
'
'        If Not SkipDisable Then DisableRedraw
'        SendMessageLngPtr RichTextBox1.hwnd, EM_SETTEXTMODE, TM_MULTICODEPAGE + TM_PLAINTEXT, ByVal 0&
'        SendMessageStruct RichTextBox1.hwnd, EM_EXSETSEL, 0, nRange
'
'        SendMessageString RichTextBox1.hwnd, EM_SETTEXTEX, VarPtr(st), newText
'        SendMessageStruct RichTextBox1.hwnd, EM_EXSETSEL, 0, xRedrawSel
'        If Not SkipDisable Then EnableRedraw
'
'    End If
'
'End Function


'Friend Function TimerProc() As Boolean
'    KillTimer RichTextBox1.hwnd, ObjPtr(Me)
'
'    RefreshView
'
'    SetTimer RichTextBox1.hwnd, ObjPtr(Me), xInterval, AddressOf FileProc
'End Function
'
'Private Sub YieldLatency()
'    Static yLatency As Single
'    If yLatency = 0 Then
'        yLatency = Timer
'    Else
'        If Timer - yLatency >= 0.25 Then
'            yLatency = Timer
'            DoEvents
'        End If
'    End If
'End Sub

Private Function GetLineText(ByVal Index As Long) As String
    Dim xLine As RangeType
    Dim st As SetTextEx
    st.Flags = ST_SELECTION And ST_NEWCHARS
    xLine.StartPos = SendMessageLngPtr(RichTextBox1.hwnd, EM_LINEINDEX, Index, 0&)
    If xLine.StartPos = -1 Then Exit Function
    xLine.StopPos = xLine.StartPos + SendMessageLngPtr(RichTextBox1.hwnd, EM_LINELENGTH, xLine.StartPos, 0&)
    GetLineText = Replace(Replace(GetTextRange(RichTextBox1.hwnd, xLine.StartPos, xLine.StopPos), vbCr, ""), vbLf, "")

End Function

Private Sub SetLineText(ByVal Index As Long, ByVal Text As String)
    Dim xLine As RangeType
    Dim st As SetTextEx
    st.Flags = ST_SELECTION And ST_NEWCHARS
    xLine.StartPos = SendMessageLngPtr(RichTextBox1.hwnd, EM_LINEINDEX, Index, 0&)
    If xLine.StartPos = -1 Then Exit Sub
    xLine.StopPos = xLine.StartPos + SendMessageLngPtr(RichTextBox1.hwnd, EM_LINELENGTH, xLine.StartPos, 0&)
    
    SendMessageStruct RichTextBox1.hwnd, EM_EXSETSEL, 0, xLine
    SendMessageString RichTextBox1.hwnd, EM_SETTEXTEX, ByVal VarPtr(st), Text
    
End Sub


Public Sub Cut()
    If Not Locked Then
        If RangeLength(xRedrawSel) > 0 Then
            xUndoTemp1 = RichTextBox1.SelText
            xUndoTemp2 = ""
            
            SendMessageLngPtr RichTextBox1.hwnd, WM_CUT, 0, ByVal 0&
        End If
    End If
End Sub
Public Sub Copy()
    SendMessageLngPtr RichTextBox1.hwnd, WM_COPY, 0, ByVal 0&
End Sub
Public Sub Paste()
    If Not Locked Then
        xUndoTemp1 = RichTextBox1.SelText
        xUndoTemp2 = Clipboard.GetText(vbCFText)

        If RichTextBox1.SelLength > 0 Then
            xRedrawRange.StartPos = RichTextBox1.SelStart
            xRedrawRange.StopPos = RichTextBox1.SelStart + RichTextBox1.SelLength
        End If
        If Len(xUndoTemp2) > 1 Then
            xRedrawSel.StartPos = RichTextBox1.SelStart
            xRedrawSel.StopPos = RichTextBox1.SelStart + Len(xUndoTemp2)
        End If
        
        Cancel = True
        
       
        SendMessageLngPtr RichTextBox1.hwnd, WM_PASTE, 0, ByVal 0&
        
        Cancel = False
    End If
End Sub
Public Sub Clear()
    If Not Locked Then
        xUndoTemp1 = RichTextBox1.SelText
        xUndoTemp2 = Chr(8)
        SendMessageLngPtr RichTextBox1.hwnd, WM_CLEAR, 0, ByVal 0&
    End If
End Sub
Public Property Get UndoDirty() As Boolean
    UndoDirty = xUndoDirty
End Property
Public Property Let UndoDirty(ByVal newVal As Boolean)
    xUndoDirty = newVal
    ResetUndoRedo
End Property
Public Property Get UndoStack() As Long
    UndoStack = xUndoStack
End Property
Public Property Let UndoStack(ByVal newVal As Long)
    xUndoStack = newVal
    ResetUndoRedo
End Property

Private Sub ResetUndoRedo()

'    ReDim xUndoSels(0 To 1) As RangeType
'    ReDim xUndoText(0 To 1) As String
'    xUndoSels(1).StartPos = 0
'    xUndoSels(1).StopPos = 0
'    xUndoText(1) = ""
    
    
    ReDim xUndoActs(0 To 1) As UndoType
    xUndoActs(0).PriorSelRange.StartPos = 0
    xUndoActs(0).PriorSelRange.StopPos = 0
    xUndoActs(0).PriorTextData = ""

    xUndoStage = 1
    xUndoDirty = False
    xUndoBuffer = xUndoStack
    xUndoSubs = 0
    
End Sub

Private Sub AddUndo()
        
    If (xUndoStack <> 0) Then
    
        Dim cnt As Long
        If (UBound(xUndoActs) = xUndoBuffer) And Not (xUndoStack = -1) Then
            For cnt = (LBound(xUndoActs) + 2) To UBound(xUndoActs)
                xUndoActs(cnt - 1) = xUndoActs(cnt)
            Next
        ElseIf (UBound(xUndoActs) < xUndoBuffer) Or (xUndoStack = -1) Then
            If xUndoSubs = 1 Or xUndoSubs = 2 Then xUndoStage = xUndoStage + 1
            ReDim Preserve xUndoActs(0 To xUndoStage) As UndoType
        End If

        xUndoActs(xUndoStage).PriorSelRange.StartPos = xUndoActs(0).PriorSelRange.StartPos
        xUndoActs(xUndoStage).PriorSelRange.StopPos = xUndoActs(0).PriorSelRange.StopPos
        
        xUndoActs(xUndoStage).AfterSelRange.StartPos = xUndoActs(0).AfterSelRange.StartPos
        xUndoActs(xUndoStage).AfterSelRange.StopPos = xUndoActs(0).AfterSelRange.StopPos
        
        xUndoActs(xUndoStage).PriorTextData = xUndoTemp1
        xUndoActs(xUndoStage).AfterTextData = xUndoTemp2

'        Debug.Print "Entry: Prior (" & xUndoActs(xUndoStage).PriorSelRange.StartPos & ", " & xUndoActs(xUndoStage).PriorSelRange.StopPos & ")"
'        Debug.Print "     : After (" & xUndoActs(xUndoStage).AfterSelRange.StartPos & ", " & xUndoActs(xUndoStage).AfterSelRange.StopPos & ")"
'        Debug.Print "     : Knitt (" & xUndoActs(xUndoStage).PriorTextData & ", " & xUndoActs(xUndoStage).AfterTextData & ")"

        If (Not xUndoSubs = 1) Then xUndoStage = xUndoStage + 1

        xUndoSubs = 0
        xUndoDirty = False
        
    ElseIf xUndoDirty Then
        ResetUndoRedo
    End If

End Sub

Public Function CanUndo() As Boolean
    CanUndo = ((UBound(xUndoActs) > 0) And (xUndoStage > 1)) And (Not Locked)
End Function
Public Function CanRedo() As Boolean
    CanRedo = ((xUndoStage < UBound(xUndoActs)) And (UBound(xUndoActs) > 0)) And (Not Locked)
End Function

Public Sub Undo()
    If CanUndo Then
        Cancel = True
        
        If xUndoStage > UBound(xUndoActs) Then xUndoStage = xUndoStage - 1
        If xUndoStage > 0 Then xUndoStage = xUndoStage + xUndoSubs
        xUndoSubs = xUndoSubs - 1
        If xUndoSubs < -1 Then xUndoSubs = -1
        If xUndoStage <= 1 Then xUndoSubs = 0
        
'        Debug.Print "Undo : Stage (" & xUndoStage & ", " & xUndoSubs & ")"
'        Debug.Print "     : Prior (" & xUndoActs(xUndoStage).PriorSelRange.StartPos & ", " & xUndoActs(xUndoStage).PriorSelRange.StopPos & ")"
'        Debug.Print "     : After (" & xUndoActs(xUndoStage).AfterSelRange.StartPos & ", " & xUndoActs(xUndoStage).AfterSelRange.StopPos & ")"
'        Debug.Print "     : Knitt (" & xUndoActs(xUndoStage).PriorTextData & ", " & xUndoActs(xUndoStage).AfterTextData & ")"
'
        SendMessageStruct RichTextBox1.hwnd, EM_EXSETSEL, 0, xUndoActs(xUndoStage).AfterSelRange

        RichTextBox1.SelText = String(Len(xUndoActs(xUndoStage).AfterTextData), Chr(0))
        RichTextBox1.SelStart = xUndoActs(xUndoStage).PriorSelRange.StartPos
        RichTextBox1.SelLength = 0
        RichTextBox1.SelText = xUndoActs(xUndoStage).PriorTextData
    
        SendMessageStruct RichTextBox1.hwnd, EM_EXSETSEL, 0, xUndoActs(xUndoStage).PriorSelRange
        
        Cancel = False
        xUndoDirty = True
        FindVisRange
        
        RaiseEvent Change
    End If
End Sub

Public Sub Redo()
    If CanRedo Then
        
        If xUndoStage < UBound(xUndoActs) Then xUndoStage = xUndoStage + xUndoSubs
        xUndoSubs = xUndoSubs + 1
        If xUndoSubs > 1 Then xUndoSubs = 1
        If xUndoStage >= UBound(xUndoActs) Then xUndoSubs = 0
        If xUndoStage > UBound(xUndoActs) Then xUndoStage = UBound(xUndoActs)
        
'        Debug.Print "Redo : Stage (" & xUndoStage & ", " & xUndoSubs & ")"
'        Debug.Print "     : Prior (" & xUndoActs(xUndoStage).PriorSelRange.StartPos & ", " & xUndoActs(xUndoStage).PriorSelRange.StopPos & ")"
'        Debug.Print "     : After (" & xUndoActs(xUndoStage).AfterSelRange.StartPos & ", " & xUndoActs(xUndoStage).AfterSelRange.StopPos & ")"
'        Debug.Print "     : Knitt (" & xUndoActs(xUndoStage).PriorTextData & ", " & xUndoActs(xUndoStage).AfterTextData & ")"
'
        Cancel = True

        SendMessageStruct RichTextBox1.hwnd, EM_EXSETSEL, 0, xUndoActs(xUndoStage).PriorSelRange
        
        If GetTextRange(RichTextBox1.hwnd, xUndoActs(xUndoStage).PriorSelRange.StartPos, xUndoActs(xUndoStage).PriorSelRange.StopPos) = xUndoActs(xUndoStage).PriorTextData Then
            RichTextBox1.SelText = ""
        ElseIf GetTextRange(RichTextBox1.hwnd, xUndoActs(xUndoStage).AfterSelRange.StartPos, xUndoActs(xUndoStage).AfterSelRange.StopPos) = xUndoActs(xUndoStage).AfterTextData Then
            RichTextBox1.SelLength = 0
        End If
        RichTextBox1.SelText = xUndoActs(xUndoStage).AfterTextData
    
        SendMessageStruct RichTextBox1.hwnd, EM_EXSETSEL, 0, xUndoActs(xUndoStage).AfterSelRange
                
        Cancel = False
        xUndoDirty = True
        
        FindVisRange
        
        RaiseEvent Change
    End If
End Sub
Private Sub mnuRedo_Click()
    Redo
End Sub

Private Sub mnuUndo_Click()
    Undo
End Sub
