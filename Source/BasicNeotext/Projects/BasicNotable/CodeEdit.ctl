VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.UserControl CodeEdit 
   AutoRedraw      =   -1  'True

   ClientHeight    =   2610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4995
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   2610
   ScaleWidth      =   4995
   ToolboxBitmap   =   "CodeEdit.ctx":0000
   Begin RichTextLib.RichTextBox txtScript1 
      CausesValidation=   0   'False
      Height          =   2190
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   3863
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      RightMargin     =   2.00000e5
      OLEDragMode     =   0
      OLEDropMode     =   1
      TextRTF         =   $"CodeEdit.ctx":0312
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
   Begin RichTextLib.RichTextBox txtScript2 
      CausesValidation=   0   'False
      Height          =   2190
      Left            =   1605
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   315
      Visible         =   0   'False
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   3863
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      RightMargin     =   40
      OLEDragMode     =   0
      OLEDropMode     =   1
      TextRTF         =   $"CodeEdit.ctx":0395
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
      Begin VB.Menu mnuComment 
         Caption         =   "Co&mment"
      End
      Begin VB.Menu mnuUncomment 
         Caption         =   "Uncomme&nt"
      End
   End
End
Attribute VB_Name = "CodeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private xRTF As Strand

Private xFileName As String
Private xFileNum As Integer
Private xLatency As Single
Private xUndoDelay As Single

Private Type LineDef
    LineLen As Long
    LineClr As Long
End Type
    
Private Type RTFColorType
    Red As Integer
    Green As Integer
    Blue As Integer
End Type

Private Type RTFColorMapType
    Fore As RTFColorType
    Text As RTFColorType
    Variable As RTFColorType
    Statement As RTFColorType
    Value As RTFColorType
    Operator As RTFColorType
    Comment As RTFColorType
    Error As RTFColorType
    Dream1 As RTFColorType
    Dream2 As RTFColorType
    Dream3 As RTFColorType
    Dream4 As RTFColorType
    Dream5 As RTFColorType
    Dream6 As RTFColorType
End Type

Public Enum ColoringFileMode
    ShowingPriority = 1
    StoringPriority = 2
End Enum

Public Enum ColoringIndexes
    TextIndex = 0
    VariableIndex = 1
    StatementIndex = 2
    ValueIndex = 3
    OperatorIndex = 4
    CommentIndex = 5
    ErrorIndex = 6
    Dream1Index = 7
    Dream2Index = 8
    Dream3Index = 9
    Dream4Index = 10
    Dream5Index = 11
    Dream6Index = 12
End Enum

Private RTFColorMap As RTFColorMapType

Private script_ForeColor As OLE_COLOR
Private script_TextColor As OLE_COLOR
Private script_VariableColor As OLE_COLOR
Private script_StatementColor As OLE_COLOR
Private script_ValueColor As OLE_COLOR
Private script_OperatorColor As OLE_COLOR
Private script_CommentColor As OLE_COLOR
Private script_ErrorColor As OLE_COLOR
Private script_Dream1Color As OLE_COLOR
Private script_Dream2Color As OLE_COLOR
Private script_Dream3Color As OLE_COLOR
Private script_Dream4Color As OLE_COLOR
Private script_Dream5Color As OLE_COLOR
Private script_Dream6Color As OLE_COLOR

Private num_Fore As String
Private num_Text As String
Private num_Variable As String
Private num_Statement As String
Private num_Value As String
Private num_Operator As String
Private num_Comment As String
Private num_Error As String
Private num_Dream1 As String
Private num_Dream2 As String
Private num_Dream3 As String
Private num_Dream4 As String
Private num_Dream5 As String
Private num_Dream6 As String

Private xCancel As Long

Private xRefresh As Boolean
Private xLanguage As String
Private xErrorLine As Long
Private xUndoStack As Long
Private xUndoDirty As Boolean

Private xUndoText() As String
Private xUndoStart() As Long
Private xUndoLength() As Long
Private xUndoStage As Long
Private xUndoBuffer As Long

Private xRedrawSel As RangeType
Private xRedrawVScroll As Long
Private xRedrawHScroll As Long
Private xRedrawRange As RangeType
Private xRedrawSizes As RangeType
Private xCountLength As RangeType
Private xDefined() As LineDef

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
Public Event OLEDragDrop(Data As Object, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLEDragOver(Data As Object, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLESetData(Data As Object, DataFormat As Integer)
Public Event OLEStartDrag(Data As Object, AllowedEffects As Long)
Public Event SelChange()

Public Event QueryDefine(ByVal LineNum As Long, ByRef LineText As String, ByRef ColorIndex As ColoringIndexes)

Private WithEvents txtScript As RichTextBox
Attribute txtScript.VB_VarHelpID = -1

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

Public Property Get Text() As String
    Text = txtScript.Text
End Property
Public Property Let Text(ByVal newVal As String)
    Cancel = True
    txtScript.Text = newVal
    Retext
    ResetUndoRedo
    Cancel = False
    If xRefresh Then RefreshView
End Property

Public Sub Concat(ByVal Data As String)
    Cancel = True
    txtScript.Text = txtScript.Text & Data
    Retext Data
    Cancel = False
    AddUndo
    If xRefresh Then RefreshView
End Sub

Public Property Get SelText() As String
    SelText = txtScript.SelText
End Property
Public Property Let SelText(ByVal newVal As String)
    Cancel = True
    txtScript.SelText = newVal
    Retext txtScript.SelText
    ResetUndoRedo
    Cancel = False
    If xRefresh Then RefreshView
End Property

Public Property Get SelStart() As Long
    SelStart = txtScript.SelStart
End Property
Public Property Let SelStart(ByVal newVal As Long)
    txtScript.SelStart = newVal
    Refresh
End Property

Public Property Get SelLength() As Long
    SelLength = txtScript.SelLength
End Property
Public Property Let SelLength(ByVal newVal As Long)
    txtScript.SelLength = newVal
End Property

Friend Property Get FileNum() As Long
    FileNum = xFileNum
End Property

Public Property Get FileName() As String
    FileName = xFileName
End Property
Public Property Let FileName(ByVal newVal As String)
    CloseFileName newVal

    If (newVal <> "") Then
        xFileNum = FreeFile
        xFileName = newVal
        Open xFileName For Binary Shared As #xFileNum
    End If
End Property

Friend Property Get TextHeight() As Long
    Static pValue As Single
    If pValue = 0 Then pValue = PixelPerPoint
    TextHeight = txtScript.Font.Size * pValue
End Property

Public Property Get Tag()
    If txtScript2.Tag = True Then
        Set Tag = PtrObj(txtScript1.Tag)
    Else
        Tag = txtScript1.Tag
    End If
End Property
Public Property Let Tag(ByVal newVal)
    txtScript1.Tag = newVal
    txtScript2.Tag = False
End Property
Public Property Set Tag(ByRef newVal)
    txtScript1.Tag = ObjPtr(newVal)
    txtScript2.Tag = True
End Property

Public Property Get WordWrap() As Boolean
    WordWrap = txtScript2.Visible
End Property
Public Property Let WordWrap(ByVal newVal As Boolean)
    If ((Not newVal) And (txtScript1.Visible = False)) Or _
         (newVal And (txtScript1.Visible = True)) Then
         
        If xHooked Then HookObj txtScript
        Dim mState As Boolean
        mState = xRefresh
        xRefresh = False
        Cancel = True
        txtScript.Visible = False
        If Not newVal Then
            txtScript1.BackColor = txtScript.BackColor
            txtScript1.Font.Name = txtScript.Font.Name
            txtScript1.Font.Italic = txtScript.Font.Italic
            txtScript1.Font.Bold = txtScript.Font.Bold
            txtScript1.Font.StrikeThrough = txtScript.Font.StrikeThrough
            txtScript1.Font.Underline = txtScript.Font.Underline
            txtScript1.Font.Size = txtScript.Font.Size
            txtScript1.Font.Weight = txtScript.Font.Weight
            txtScript1.Font.Charset = txtScript.Font.Charset
            txtScript1.Enabled = txtScript.Enabled
            txtScript1.TextRTF = txtScript.TextRTF
            txtScript1.SelStart = txtScript.SelStart
            txtScript1.SelLength = txtScript.SelLength
            txtScript1.Locked = txtScript.Locked
            txtScript1.Visible = True
            Set txtScript = Nothing
            Set txtScript = txtScript1
            txtScript2.Text = ""
        Else
            txtScript2.BackColor = txtScript.BackColor
            txtScript2.Font.Name = txtScript.Font.Name
            txtScript2.Font.Italic = txtScript.Font.Italic
            txtScript2.Font.Bold = txtScript.Font.Bold
            txtScript2.Font.StrikeThrough = txtScript.Font.StrikeThrough
            txtScript2.Font.Underline = txtScript.Font.Underline
            txtScript2.Font.Size = txtScript.Font.Size
            txtScript1.Font.Weight = txtScript.Font.Weight
            txtScript1.Font.Charset = txtScript.Font.Charset
            txtScript2.Enabled = txtScript.Enabled
            txtScript2.TextRTF = txtScript.TextRTF
            txtScript2.SelStart = txtScript.SelStart
            txtScript2.SelLength = txtScript.SelLength
            txtScript2.Locked = txtScript.Locked
            txtScript2.Visible = True
            Set txtScript = Nothing
            Set txtScript = txtScript2
            txtScript1.Text = ""
        End If
        xRefresh = mState
        Cancel = False
        If xHooked Then HookObj txtScript
        UserControl_Resize
    End If
    AdjustWordWrap
End Property

Private Sub Retext(Optional ByVal Data As String = "")

    xCountLength.StartPos = SendMessageLngPtr(txtScript.hwnd, EM_GETLINECOUNT, 0&, 0&)
    xCountLength.StopPos = SendMessageLngPtr(txtScript.hwnd, WM_GETTEXTLENGTH, 0&, 0&)

End Sub

Public Sub Redraw()
    Dim mState As Boolean
    mState = xRefresh
    xRefresh = True
    If xRefresh Then RefreshView
    xRefresh = mState
    BackColor = BackColor
End Sub

Friend Sub RefreshView(Optional ByVal Full As Boolean = False)
    Dim tmp As RangeType
    tmp = VisibleRange
   
    WalkRTFView tmp.StartPos, tmp.StopPos
End Sub

Public Property Let LineDefines(Optional ByVal LineNumber As Long = -1, ByVal newVal As ColoringIndexes)
    If xLanguage = "Defined" Then
    
        Dim lines As Long
        Dim dirty As Boolean
        lines = SendMessageLngPtr(txtScript.hwnd, EM_GETLINECOUNT, 0&, 0&)
        
        If LineNumber = -1 Then
        
            ReDim Preserve xDefined(0 To lines) As LineDef
            For lines = 1 To UBound(xDefined)
                If Not xDefined(lines).LineClr = newVal Then
                    xDefined(lines).LineClr = newVal
                    dirty = True
                End If
            Next
            If dirty And xRefresh Then RefreshView
        Else
            If LineNumber > 0 And LineNumber < lines Then
                Dim visRange As RangeType
                visRange = VisibleRange
                visRange.StartPos = GetLineNumber(txtScript.hwnd, visRange.StartPos) + 1
                visRange.StopPos = GetLineNumber(txtScript.hwnd, visRange.StopPos) + 1
        
                ReDim Preserve xDefined(0 To lines) As LineDef
                If xDefined(LineNumber).LineClr <> newVal Then
                    xDefined(LineNumber).LineClr = newVal
                    If ((LineNumber >= visRange.StartPos) And (LineNumber <= visRange.StopPos)) Then dirty = True
                End If
            End If
            If dirty And xRefresh Then RefreshView
        End If
        
    End If
End Property

Public Property Get LineDefines(Optional ByVal LineNumber As Long = -1) As ColoringIndexes
    If xLanguage = "Defined" Then
        If LineNumber = -1 Then LineNumber = Me.LineNumber
        Dim lines As Long
        lines = SendMessageLngPtr(txtScript.hwnd, EM_GETLINECOUNT, 0&, 0&)
        If LineNumber > 0 And LineNumber < lines Then
            ReDim Preserve xDefined(0 To lines) As LineDef
            LineDefines = xDefined(LineNumber).LineClr
        End If
    End If
End Property

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

Public Property Get UndoDelay() As Single
    UndoDelay = xUndoDelay
End Property
Public Property Let UndoDelay(ByVal newVal As Single)
    xUndoDelay = newVal
End Property

Public Property Get Locked() As Boolean
    Locked = txtScript.Locked
End Property
Public Property Let Locked(ByVal newVal As Boolean)
    txtScript.Locked = newVal
End Property

Public Property Get AutoRedraw() As Boolean
    AutoRedraw = xRefresh
End Property
Public Property Let AutoRedraw(ByVal newVal As Boolean)
    xRefresh = newVal
End Property

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

Public Property Get hwnd() As Long
    hwnd = txtScript.hwnd
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = txtScript.BackColor
End Property
Public Property Let BackColor(ByVal newVal As OLE_COLOR)
    Cancel = True
    txtScript.BackColor = newVal
    Cancel = False
    If xRefresh Then RefreshView
End Property

Public Property Get ForeColor() As OLE_COLOR
    ColorText = script_TextColor
End Property
Public Property Let ForeColor(newVal As OLE_COLOR)
    Cancel = True
    script_TextColor = newVal
    Cancel = False
    If xRefresh Then RefreshView
End Property

Public Property Get ColorText() As OLE_COLOR
    ColorText = script_TextColor
End Property
Public Property Let ColorText(newVal As OLE_COLOR)
    script_TextColor = newVal
    If xRefresh Then RefreshView
End Property

Public Property Get ColorVariable() As OLE_COLOR
    ColorVariable = script_VariableColor
End Property
Public Property Let ColorVariable(newVal As OLE_COLOR)
    script_VariableColor = newVal
    If xRefresh Then RefreshView
End Property

Public Property Get ColorStatement() As OLE_COLOR
    ColorStatement = script_StatementColor
End Property
Public Property Let ColorStatement(newVal As OLE_COLOR)
    script_StatementColor = newVal
    If xRefresh Then RefreshView
End Property

Public Property Get ColorValue() As OLE_COLOR
    ColorValue = script_ValueColor
End Property
Public Property Let ColorValue(newVal As OLE_COLOR)
    script_ValueColor = newVal
    If xRefresh Then RefreshView
End Property

Public Property Get ColorOperator() As OLE_COLOR
    ColorOperator = script_OperatorColor
End Property
Public Property Let ColorOperator(newVal As OLE_COLOR)
    script_OperatorColor = newVal
    If xRefresh Then RefreshView
End Property

Public Property Get ColorComment() As OLE_COLOR
    ColorComment = script_CommentColor
End Property
Public Property Let ColorComment(newVal As OLE_COLOR)
    script_CommentColor = newVal
    If xRefresh Then RefreshView
End Property

Public Property Get ColorError() As OLE_COLOR
    ColorError = script_ErrorColor
End Property
Public Property Let ColorError(newVal As OLE_COLOR)
    script_ErrorColor = newVal
    If xRefresh Then RefreshView
End Property

Public Property Get ColorDream1() As OLE_COLOR
    ColorDream1 = script_Dream1Color
End Property
Public Property Let ColorDream1(newVal As OLE_COLOR)
    script_Dream1Color = newVal
    If xRefresh Then RefreshView
End Property

Public Property Get ColorDream2() As OLE_COLOR
    ColorDream2 = script_Dream2Color
End Property
Public Property Let ColorDream2(newVal As OLE_COLOR)
    script_Dream2Color = newVal
    If xRefresh Then RefreshView
End Property

Public Property Get ColorDream3() As OLE_COLOR
    ColorDream3 = script_Dream3Color
End Property
Public Property Let ColorDream3(newVal As OLE_COLOR)
    script_Dream3Color = newVal
    If xRefresh Then RefreshView
End Property

Public Property Get ColorDream4() As OLE_COLOR
    ColorDream4 = script_Dream4Color
End Property
Public Property Let ColorDream4(newVal As OLE_COLOR)
    script_Dream4Color = newVal
    If xRefresh Then RefreshView
End Property

Public Property Get ColorDream5() As OLE_COLOR
    ColorDream5 = script_Dream5Color
End Property
Public Property Let ColorDream5(newVal As OLE_COLOR)
    script_Dream5Color = newVal
    If xRefresh Then RefreshView
End Property

Public Property Get ColorDream6() As OLE_COLOR
    ColorDream6 = script_Dream6Color
End Property
Public Property Let ColorDream6(newVal As OLE_COLOR)
    script_Dream6Color = newVal
    If xRefresh Then RefreshView
End Property

Public Property Get LineNumber() As Long
    LineNumber = txtScript.GetLineFromChar(txtScript.SelStart)
End Property

Public Property Get Enabled() As Boolean
    Enabled = txtScript.Enabled
End Property
Public Property Let Enabled(ByVal newVal As Boolean)
    txtScript.Enabled = newVal
End Property

Public Property Get TextRTF() As String
    TextRTF = txtScript.TextRTF
End Property

Public Property Get Language() As String
    Language = xLanguage
End Property
Public Property Let Language(ByVal newVal As String)
    Cancel = True
    Select Case LCase(Replace(newVal, " ", ""))
        Case "none", "plaintext", "plain", "text", "txt", "plaintxt", "text/plain", ""
            xLanguage = "PlainText"
            txtScript.BackColor = RGB(&HFF, &HFF, &HFF)
        Case "js", "jscript", "javascript", "text/javascript", "application/javascript"
            xLanguage = "JScript"
            txtScript.BackColor = RGB(&HFF, &HFF, &HFF)
        Case "vb", "vbs", "visualbasicscript", "visualbasic", "vbscript", "text/vbscript", "application/vbscript"
            xLanguage = "VBScript"
            txtScript.BackColor = RGB(&HFF, &HFF, &HFF)
        Case "rain", "bow", "rainbow", "dream", "ream", "reem"
            xLanguage = "Rainbow"
            txtScript.BackColor = RGB(&HFF, &HFF, &HFF) 'RGB(&H0, &H0, &H0)
        Case "defined", "custom", "define", "line", "lines"
            xLanguage = "Defined"
            txtScript.BackColor = RGB(&HFF, &HFF, &HFF) 'RGB(&H0, &H0, &H0)
            LineDefines = -1
        Case Else
            xLanguage = "Unknown"
            txtScript.BackColor = RGB(&HFF, &HFF, &HFF)
    End Select
    Cancel = False
    If xRefresh Then RefreshView
End Property

Public Property Let ErrorLine(ByVal newVal As Long)

    xErrorLine = newVal
    If xRefresh Then RefreshView

End Property

Public Property Get ErrorLine() As Long
    ErrorLine = xErrorLine
End Property

Friend Function FillFilePass(ByRef startByte As Long, ByRef stopByte As Long, ByRef idEvent As Long) As Boolean

    If xUndoDelay > 0 Then
    
        If CSng(Timer - xLatency) >= xUndoDelay Then
            AddUndo
            xLatency = Timer
        End If
    End If
        
    If (FileName <> "") Then
        Dim xTemp As Boolean

        If Not EOF(xFileNum) Then
            If stopByte = 0 Then stopByte = LOF(xFileNum)
        
            Dim str As String
            Dim visRange As RangeType
            visRange = VisibleRange

            Dim linecnt As Long
            If ((startByte >= visRange.StartPos) And _
                    (startByte <= visRange.StopPos)) Then
                linecnt = visRange.StopPos - visRange.StartPos
                If linecnt <= 0 Then linecnt = 100000
            Else
                linecnt = 100000
            End If
            
            xTemp = ((startByte >= visRange.StartPos) And (startByte <= visRange.StopPos)) Or _
                 ((startByte + linecnt >= visRange.StartPos) And (startByte + linecnt <= visRange.StopPos))
        
            
            visRange.StartPos = 0
            visRange.StopPos = stopByte
            If ((startByte + linecnt) > visRange.StopPos) Then
                linecnt = visRange.StopPos - startByte
                
            End If
            

            If (linecnt > 0) Then
                str = String(linecnt, Chr(0))
                Get #xFileNum, startByte + 1, str
                str = Replace(str, Chr(0), "")
                xRTF.Concat str
        
                If xTemp Then
                    DisableRedraw
                    SendMessageString txtScript.hwnd, WM_SETTEXT, ByVal 0&, xRTF.GetString(, ((startByte >= visRange.StopPos) Or EOF(xFileNum))) & Chr(0)
                    
                    Retext
                    ResetUndoRedo
                
                    EnableRedraw
                    
                    LineDefines(-1) = -1
                    
                    RefreshView
                
                    xTemp = False
                
                End If
                
                startByte = startByte + linecnt
                
            End If

            If EOF(xFileNum) Or startByte >= LOF(xFileNum) Then
                DisableRedraw
            
                SendMessageString txtScript.hwnd, WM_SETTEXT, ByVal 0&, xRTF.GetString(, True) & Chr(0)

                Retext
                ResetUndoRedo
                
                EnableRedraw
                
                LineDefines(-1) = -1
                
                RefreshView

                CloseFileName
                FillFilePass = True
            End If

        End If
            
    End If

End Function

Friend Function GetTextRange(ByRef Range As RangeType) As String

    If (Range.StopPos - Range.StartPos) > 0 Then
        Dim txt As TextRange
        Dim str As String
        Dim siz As Long

        str = Space((Range.StopPos - Range.StartPos))
        txt.Range.StartPos = Range.StartPos
        txt.Range.StopPos = Range.StopPos

        RtlMoveMemory txt.lpStr, ByVal VarPtr(str), 4&
        If SendMessageLngPtr(txtScript.hwnd, EM_GETTEXTRANGE, 0, ByVal VarPtr(txt)) Then
            RtlMoveMemory ByVal VarPtr(str), txt.lpStr, 4&
            GetTextRange = Left(StrConv(str, vbUnicode), Range.StopPos - Range.StartPos)
        End If
        
    End If

End Function

Public Function SelectRow(ByVal CursorRow As Long, Optional ByVal FullRowSelect As Boolean = True)

    Dim cnt As Integer
    Dim pNext As Integer
    Dim done As Boolean
    Dim newText As String
    Dim cPos As Long
    CursorRow = CursorRow - 1
        
    newText = txtScript.Text
    pNext = 1
    done = False
    cnt = 0
    cPos = 0
    If cnt < CursorRow Then
        cPos = 1
        Do
            pNext = InStr(pNext, newText, Chr(13))
            If pNext > 0 Then
                If cnt < CursorRow Then
                    cPos = cPos + Len(Left(newText, pNext))
                    newText = Mid(newText, pNext + 1)
                    pNext = 1
                Else
                    done = True
                End If
            Else
                done = True
            End If
            cnt = cnt + 1
        Loop Until done Or newText = ""
    End If
    
    pNext = InStr(newText, Chr(13))
    If pNext > 0 Then
        newText = Left(newText, pNext - 1)
    End If
    
    txtScript.SelStart = cPos
    If FullRowSelect Then
        txtScript.SelLength = Len(newText)
    Else
        txtScript.SelLength = 0
    End If

End Function
Public Sub TabIndent(ByVal Add As Boolean)
    ModifyIndent Add, False
End Sub

Public Sub Comment(ByVal Add As Boolean)
    ModifyIndent Add, True
End Sub

Public Function CanUndo() As Boolean
    CanUndo = ((UBound(xUndoText) > 0) And (xUndoStage > 0)) And (Not Locked) And (xUndoDelay > -1)
End Function
Public Function CanRedo() As Boolean
    CanRedo = ((xUndoStage < UBound(xUndoText)) And (UBound(xUndoText) > 0)) And (Not Locked) And (xUndoDelay > -1)
End Function

Public Sub Undo()
    If CanUndo Then
        If xUndoDelay > 0 Then AddUndo
        
        DisableRedraw
        xUndoStage = xUndoStage - 1
        txtScript.Text = xUndoText(xUndoStage)
        xRedrawSel.StartPos = xUndoStart(xUndoStage)
        xRedrawSel.StopPos = xRedrawSel.StartPos + xUndoLength(xUndoStage)
        Retext
        
        xUndoDirty = True
        
        EnableRedraw
        If xRefresh Then RefreshView
        RaiseEvent Change
    End If
End Sub

Public Sub Redo()
    If CanRedo Then
        DisableRedraw
        xUndoStage = xUndoStage + 1
        txtScript.Text = xUndoText(xUndoStage)
        xRedrawSel.StartPos = xUndoStart(xUndoStage)
        xRedrawSel.StopPos = xRedrawSel.StartPos + xUndoLength(xUndoStage)
        Retext
        
        xUndoDirty = True
        EnableRedraw
        If xRefresh Then RefreshView
        RaiseEvent Change
    End If
End Sub

Private Function CursorAtHome1() As Boolean
    If txtScript.SelStart = 0 Then
        CursorAtHome1 = True
    Else
        If Mid(txtScript.Text, txtScript.SelStart, 1) = Chr(10) Then
            CursorAtHome1 = True
        Else
            CursorAtHome1 = False
        End If
    End If
End Function
Private Function CursorAtHome2() As Boolean
    Dim nChar1 As String
    Dim nChar2 As String
    
    If txtScript.SelStart = 0 Then
        nChar1 = Chr(10)
    Else
        nChar1 = Mid(txtScript.Text, txtScript.SelStart, 1)
    End If

    If (txtScript.SelStart + 1) > SendMessageLngPtr(txtScript.hwnd, WM_GETTEXTLENGTH, 0&, 0&) Then
        nChar2 = Chr(13)
    Else
        nChar2 = Mid(txtScript.Text, txtScript.SelStart + 1, 1)
    End If
    
    Select Case nChar1
        Case Chr(10), Chr(9), " "
            Select Case nChar2
                Case Chr(9), " "
                    CursorAtHome2 = False
                Case Else
                    CursorAtHome2 = True
            End Select
        Case Else
            CursorAtHome2 = False
    End Select

End Function
Private Sub SetCursorAtHome2()
    Cancel = True
    SetCursorAtHome1
    If Not CursorAtHome2 = CursorAtHome1 Then
        Do Until CursorAtHome2 Or (txtScript.SelStart + 1 > SendMessageLngPtr(txtScript.hwnd, WM_GETTEXTLENGTH, 0&, 0&))
            txtScript.SelStart = txtScript.SelStart + 1
        Loop
    End If
    Cancel = False
End Sub
Private Sub SetCursorAtHome1()
    Cancel = True
    Do Until CursorAtHome1 Or (txtScript.SelStart = 0)
        txtScript.SelStart = txtScript.SelStart - 1
    Loop
    Cancel = False
End Sub
Private Sub txtScript_Cut()
    If Not Locked Then modEditor.SendMessage txtScript.hwnd, WM_CUT, 0, ByVal 0&
End Sub
Private Sub txtScript_Copy()
    modEditor.SendMessage txtScript.hwnd, WM_COPY, 0, ByVal 0&
End Sub
Private Sub txtScript_Paste()
    If Not Locked Then modEditor.SendMessage txtScript.hwnd, WM_PASTE, 0, ByVal 0&
End Sub
Private Sub txtScript_Clear()
    If Not Locked Then modEditor.SendMessage txtScript.hwnd, WM_CLEAR, 0, ByVal 0&
End Sub
Public Sub Cut()
    txtScript_Cut
End Sub
Public Sub Copy()
    txtScript_Copy
End Sub
Public Sub Paste()
    txtScript_Paste
End Sub
Public Sub Clear()
    txtScript_Clear
End Sub

Private Sub mnuComment_Click()
    Comment True
End Sub

Private Sub mnuCopy_Click()
    txtScript_Copy
End Sub

Private Sub mnuCut_Click()
    txtScript_Cut
End Sub

Private Sub mnuDelete_Click()
    If txtScript.SelLength = 0 Then txtScript.SelLength = 1
    txtScript.SelText = ""
End Sub

Private Sub mnuEdit_Click()
    On Error Resume Next
    RefreshEditMenu
    If Err Then Err.Clear
    On Error GoTo 0
End Sub

Private Sub mnuPaste_Click()
    txtScript_Paste
End Sub

Private Sub mnuRedo_Click()
    Redo
End Sub

Private Sub mnuSelectAll_Click()
    txtScript.SelStart = 0
    txtScript.SelLength = SendMessageLngPtr(txtScript.hwnd, WM_GETTEXTLENGTH, 0&, 0&)
End Sub

Private Sub mnuUncomment_Click()
    Comment False
End Sub

Private Sub mnuUndo_Click()
    Undo
End Sub

Private Sub txtScript_Change()

    If Not Cancel Then
        Cancel = True

        xCountLength.StartPos = SendMessageLngPtr(txtScript.hwnd, EM_GETLINECOUNT, 0&, 0&)
        xCountLength.StopPos = SendMessageLngPtr(txtScript.hwnd, WM_GETTEXTLENGTH, 0&, 0&)
        
        Dim xRedrawLines As RangeType
        xRedrawLines.StopPos = xRedrawSizes.StopPos
        xRedrawLines.StartPos = xCountLength.StartPos
        xRedrawSizes.StopPos = xRedrawSizes.StartPos
        xRedrawSizes.StartPos = xCountLength.StopPos
            
        'Debug.Print "Prior Char Size: " & xRedrawSizes.StopPos & "  Now Char Size: " & xRedrawSizes.StartPos
        'Debug.Print "Prior Line Count: " & xRedrawLines.StopPos & "  Now Line Count: " & xRedrawLines.StartPos
        
        Dim cnt As Long
        Dim cnt2 As Long
        Dim Amount As Long
        Dim StartAt As Long
        
        If Language = "Defined" Then
        
            If xRedrawLines.StartPos = xRedrawLines.StopPos Then
                StartAt = (GetLineNumber(txtScript.hwnd, xRedrawSel.StopPos) + 1)
                Amount = 0
            Else
        
                If xRedrawLines.StartPos > xRedrawLines.StopPos Then 'lines added
                    StartAt = (GetLineNumber(txtScript.hwnd, xRedrawSel.StopPos) + 1)
                    Amount = (xRedrawLines.StartPos - xRedrawLines.StopPos)
    
                    If StartAt <= UBound(xDefined) Then
                        'Debug.Print "Starts: " & StartAt & " Adding: " & Amount
                        ReDim Preserve xDefined(0 To UBound(xDefined) + Amount) As LineDef
        
                        For cnt = 1 To (UBound(xDefined) - StartAt - Amount)
                            xDefined(UBound(xDefined) - cnt) = xDefined(UBound(xDefined) - Amount - cnt)
                        Next
                        
                        For cnt = StartAt To StartAt + Amount
                            xDefined(cnt).LineClr = IIf(xDefined(cnt).LineClr = 0, -100, -Abs(xDefined(cnt).LineClr))
                            'Debug.Print "Redefine: " & cnt
                        Next
                    End If
                End If
        
                '^^maybe preform exact depending on performance to big pastes, for now::
            End If
        End If
    
        SendMessageStruct txtScript.hwnd, EM_EXGETSEL, 0, xRedrawSel
        
        If Language = "Defined" Then
            If xRedrawLines.StartPos = xRedrawLines.StopPos Then
                If StartAt <= UBound(xDefined) Then
                    xDefined(StartAt).LineClr = IIf(xDefined(StartAt).LineClr = 0, -100, -Abs(xDefined(StartAt).LineClr))
                End If
            Else
    
                'Debug.Print "Prior Selection: Start: " & xRedrawRange.StartPos & "  Stop: " & xRedrawRange.StopPos
                ''Debug.Print "Now Selection: Start: " & xRedrawSel.StartPos & "  Stop: " & xRedrawSel.StopPos
        
                If xRedrawLines.StartPos < xRedrawLines.StopPos Then 'lines removed
                    StartAt = (GetLineNumber(txtScript.hwnd, xRedrawSel.StartPos) + 1)
                    Amount = -(xRedrawLines.StartPos - xRedrawLines.StopPos)
                    'Debug.Print "Starts: " & StartAt & " Removing: " & Amount & " Redefine: " & StartAt
                    If StartAt <= UBound(xDefined) And Amount <= (GetLineNumber(txtScript.hwnd, xCountLength.StopPos) + 1) Then
                    
                        For cnt = StartAt To UBound(xDefined) - Amount
                            xDefined(cnt) = xDefined(cnt + Amount)
                        Next
                        
                        ReDim Preserve xDefined(0 To UBound(xDefined) - Amount) As LineDef
                    End If
                End If
    
                '^^maybe preform exact depending on performance to big pastes, for now::
        
            End If
        End If
            
        'condition in stops here for values
        '###################################
        xRedrawRange.StartPos = xRedrawSel.StartPos
        xRedrawRange.StopPos = xRedrawSel.StopPos
        xRedrawSizes.StartPos = xCountLength.StopPos
        xRedrawSizes.StopPos = xCountLength.StartPos
        
        If xUndoDelay = 0 Then
            AddUndo
        Else
            xLatency = Timer
        End If

        Cancel = False
    
        RaiseEvent Change

        RefreshView
    End If
            
End Sub

Private Sub txtScript_SelChange()
        
    If Not Cancel Then
        Cancel = True

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
    
       'If (Not Cancel) Then
            Cancel = True

            SendMessageStruct txtScript.hwnd, EM_EXGETSEL, 0, xRedrawSel

            Cancel = False
            
            'only what we want chaged by this event at this point is properly met
            'Debug.Print "Prior Selection: Start: " & xRedrawRange.StartPos & "  Stop: " & xRedrawRange.StopPos
            'Debug.Print "Now Selection: Start: " & xRedrawSel.StartPos & "  Stop: " & xRedrawSel.StopPos

    
        'End If
        
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

        RaiseEvent SelChange
        
        Cancel = False
    End If
    
End Sub

Private Sub txtScript_KeyDown(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyDown(KeyCode, Shift)
    Debug.Print KeyCode
    
    If KeyCode > 0 Then
    
        If (Shift = 2) And (KeyCode = 82) Then
            KeyCode = 0
            Redo
        End If
        
        If KeyCode = 45 And Shift = 1 Then
            KeyCode = 0
            txtScript_Paste
        End If
    
        If KeyCode = 45 And Shift = 2 And (txtScript.SelLength > 0) Then
            KeyCode = 0
            txtScript_Copy
        End If
    
        If KeyCode = 46 And (txtScript.SelLength > 0) Then
            KeyCode = 0
            txtScript_Clear
        End If
        
        If KeyCode = 9 Then
    
            If (txtScript.SelLength > 0) And (InStr(txtScript.SelText, vbCrLf) > 0) Then
                KeyCode = 0
                ModifyIndent (Shift = 0), False
            End If
    
        End If
        
        If KeyCode = 36 Then
        
            If Not (Shift = 1) Then
            
                If (txtScript.SelLength > 0) Then
                    txtScript.SelLength = 0
                    SetCursorAtHome2
                Else
                    If CursorAtHome1 Then
                        SetCursorAtHome2
                    ElseIf CursorAtHome2 Then
                        SetCursorAtHome1
                    Else
                        SetCursorAtHome2
                    End If
                End If
        
                KeyCode = 0
            End If
        End If
    End If
    
End Sub

Private Sub txtScript_KeyPress(KeyAscii As Integer)
    
    RaiseEvent KeyPress(KeyAscii)
    If KeyAscii > 0 Then
        If KeyAscii = 22 Then
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
            Undo
        End If
        
        If KeyAscii = 18 Then
            KeyAscii = 0
        End If
    
        If KeyAscii = 9 Then
            If (txtScript.SelLength > 0) And (InStr(txtScript.SelText, vbCrLf) > 0) Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub txtScript_Click()
    RaiseEvent Click
End Sub

Private Sub txtScript_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub txtScript_KeyUp(KeyCode As Integer, Shift As Integer)
    
    RaiseEvent KeyUp(KeyCode, Shift)
    If KeyCode > 0 Then
        If KeyCode = 45 And Shift = 1 Then
            KeyCode = 0
        End If
    
        If KeyCode = 45 And Shift = 2 And (txtScript.SelLength > 0) Then
            KeyCode = 0
        End If
    
        If KeyCode = 46 And (txtScript.SelLength > 0) Then
            KeyCode = 0
        End If
        If KeyCode = 33 Then 'pageup
            RefreshView
        End If
        
        If KeyCode = 34 Then 'pagedown
            RefreshView
        End If
    End If
End Sub

Private Sub txtScript_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        RefreshEditMenu
        PopupMenu mnuEdit
    End If
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub txtScript_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub txtScript_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
Private Sub ResetUndoRedo()
    ReDim xUndoText(0) As String
    ReDim xUndoStart(0) As Long
    ReDim xUndoLength(0) As Long
    xUndoStage = 0
    xUndoDirty = False
    xUndoBuffer = xUndoStack
    xUndoText(0) = txtScript.Text
    xUndoStart(0) = txtScript.SelStart
    xUndoLength(0) = txtScript.SelLength
End Sub

Private Sub AddUndo()

    If (xUndoStack <> 0) Then
    
        If ((UBound(xUndoText) < xUndoBuffer) Or (xUndoStack = -1)) Or (Not (xUndoStage = UBound(xUndoText))) Then
        
            ReDim Preserve xUndoText(xUndoStage + 1) As String
            ReDim Preserve xUndoStart(xUndoStage + 1) As Long
            ReDim Preserve xUndoLength(xUndoStage + 1) As Long
            xUndoStage = xUndoStage + 1
        ElseIf (UBound(xUndoText) = xUndoBuffer) Or (xUndoStage = UBound(xUndoText)) Then
        
            Dim cnt As Long
            For cnt = (LBound(xUndoText) + 1) To UBound(xUndoText)
                xUndoText(cnt - 1) = xUndoText(cnt)
                xUndoStart(cnt - 1) = xUndoStart(cnt)
                xUndoLength(cnt - 1) = xUndoLength(cnt)
            Next
    
        End If

        xUndoText(UBound(xUndoText)) = txtScript.Text
        xUndoStart(UBound(xUndoStart)) = txtScript.SelStart
        xUndoLength(UBound(xUndoLength)) = txtScript.SelLength
    
        xUndoDirty = True
        
    ElseIf xUndoDirty Then
        ResetUndoRedo
    End If

End Sub

Private Sub txtScript_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub txtScript_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
    If xHooked Then HookObj txtScript
    
    DisableRedraw
    
    Dim pt As PointType
    
    pt.X = (X / Screen.TwipsPerPixelX)
    pt.Y = (Y / Screen.TwipsPerPixelY)
    txtScript.SelStart = SendMessageStruct(txtScript.hwnd, EM_CHARFROMPOS, 0&, pt)
    txtScript.SelLength = 0

    If Data.GetFormat(ClipBoardConstants.vbCFText) Then
        txtScript.SelText = Data.GetData(ClipBoardConstants.vbCFText)
    End If
    
    EnableRedraw

    RefreshView
    
    If xHooked Then HookObj txtScript
        
    Call txtProxy_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub txtScript_OLEDragOver(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    
    Dim pt As PointType
    
    pt.X = (X / Screen.TwipsPerPixelX)
    pt.Y = (Y / Screen.TwipsPerPixelY)
    txtScript.SelStart = SendMessageStruct(txtScript.hwnd, EM_CHARFROMPOS, 0&, pt)
    
    Call txtProxy_OLEDragOver(Data, Effect, Button, Shift, X, Y, State)

End Sub


Private Sub txtScript_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub txtScript_OLESetData(Data As RichTextLib.DataObject, DataFormat As Integer)
    Call txtProxy_OLESetData(Data, DataFormat)
End Sub

Private Sub txtScript_OLEStartDrag(Data As RichTextLib.DataObject, AllowedEffects As Long)
    Call txtProxy_OLEStartDrag(Data, AllowedEffects)
End Sub


Private Sub txtProxy_OLEDragDrop(Data As Object, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub txtProxy_OLEDragOver(Data As Object, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)

    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)

End Sub

Private Sub txtProxy_OLESetData(Data As Object, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub txtProxy_OLEStartDrag(Data As Object, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub


Private Sub UserControl_Initialize()
    Cancel = True
    
    ReDim xDefined(0 To 0) As LineDef
    Set xRTF = New Strand
    
    xRefresh = True
    Set txtScript = txtScript1
    
    SendMessageLngPtr txtScript1.hwnd, EM_SETUNDOLIMIT, 0, ByVal 0&
    SendMessageLngPtr txtScript2.hwnd, EM_SETUNDOLIMIT, 0, ByVal 0&
    
    txtScript1.TextRTF = "{\rtf1\ansi\deff0" & CreateFontTable & CreateColorTable & GetLineHeader & "\line\cf0 }"
    txtScript2.TextRTF = "{\rtf1\ansi\deff0" & CreateFontTable & CreateColorTable & GetLineHeader & "\line\cf0 }"

    txtScript1.SelTabCount = 4
    txtScript2.SelTabCount = 4

    xUndoStack = 150
    
    ResetUndoRedo

    Cancel = False
    If Not IsRunningMode Then
        xHooked = True
        HookObj txtScript
    End If
    
    xLatency = Timer
    SetTimer txtScript.hwnd, ObjPtr(Me), 50, AddressOf FileProc

End Sub
Private Sub AdjustWordWrap()
    If WordWrap Then
        Dim rct As RectType
        If GetClientRect(txtScript.hwnd, rct) Then
            txtScript.RightMargin = IIf(WordWrap, (rct.Right - rct.Left) * Screen.TwipsPerPixelX, 200000)
        Else
            txtScript.RightMargin = IIf(WordWrap, txtScript.Width - (30 * Screen.TwipsPerPixelX), 200000)
        End If
    Else
        txtScript.RightMargin = 200000
    End If
End Sub
Private Sub RefreshEditMenu()
    mnuUndo.Enabled = CanUndo And Enabled And Not Locked
    mnuRedo.Enabled = CanRedo And Enabled And Not Locked
    mnuCut.Enabled = Enabled And Not Locked
    mnuCopy.Enabled = (SelLength > 0)
    mnuPaste.Enabled = Enabled And Not Locked
    mnuDelete.Enabled = Enabled And Not Locked
    mnuSelectAll.Enabled = Enabled
    mnuComment.Enabled = (Enabled And Not Locked And (CountWord(txtScript.SelText, vbCrLf) > 1))
    mnuUncomment.Enabled = (Enabled And Not Locked And (CountWord(txtScript.SelText, vbCrLf) > 1))
End Sub

'Private Sub Cache_Initialize()
'    xSource = GetTemporaryFile
'    xHandle = FreeFile
'    Open xSource For Random As #xHandle
'End Sub
'
'Private Sub Cache_Terminate()
'    Close #xHandle
'    Kill xSource
'End Sub

Private Sub UserControl_Show()
    AdjustWordWrap
End Sub

Friend Sub CloseFileName(Optional ByVal PreSet As String = "")
    If (xFileName <> "") Or (PreSet = "") Then
        Close #xFileNum
        xFileName = ""
        xFileNum = 0
        xRTF.Reset
    End If
End Sub

Private Sub UserControl_Terminate()

    KillTimer txtScript.hwnd, ObjPtr(Me)
    
    If xHooked Then

        HookObj txtScript
        xHooked = False
    End If
    
    Cancel = True
    
    CloseFileName

    Set txtScript = Nothing
    
    xRTF.Reset
    Set xRTF = Nothing
    
    Cancel = False

End Sub

Private Sub UserControl_InitProperties()
    
    Set txtScript = txtScript1
    
    xLanguage = "Unknown"
    
    txtScript.Font.Name = "Lucida Console"
    txtScript.Font.Italic = False
    txtScript.Font.Bold = False
    txtScript.Font.StrikeThrough = False
    txtScript.Font.Underline = False
    txtScript.Font.Size = 10
    txtScript.Font.Weight = 400
    txtScript.Font.Charset = 0
    txtScript.BackColor = SystemColorConstants.vbWindowBackground
    
    txtScript.Locked = False
    txtScript.Enabled = True
    
    script_CommentColor = &H8000&
    script_ErrorColor = &HFF&
    script_OperatorColor = &H808080
    script_StatementColor = &HFF0000
    script_TextColor = SystemColorConstants.vbWindowText
    script_ValueColor = &H404040
    script_VariableColor = &H800000
    
    script_Dream1Color = &HC000C0
    script_Dream2Color = &HC00000
    script_Dream3Color = &HC000&
    script_Dream4Color = &HFFFF&
    script_Dream5Color = &H80FF&
    script_Dream6Color = &HFF&
    
    AdjustWordWrap

End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    Enabled = PropBag.ReadProperty("Enabled", True)
    
    Language = PropBag.ReadProperty("Language", "Unknown")
    
    Font.Name = PropBag.ReadProperty("FontName", "Lucida Console")
    Font.Italic = PropBag.ReadProperty("FontItalic", False)
    Font.Bold = PropBag.ReadProperty("FontBold", False)
    Font.StrikeThrough = PropBag.ReadProperty("FontStrikeThrough", False)
    Font.Underline = PropBag.ReadProperty("FontUnderline", False)
    Font.Size = PropBag.ReadProperty("FontSize", 10)
    Font.Weight = PropBag.ReadProperty("FontWeight", 400)
    Font.Charset = PropBag.ReadProperty("FontCharset", 0)
    
    Locked = PropBag.ReadProperty("Locked", False)
    
    BackColor = PropBag.ReadProperty("BackColor", SystemColorConstants.vbWindowBackground)
    
    ColorComment = PropBag.ReadProperty("ColorComment", &H8000&)
    ColorError = PropBag.ReadProperty("ColorError", &HFF&)
    ColorOperator = PropBag.ReadProperty("ColorOperator", &H808080)
    ColorStatement = PropBag.ReadProperty("ColorStatement", &HFF0000)
    ColorText = PropBag.ReadProperty("ColorText", SystemColorConstants.vbWindowText)
    ColorValue = PropBag.ReadProperty("ColorValue", &H404040)
    ColorVariable = PropBag.ReadProperty("ColorVariable", &H800000)
    
    AutoRedraw = PropBag.ReadProperty("AutoRedraw", True)
    
    ColorDream1 = PropBag.ReadProperty("ColorDream1", &H800080)
    ColorDream2 = PropBag.ReadProperty("ColorDream2", &H800000)
    ColorDream3 = PropBag.ReadProperty("ColorDream3", &H808000)
    ColorDream4 = PropBag.ReadProperty("ColorDream4", &H8000&)
    ColorDream5 = PropBag.ReadProperty("ColorDream5", &H8080&)
    ColorDream6 = PropBag.ReadProperty("ColorDream6", &H4080&)

    WordWrap = PropBag.ReadProperty("WordWrap", False)
    xUndoDelay = PropBag.ReadProperty("UndoDelay", 0)
    
End Sub

Private Sub UserControl_Resize()

    txtScript.Left = 0
    txtScript.Top = 0
    txtScript.Width = UserControl.Width
    txtScript.Height = UserControl.Height

    RefreshView

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    PropBag.WriteProperty "Enabled", Enabled, True
    PropBag.WriteProperty "Language", Language, "Unknown"
    
    PropBag.WriteProperty "FontName", Font.Name, "Lucida Console"
    PropBag.WriteProperty "FontItalic", Font.Italic, False
    PropBag.WriteProperty "FontBold", Font.Bold, False
    PropBag.WriteProperty "FontStrikeThrough", Font.StrikeThrough, False
    PropBag.WriteProperty "FontUnderline", Font.Underline, False
    PropBag.WriteProperty "FontSize", Font.Size, 10
    PropBag.WriteProperty "FontWeight", Font.Weight, 400
    PropBag.WriteProperty "FontCharset", Font.Charset, 0
    PropBag.WriteProperty "Locked", Locked, False
    PropBag.WriteProperty "WordWrap", WordWrap, False
    PropBag.WriteProperty "BackColor", BackColor, SystemColorConstants.vbWindowBackground
    
    PropBag.WriteProperty "ColorComment", ColorComment, &H8000&
    PropBag.WriteProperty "ColorError", ColorError, &HFF&
    PropBag.WriteProperty "ColorOperator", ColorOperator, &H808080
    PropBag.WriteProperty "ColorStatement", ColorStatement, &HFF0000
    PropBag.WriteProperty "ColorText", ColorText, SystemColorConstants.vbWindowText
    PropBag.WriteProperty "ColorValue", ColorValue, &H404040
    PropBag.WriteProperty "ColorVariable", ColorVariable, &H800000
    
    PropBag.WriteProperty "AutoRedraw", AutoRedraw, True
    
    PropBag.WriteProperty "ColorDream1", ColorDream1, &HC000C0
    PropBag.WriteProperty "ColorDream2", ColorDream2, &HC00000
    PropBag.WriteProperty "ColorDream3", ColorDream3, &HC000&
    PropBag.WriteProperty "ColorDream4", ColorDream4, &HFFFF&
    PropBag.WriteProperty "ColorDream5", ColorDream5, &H80FF&
    PropBag.WriteProperty "ColorDream6", ColorDream6, &HFF&
    
    PropBag.WriteProperty "UndoDelay", xUndoDelay, 0
End Sub

Private Function RemoveNextArg(ByRef TheParams As String, ByVal TheSeperator As String, Optional ByVal RemoveSeperator As Boolean = True) As String
    Dim retVal As String
    If InStr(TheParams, TheSeperator) > 0 Then
        retVal = Left(TheParams, InStr(TheParams, TheSeperator) - 1) & IIf(RemoveSeperator, "", TheSeperator)
        TheParams = Mid(TheParams, InStr(TheParams, TheSeperator) + Len(TheSeperator))
    Else
        retVal = TheParams
        TheParams = ""
    End If
    RemoveNextArg = retVal
End Function

Private Function TrimStrip(ByVal TheStr As String, ByVal TheChar As String) As String
    Do While Left(TheStr, Len(TheChar)) = TheChar
        TheStr = Mid(TheStr, Len(TheChar) + 1)
    Loop
    Do While Right(TheStr, Len(TheChar)) = TheChar
        TheStr = Left(TheStr, Len(TheStr) - Len(TheChar))
    Loop
    TrimStrip = TheStr
End Function

Friend Function DisableRedraw()
    Cancel = True
        
    LockWindowUpdate txtScript.hwnd

    xRedrawVScroll = GetScrollPos(txtScript.hwnd, SB_VERT)
    xRedrawHScroll = GetScrollPos(txtScript.hwnd, SB_HORZ)
    
    SendMessageStruct txtScript.hwnd, EM_EXGETSEL, 0, xRedrawSel
    txtScript.HideSelection = True
    SendMessageLngPtr txtScript.hwnd, EM_HIDESELECTION, True, 0
    
End Function

Friend Function EnableRedraw()
    On Error GoTo clearerr
    
    If xRedrawVScroll <> GetScrollPos(txtScript.hwnd, SB_VERT) Then
        SendMessageLngPtr txtScript.hwnd, WM_VSCROLL, SB_THUMBPOSITION Or &H10000 * xRedrawVScroll, 0&
    End If
    If xRedrawHScroll <> GetScrollPos(txtScript.hwnd, SB_HORZ) Then
        SendMessageLngPtr txtScript.hwnd, WM_HSCROLL, SB_THUMBPOSITION Or &H10000 * xRedrawHScroll, 0&
    End If

    GoTo passclear
clearerr:
    Err.Clear
passclear:
    
    SendMessageStruct txtScript.hwnd, EM_EXSETSEL, 0, xRedrawSel
    txtScript.HideSelection = False
    SendMessageLngPtr txtScript.hwnd, EM_HIDESELECTION, False, 0
    
    LockWindowUpdate 0&
    
    Cancel = False
    On Error GoTo 0
End Function

Private Function VisibleRange() As RangeType
    Dim pt As PointType

    pt.X = 0
    pt.Y = 0
    VisibleRange.StartPos = SendMessageStruct(txtScript.hwnd, EM_CHARFROMPOS, 0&, pt)
    pt.X = txtScript.Width / Screen.TwipsPerPixelX
    pt.Y = txtScript.Height / Screen.TwipsPerPixelY
    VisibleRange.StopPos = SendMessageStruct(txtScript.hwnd, EM_CHARFROMPOS, 0&, pt)

End Function

'Private Function VisibleRange() As RangeType
'    Dim pt As PointType
'
'    pt.X = 0
'    pt.Y = 0
'    VisibleRange.StartPos = SendMessageStruct(txtScript.hWnd, EM_CHARFROMPOS, 0&, pt)
'    pt.X = (txtScript.Width / Screen.TwipsPerPixelX) - 1
'    pt.Y = (txtScript.Height / Screen.TwipsPerPixelY) - 1
'    VisibleRange.StopPos = SendMessageStruct(txtScript.hWnd, EM_CHARFROMPOS, 0&, pt)
'
'    pt.X = SendMessageLngPtr(txtScript.hWnd, EM_LINEFROMCHAR, VisibleRange.StopPos, 0&)
'    pt.Y = SendMessageLngPtr(txtScript.hWnd, EM_LINEINDEX, pt.X, 0&) + _
'            SendMessageLngPtr(txtScript.hWnd, EM_LINELENGTH, pt.X, 0&)
'    If pt.Y > VisibleRange.StopPos Then VisibleRange.StopPos = pt.Y
'    If VisibleRange.StopPos > xCountLength.StopPos Then VisibleRange.StopPos = xCountLength.StopPos
'
'End Function

Public Property Get Font() As StdFont
    UserControl.Font.Bold = UserControl.FontBold
    UserControl.Font.Italic = UserControl.FontItalic
    UserControl.Font.Name = UserControl.FontName
    UserControl.Font.Size = UserControl.FontSize
    UserControl.Font.StrikeThrough = UserControl.FontStrikethru
    UserControl.Font.Underline = UserControl.FontUnderline
    Set txtScript.Font = UserControl.Font
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByRef newVal As StdFont)
    Cancel = True
    Set UserControl.Font = newVal
    Set txtScript.Font = newVal
    UserControl.FontBold = UserControl.Font.Bold
    UserControl.FontItalic = UserControl.Font.Italic
    UserControl.FontName = UserControl.Font.Name
    UserControl.FontSize = UserControl.Font.Size
    UserControl.FontStrikethru = UserControl.Font.StrikeThrough
    UserControl.FontUnderline = UserControl.Font.Underline
    Cancel = False
    Refresh
End Property

Private Function CreateFontTable() As String
    
    Dim fontTable As New Strand
    fontTable.Concat "{\fonttbl{\f0 "
    fontTable.Concat txtScript.Font.Name
    fontTable.Concat ";}{\f1\fcharset0 "
    fontTable.Concat txtScript.Font.Name
    fontTable.Concat ";}}"
    
    CreateFontTable = fontTable.GetString

    Set fontTable = Nothing
End Function
Private Function CreateColorTable() As String
 
    Dim retStr As New Strand
    Dim Red As Long
    Dim Green As Long
    Dim Blue As Long

    With retStr

        .Concat "{\colortbl"
        
        ConvertColor script_TextColor, Red, Green, Blue
        .Concat "\red"
        .Concat Trim(str(Red))
        .Concat "\green"
        .Concat Trim(str(Green))
        .Concat "\blue"
        .Concat Trim(str(Blue))
        .Concat ";"
        RTFColorMap.Text.Red = Red
        RTFColorMap.Text.Green = Green
        RTFColorMap.Text.Blue = Blue
        num_Text = ""

        ConvertColor script_VariableColor, Red, Green, Blue
        .Concat "\red"
        .Concat Trim(str(Red))
        .Concat "\green"
        .Concat Trim(str(Green))
        .Concat "\blue"
        .Concat Trim(str(Blue))
        .Concat ";"
        RTFColorMap.Variable.Red = Red
        RTFColorMap.Variable.Green = Green
        RTFColorMap.Variable.Blue = Blue
        num_Variable = ""

        ConvertColor script_StatementColor, Red, Green, Blue
        .Concat "\red"
        .Concat Trim(str(Red))
        .Concat "\green"
        .Concat Trim(str(Green))
        .Concat "\blue"
        .Concat Trim(str(Blue))
        .Concat ";"
        RTFColorMap.Statement.Red = Red
        RTFColorMap.Statement.Green = Green
        RTFColorMap.Statement.Blue = Blue
        num_Statement = ""

        ConvertColor script_ValueColor, Red, Green, Blue
        .Concat "\red"
        .Concat Trim(str(Red))
        .Concat "\green"
        .Concat Trim(str(Green))
        .Concat "\blue"
        .Concat Trim(str(Blue))
        .Concat ";"
        RTFColorMap.Value.Red = Red
        RTFColorMap.Value.Green = Green
        RTFColorMap.Value.Blue = Blue
        num_Value = ""

        ConvertColor script_OperatorColor, Red, Green, Blue
        .Concat "\red"
        .Concat Trim(str(Red))
        .Concat "\green"
        .Concat Trim(str(Green))
        .Concat "\blue"
        .Concat Trim(str(Blue))
        .Concat ";"
        RTFColorMap.Operator.Red = Red
        RTFColorMap.Operator.Green = Green
        RTFColorMap.Operator.Blue = Blue
        num_Operator = ""

        ConvertColor script_CommentColor, Red, Green, Blue
        .Concat "\red"
        .Concat Trim(str(Red))
        .Concat "\green"
        .Concat Trim(str(Green))
        .Concat "\blue"
        .Concat Trim(str(Blue))
        .Concat ";"
        RTFColorMap.Comment.Red = Red
        RTFColorMap.Comment.Green = Green
        RTFColorMap.Comment.Blue = Blue
        num_Comment = ""

        ConvertColor script_ErrorColor, Red, Green, Blue
        .Concat "\red"
        .Concat Trim(str(Red))
        .Concat "\green"
        .Concat Trim(str(Green))
        .Concat "\blue"
        .Concat Trim(str(Blue))
        .Concat ";"
        RTFColorMap.Error.Red = Red
        RTFColorMap.Error.Green = Green
        RTFColorMap.Error.Blue = Blue
        num_Error = ""
        
        ConvertColor script_Dream1Color, Red, Green, Blue
        .Concat "\red"
        .Concat Trim(str(Red))
        .Concat "\green"
        .Concat Trim(str(Green))
        .Concat "\blue"
        .Concat Trim(str(Blue))
        .Concat ";"
        RTFColorMap.Dream1.Red = Red
        RTFColorMap.Dream1.Green = Green
        RTFColorMap.Dream1.Blue = Blue
        num_Dream1 = ""

        ConvertColor script_Dream2Color, Red, Green, Blue
        .Concat "\red"
        .Concat Trim(str(Red))
        .Concat "\green"
        .Concat Trim(str(Green))
        .Concat "\blue"
        .Concat Trim(str(Blue))
        .Concat ";"
        RTFColorMap.Dream2.Red = Red
        RTFColorMap.Dream2.Green = Green
        RTFColorMap.Dream2.Blue = Blue
        num_Dream2 = ""
        
        ConvertColor script_Dream3Color, Red, Green, Blue
        .Concat "\red"
        .Concat Trim(str(Red))
        .Concat "\green"
        .Concat Trim(str(Green))
        .Concat "\blue"
        .Concat Trim(str(Blue))
        .Concat ";"
        RTFColorMap.Dream3.Red = Red
        RTFColorMap.Dream3.Green = Green
        RTFColorMap.Dream3.Blue = Blue
        num_Dream3 = ""
        
        ConvertColor script_Dream4Color, Red, Green, Blue
        .Concat "\red"
        .Concat Trim(str(Red))
        .Concat "\green"
        .Concat Trim(str(Green))
        .Concat "\blue"
        .Concat Trim(str(Blue))
        .Concat ";"
        RTFColorMap.Dream4.Red = Red
        RTFColorMap.Dream4.Green = Green
        RTFColorMap.Dream4.Blue = Blue
        num_Dream4 = ""
        
        ConvertColor script_Dream5Color, Red, Green, Blue
        .Concat "\red"
        .Concat Trim(str(Red))
        .Concat "\green"
        .Concat Trim(str(Green))
        .Concat "\blue"
        .Concat Trim(str(Blue))
        .Concat ";"
        RTFColorMap.Dream5.Red = Red
        RTFColorMap.Dream5.Green = Green
        RTFColorMap.Dream5.Blue = Blue
        num_Dream5 = ""
        
        ConvertColor script_Dream6Color, Red, Green, Blue
        .Concat "\red"
        .Concat Trim(str(Red))
        .Concat "\green"
        .Concat Trim(str(Green))
        .Concat "\blue"
        .Concat Trim(str(Blue))
        .Concat ";"
        RTFColorMap.Dream6.Red = Red
        RTFColorMap.Dream6.Green = Green
        RTFColorMap.Dream6.Blue = Blue
        num_Dream6 = ""

        .Concat "}"
        
        Dim num As Integer
        num = 0

        Do
            Select Case num
                Case 0
                    Red = RTFColorMap.Text.Red
                    Green = RTFColorMap.Text.Green
                    Blue = RTFColorMap.Text.Blue
                Case 1
                    Red = RTFColorMap.Variable.Red
                    Green = RTFColorMap.Variable.Green
                    Blue = RTFColorMap.Variable.Blue
                Case 2
                    Red = RTFColorMap.Statement.Red
                    Green = RTFColorMap.Statement.Green
                    Blue = RTFColorMap.Statement.Blue
                Case 3
                    Red = RTFColorMap.Value.Red
                    Green = RTFColorMap.Value.Green
                    Blue = RTFColorMap.Value.Blue
                Case 4
                    Red = RTFColorMap.Operator.Red
                    Green = RTFColorMap.Operator.Green
                    Blue = RTFColorMap.Operator.Blue
                Case 5
                    Red = RTFColorMap.Comment.Red
                    Green = RTFColorMap.Comment.Green
                    Blue = RTFColorMap.Comment.Blue
                Case 6
                    Red = RTFColorMap.Error.Red
                    Green = RTFColorMap.Error.Green
                    Blue = RTFColorMap.Error.Blue
                Case 7
                    Red = RTFColorMap.Dream1.Red
                    Green = RTFColorMap.Dream1.Green
                    Blue = RTFColorMap.Dream1.Blue
                Case 8
                    Red = RTFColorMap.Dream2.Red
                    Green = RTFColorMap.Dream2.Green
                    Blue = RTFColorMap.Dream2.Blue
                Case 9
                    Red = RTFColorMap.Dream3.Red
                    Green = RTFColorMap.Dream3.Green
                    Blue = RTFColorMap.Dream3.Blue
                Case 10
                    Red = RTFColorMap.Dream4.Red
                    Green = RTFColorMap.Dream4.Green
                    Blue = RTFColorMap.Dream4.Blue
                Case 11
                    Red = RTFColorMap.Dream5.Red
                    Green = RTFColorMap.Dream5.Green
                    Blue = RTFColorMap.Dream5.Blue
                Case 12
                    Red = RTFColorMap.Dream6.Red
                    Green = RTFColorMap.Dream6.Green
                    Blue = RTFColorMap.Dream6.Blue
            End Select

            If RTFColorMap.Text.Red = Red And RTFColorMap.Text.Green = Green And RTFColorMap.Text.Blue = Blue Then
                If num_Text = "" Then num_Text = Trim(num)
            End If
            If RTFColorMap.Variable.Red = Red And RTFColorMap.Variable.Green = Green And RTFColorMap.Variable.Blue = Blue Then
                If num_Variable = "" Then num_Variable = Trim(num)
            End If
            If RTFColorMap.Statement.Red = Red And RTFColorMap.Statement.Green = Green And RTFColorMap.Statement.Blue = Blue Then
                If num_Statement = "" Then num_Statement = Trim(num)
            End If
            If RTFColorMap.Value.Red = Red And RTFColorMap.Value.Green = Green And RTFColorMap.Value.Blue = Blue Then
                If num_Value = "" Then num_Value = Trim(num)
            End If
            If RTFColorMap.Operator.Red = Red And RTFColorMap.Operator.Green = Green And RTFColorMap.Operator.Blue = Blue Then
                If num_Operator = "" Then num_Operator = Trim(num)
            End If
            If RTFColorMap.Comment.Red = Red And RTFColorMap.Comment.Green = Green And RTFColorMap.Comment.Blue = Blue Then
                If num_Comment = "" Then num_Comment = Trim(num)
            End If
            If RTFColorMap.Error.Red = Red And RTFColorMap.Error.Green = Green And RTFColorMap.Error.Blue = Blue Then
                If num_Error = "" Then num_Error = Trim(num)
            End If
            If RTFColorMap.Dream1.Red = Red And RTFColorMap.Dream1.Green = Green And RTFColorMap.Dream1.Blue = Blue Then
                If num_Dream1 = "" Then num_Dream1 = Trim(num)
            End If
            If RTFColorMap.Dream2.Red = Red And RTFColorMap.Dream2.Green = Green And RTFColorMap.Dream2.Blue = Blue Then
                If num_Dream2 = "" Then num_Dream2 = Trim(num)
            End If
            If RTFColorMap.Dream3.Red = Red And RTFColorMap.Dream3.Green = Green And RTFColorMap.Dream3.Blue = Blue Then
                If num_Dream3 = "" Then num_Dream3 = Trim(num)
            End If
            If RTFColorMap.Dream4.Red = Red And RTFColorMap.Dream4.Green = Green And RTFColorMap.Dream4.Blue = Blue Then
                If num_Dream4 = "" Then num_Dream4 = Trim(num)
            End If
            If RTFColorMap.Dream5.Red = Red And RTFColorMap.Dream5.Green = Green And RTFColorMap.Dream5.Blue = Blue Then
                If num_Dream5 = "" Then num_Dream5 = Trim(num)
            End If
            If RTFColorMap.Dream6.Red = Red And RTFColorMap.Dream6.Green = Green And RTFColorMap.Dream6.Blue = Blue Then
                If num_Dream6 = "" Then num_Dream6 = Trim(num)
            End If
            
            num = num + 1
        Loop Until num > 12
        
    End With
    
    CreateColorTable = retStr.GetString

    Set retStr = Nothing
End Function

Public Function ConvertColor(ByVal Color As Variant, Optional ByRef Red As Long, Optional ByRef Green As Long, Optional ByRef Blue As Long) As Long
On Error GoTo catch
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
    Red = CByte("&h" & Mid(Color, 5, 2))
    Green = CByte("&h" & Mid(Color, 3, 2))
    Blue = CByte("&h" & Mid(Color, 1, 2))
    ConvertColor = RGB(Red, Green, Blue)
    Exit Function
HTMLorHexColor:
    Red = Val("&H" & Left(Color, 2))
    Green = Val("&H" & Mid(Color, 3, 2))
    Blue = Val("&H" & Right(Color, 2))
    ConvertColor = RGB(Red, Green, Blue)
    Exit Function
catch:
    Err.Clear
    ConvertColor = 0
End Function

Private Function GetLineHeader() As String
    Dim FontNum As Integer
    Dim FontSize As Integer
    Dim FontString As String
    Dim lineHeader As String
   
    FontNum = 0
    FontSize = TextHeight
    FontString = ""
    If UserControl.Font.Bold Then FontString = "\b"
    If UserControl.Font.StrikeThrough Then FontString = FontString + "\strike"
    If UserControl.Font.Underline Then FontString = FontString + "\ul"
    
    GetLineHeader = "\plain\fs" + Trim(CStr(FontSize)) + "\f" + Trim(str(FontNum)) + FontString + "\cf" & num_Text
    
End Function

Friend Sub WalkRTFView(Optional ByVal StartPos As Long = -1, Optional ByVal StopPos As Long = -1)

    Static stacks As Integer
    stacks = stacks + 1

    If stacks > 1 Then
        stacks = stacks - 1
        Exit Sub
    End If
    
    'If Language = "Unknown" Then Exit Sub
    
    Dim ByLine As Boolean
    Dim startLine As Long
    Dim stopLine As Long
    Dim skipTo As Long

    Dim line As Long
    Dim Header As String
    Dim DoSwap As Boolean
    Dim lapse As Single

    Dim fullSize As Long
    Dim FullText As New Strand

    Dim LineText As New Strand

    Dim ni As New Strand
    Dim sr As RangeType

    fullSize = xCountLength.StopPos

    If (StartPos > -1) And (StopPos > -1) Then
        sr.StartPos = StartPos
        sr.StopPos = StopPos
        ByLine = True
    Else
        sr.StartPos = -1
        sr.StopPos = 0
        ByLine = False
    End If

    Header = GetLineHeader
    line = 1

    If ByLine Then
        line = GetLineNumber(txtScript.hwnd, sr.StartPos) + 1
        If line < 1 Then line = 1
        startLine = line
        stopLine = GetLineNumber(txtScript.hwnd, sr.StopPos) + 1
        If stopLine < startLine Then stopLine = startLine

        FullText.Concat Mid(txtScript.Text, sr.StartPos + 1)
        
    End If

    If Not ByLine Then
        FullText.Concat txtScript.Text
        ni.Concat "{\rtf1\ansi\deff0\fs" + Trim(CInt(TextHeight))
        ni.Concat CreateFontTable
        ni.Concat CreateColorTable
        ni.Concat Header
    End If

    If Language = "JScript" Then
        If xCountLength.StopPos > 0 And sr.StartPos > 1 Then
            Dim Find As Long
            If (InStrRev(Left(txtScript.Text, sr.StartPos), "/*") > InStr(Left(txtScript.Text, sr.StartPos), "*/")) _
                And Not ((InStrRev(Left(txtScript.Text, sr.StartPos), "/*") > InStr(Left(txtScript.Text, sr.StartPos), "//"))) Then
                DoSwap = Not ((WordCount(Left(txtScript.Text, sr.StartPos), "/*") + WordCount(Left(txtScript.Text, sr.StartPos), "*/")) Mod 2 = 0)
            ElseIf (InStr(Left(txtScript.Text, sr.StartPos), "/*") > InStrRev(Left(txtScript.Text, sr.StartPos), "*/")) _
                And Not ((InStrRev(Left(txtScript.Text, sr.StartPos), "/*") > InStr(Left(txtScript.Text, sr.StartPos), "//"))) Then
                DoSwap = Not ((WordCount(Left(txtScript.Text, sr.StartPos), "*/") + WordCount(Left(txtScript.Text, sr.StartPos), "/*")) Mod 2 = 0)
            ElseIf (InStrRev(Left(txtScript.Text, sr.StartPos), "//") = 0) Then

            Else
                Find = InStrRev(Left(txtScript.Text, sr.StartPos), "/*")
                If (Find > 0) And DoSwap Then
                    Find = WordCount(Left(txtScript.Text, Find - 1), "'") + WordCount(Left(txtScript.Text, Find - 1), """")
                    If ((Find > 0) And (Find Mod 2 = 0)) Or (Find = 0) Then
                        DoSwap = False
                    End If
                End If
                Find = InStrRev(Left(txtScript.Text, sr.StartPos), "*/")
                If (Find > 0) And Not DoSwap Then
                    Find = WordCount(Left(txtScript.Text, Find - 1), "/*") + WordCount(Left(txtScript.Text, Find - 1), "*/")
                    If ((Find > 0) And (Find Mod 2 = 1)) Then
                        DoSwap = False
                    End If
                End If
            End If
        End If
    ElseIf Language = "Defined" Then
        Dim lines As Long
        lines = xCountLength.StartPos ' SendMessageLngPtr(txtScript.hWnd, EM_GETLINECOUNT, 0&, 0&)
        If UBound(xDefined) <> lines Then ReDim Preserve xDefined(0 To lines) As LineDef
    End If

    lapse = Timer
    Do Until (FullText.Size = 0) Or (ByLine And (line > stopLine))
    
        If Not ByLine Then ni.Concat Header

        If ByLine And (line = startLine) Then
            sr.StartPos = fullSize - FullText.Size
        End If

        If InStr(FullText.GetString, vbCrLf) > 0 Then
            LineText.Concat FullText.GetString(vbCrLf, True), True
        Else
            LineText.Concat FullText.GetString(, True), True
        End If

        If ByLine And (line = stopLine) Then
            sr.StopPos = fullSize - FullText.Size
        End If

        If ByLine And (line = startLine) Then
            ni.Reset
            ni.Concat "{\rtf1\ansi\deff0\fs" + Trim(CInt(TextHeight))
            ni.Concat CreateFontTable
            ni.Concat CreateColorTable
            ni.Concat Header
        ElseIf ByLine And ((line > startLine) And (line <= stopLine)) Then
            ni.Concat Header
        End If

        If Not (LineText.Size = 0) Then
            If (Not ByLine) Or (ByLine And ((line >= startLine) And (line <= stopLine))) Then

                If (line = xErrorLine) Then
                    ni.Concat "\cf"
                    ni.Concat num_Error
                    ni.Concat " "
                    ni.Concat Replace(Replace(Replace(LineText.GetString, "\", "\\"), "{", "\'7b"), "}", "\'7d")
                Else
                
                
                    Select Case xLanguage
                        Case "VBScript"
                            VBS_CreateScriptLine LineText, ni
                        Case "JScript"
                            JS_CreateScriptLine LineText, ni, DoSwap
                        Case "PlainText", "Unknown"
                            TXT_CreateScriptLine LineText, ni
                        Case "Rainbow", "Dream"
                            BOW_CreateScriptLine LineText, ni, line
                        Case "Defined", "Custom"
                            DEF_CreateScriptLine LineText, ni, line
                    End Select
                End If
            End If

        End If

        If Not ByLine Then ni.Concat "\line " & vbCrLf

        If ByLine And (line = (stopLine - 1)) Then
            ni.Concat "\line\fs" + Trim(CInt(TextHeight)) + " }" & vbCrLf

            SetTextRTF ni.GetString(), sr
            ni.Reset
        
        ElseIf ByLine And ((line >= startLine) And (line < stopLine)) Then
            ni.Concat "\line " & vbCrLf
        End If
        
        line = line + 1
    Loop

    If Not ByLine Then
        ni.Concat "\line\fs" + Trim(CInt(TextHeight)) + " }" & vbCrLf

        SetTextRTF ni.GetString, sr

    End If

    ni.Reset
    Set ni = Nothing

    stacks = stacks - 1
Exit Sub
catch:
    Err.Clear
    stacks = stacks - 1
End Sub

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

Friend Sub ModifyIndent(ByVal Add As Boolean, ByVal AsComment As Boolean)

    If Not Locked Then
        
        Dim startLine As Long
        Dim stopLine As Long
        
        Dim line As Long
        Dim UseChar As String
        Dim nPos As Long
        
        Dim allText As String
        Dim LineText As String
        
        Dim sr As RangeType
        Dim ni As New Strand
                
        If AsComment Then
            If xLanguage = "VBScript" Then
                UseChar = "'"
            ElseIf xLanguage = "JScript" Then
                UseChar = "//"
            ElseIf xLanguage = "Defined" Then
                If GetFileExt(xFileName, True, True) = "nsh" Or GetFileExt(xFileName, True, True) = "nsi" _
                    Or (InStr(txtScript.Text, "SectionEnd", VbCompareMethod.vbTextCompare) > 0) _
                    Or (InStr(txtScript.Text, "!macro", VbCompareMethod.vbTextCompare) > 0) Then
                    UseChar = ";"
                ElseIf GetFileExt(xFileName, True, True) = "bat" Or _
                    (InStr(txtScript.Text, "rem", VbCompareMethod.vbTextCompare) > 0) _
                    Or (InStr(txtScript.Text, "cd ", VbCompareMethod.vbTextCompare) > 0) Then
                
                    UseChar = "rem "
                    
                End If
            End If
        Else
            UseChar = vbTab
        End If

        sr.StartPos = txtScript.SelStart
        sr.StopPos = txtScript.SelStart + txtScript.SelLength
        
        allText = txtScript.SelText
        nPos = txtScript.SelLength
        
        line = GetLineNumber(txtScript.hwnd, sr.StartPos) + 1
        If line < 1 Then line = 1
        
        startLine = line
        stopLine = GetLineNumber(txtScript.hwnd, sr.StopPos) + 1
        If stopLine < startLine Then stopLine = startLine
        
        Do Until (allText = "") Or (line > stopLine)
        
            LineText = RemoveNextArg(allText, vbCrLf, False)
    
            If (line >= startLine) And (line <= stopLine) Then

                If (line <= UBound(xDefined)) Then
                    If xDefined(line).LineClr >= 0 Then xDefined(line).LineClr = -100
                End If
                
                If Add Then
                    ni.Concat UseChar & LineText
                Else
                
                    If Not AsComment Then
                    
                        If Left(LineText, Len(UseChar)) = UseChar Then
                            ni.Concat Mid(LineText, Len(UseChar) + 1)
                        Else
                            ni.Concat LineText
                        End If
                    Else
                    
                        nPos = InStr(1, LineText, UseChar)
                        If nPos > 0 Then
                            If InStr(1, LineText, """") > 0 Then
                                If InStr(1, LineText, """") < nPos Then
                                    nPos = 0
                                End If
                            ElseIf InStr(1, LineText, "'") > 0 Then
                                If InStr(1, LineText, "'") < nPos Then
                                    nPos = 0
                                End If
                            End If
                        End If
                        
                        If nPos > 0 Then
                            ni.Concat Left(LineText, nPos - 1) & Mid(LineText, nPos + Len(UseChar))
                        Else
                            ni.Concat LineText
                        End If
                    End If
                    
                End If
    
            End If
            
            line = line + 1
        
        Loop
        
        txtScript.SelText = ni.GetString

        AddUndo
        
        RefreshView
        
        RaiseEvent Change
    End If

End Sub

Private Function IsAlphaNumeric(ByVal Text As String) As Boolean
    Dim cnt As Integer
    Dim C2 As Integer
    Dim retVal As Boolean
    retVal = True
    If Len(Text) > 0 Then
        For cnt = 1 To Len(Text)
            If (Asc(LCase(Mid(Text, cnt, 1))) = 46) Then
                C2 = C2 + 1
            ElseIf (Not IsNumeric(Mid(Text, cnt, 1))) And (Not (Asc(LCase(Mid(Text, cnt, 1))) >= 97 And Asc(LCase(Mid(Text, cnt, 1))) <= 122)) Then
                retVal = False
                Exit For
            End If
        Next
    Else
        retVal = False
    End If
    IsAlphaNumeric = retVal And (C2 <= 1)
End Function

Private Sub DEF_CreateScriptLine(ByRef LineText As Strand, ByRef doneLine As Strand, ByVal line As Long)
    If xDefined(line).LineClr < 0 Then
        Dim TempText As String
        Dim TempIndex As Long
        TempText = LineText.GetString()
        TempIndex = IIf(xDefined(line).LineClr = -100, 0, -xDefined(line).LineClr)
        RaiseEvent QueryDefine(line, TempText, TempIndex)
        If Not TempIndex >= 0 Then TempIndex = 0
        xDefined(line).LineClr = TempIndex
        TempText = RemoveNextArg(RemoveNextArg(RemoveNextArg(TempText, vbCrLf), vbCr), vbLf)
        If Not (TempText = LineText.GetString()) Then LineText.Concat TempText, True
    End If
    Select Case xDefined(line).LineClr
        Case ColoringIndexes.VariableIndex
            doneLine.Concat "\cf" & num_Variable & " "
            doneLine.Concat Replace(Replace(Replace(LineText.GetString, "\", "\\"), "{", "\'7b"), "}", "\'7d")
        Case ColoringIndexes.StatementIndex
            doneLine.Concat "\cf" & num_Statement & " "
            doneLine.Concat Replace(Replace(Replace(LineText.GetString, "\", "\\"), "{", "\'7b"), "}", "\'7d")
        Case ColoringIndexes.ValueIndex
            doneLine.Concat "\cf" & num_Value & " "
            doneLine.Concat Replace(Replace(Replace(LineText.GetString, "\", "\\"), "{", "\'7b"), "}", "\'7d")
        Case ColoringIndexes.OperatorIndex
            doneLine.Concat "\cf" & num_Operator & " "
            doneLine.Concat Replace(Replace(Replace(LineText.GetString, "\", "\\"), "{", "\'7b"), "}", "\'7d")
        Case ColoringIndexes.CommentIndex
            doneLine.Concat "\cf" & num_Comment & " "
            doneLine.Concat Replace(Replace(Replace(LineText.GetString, "\", "\\"), "{", "\'7b"), "}", "\'7d")
        Case ColoringIndexes.ErrorIndex
            doneLine.Concat "\cf" & num_Error & " "
            doneLine.Concat Replace(Replace(Replace(LineText.GetString, "\", "\\"), "{", "\'7b"), "}", "\'7d")
        Case ColoringIndexes.Dream1Index
            doneLine.Concat "\cf" & num_Dream1 & " "
            doneLine.Concat Replace(Replace(Replace(LineText.GetString, "\", "\\"), "{", "\'7b"), "}", "\'7d")
        Case ColoringIndexes.Dream2Index
            doneLine.Concat "\cf" & num_Dream2 & " "
            doneLine.Concat Replace(Replace(Replace(LineText.GetString, "\", "\\"), "{", "\'7b"), "}", "\'7d")
        Case ColoringIndexes.Dream3Index
            doneLine.Concat "\cf" & num_Dream3 & " "
            doneLine.Concat Replace(Replace(Replace(LineText.GetString, "\", "\\"), "{", "\'7b"), "}", "\'7d")
        Case ColoringIndexes.Dream4Index
            doneLine.Concat "\cf" & num_Dream4 & " "
            doneLine.Concat Replace(Replace(Replace(LineText.GetString, "\", "\\"), "{", "\'7b"), "}", "\'7d")
        Case ColoringIndexes.Dream5Index
            doneLine.Concat "\cf" & num_Dream5 & " "
            doneLine.Concat Replace(Replace(Replace(LineText.GetString, "\", "\\"), "{", "\'7b"), "}", "\'7d")
        Case ColoringIndexes.Dream6Index
            doneLine.Concat "\cf" & num_Dream6 & " "
            doneLine.Concat Replace(Replace(Replace(LineText.GetString, "\", "\\"), "{", "\'7b"), "}", "\'7d")
        Case Else
            doneLine.Concat "\cf" & num_Text & " "
            doneLine.Concat Replace(Replace(Replace(LineText.GetString, "\", "\\"), "{", "\'7b"), "}", "\'7d")
    End Select
    
End Sub

Private Sub BOW_CreateScriptLine(ByRef LineText As Strand, ByRef doneLine As Strand, ByVal line As Long)
    Do While line > 6
       line = line - 6
    Loop
    Select Case line
        Case 1
            doneLine.Concat "\cf" & num_Dream1 & " "
            doneLine.Concat Replace(Replace(Replace(LineText.GetString, "\", "\\"), "{", "\'7b"), "}", "\'7d")
        Case 2
            doneLine.Concat "\cf" & num_Dream2 & " "
            doneLine.Concat Replace(Replace(Replace(LineText.GetString, "\", "\\"), "{", "\'7b"), "}", "\'7d")
        Case 3
            doneLine.Concat "\cf" & num_Dream3 & " "
            doneLine.Concat Replace(Replace(Replace(LineText.GetString, "\", "\\"), "{", "\'7b"), "}", "\'7d")
        Case 4
            doneLine.Concat "\cf" & num_Dream4 & " "
            doneLine.Concat Replace(Replace(Replace(LineText.GetString, "\", "\\"), "{", "\'7b"), "}", "\'7d")
        Case 5
            doneLine.Concat "\cf" & num_Dream5 & " "
            doneLine.Concat Replace(Replace(Replace(LineText.GetString, "\", "\\"), "{", "\'7b"), "}", "\'7d")
        Case 6
            doneLine.Concat "\cf" & num_Dream6 & " "
            doneLine.Concat Replace(Replace(Replace(LineText.GetString, "\", "\\"), "{", "\'7b"), "}", "\'7d")
    End Select
    
End Sub

Private Sub TXT_CreateScriptLine(ByRef LineText As Strand, ByRef doneLine As Strand)

    doneLine.Concat "\cf" & num_Text & " "
    doneLine.Concat Replace(Replace(Replace(LineText.GetString, "\", "\\"), "{", "\'7b"), "}", "\'7d")
    
End Sub


Private Function JS_IsValue(ByVal QStr As String) As Boolean
    If IsNumeric(QStr) Then
        JS_IsValue = True
    ElseIf UCase(Left(QStr, 2)) = "0X" And IsNumeric(Mid(QStr, 3)) Then
        JS_IsValue = True
    Else
        JS_IsValue = False
    End If
End Function

Private Function JS_IsQuote(ByVal QStr As String) As Boolean
    Dim isRet As Boolean
    isRet = False
    Select Case QStr
        Case "'", """"
            isRet = True
    End Select
    JS_IsQuote = isRet
End Function

Private Function JS_IsEscapeChar(ByVal Char As String) As String
    Dim retChar As String
    retChar = ""
    Select Case Char
        Case "\"
            retChar = "\\"
        Case "{"
            retChar = "\'7b"
        Case "}"
            retChar = "\'7d"
    End Select
    JS_IsEscapeChar = retChar
End Function

Private Function JS_IsOperator(ByVal Opstr As String) As Boolean
    Select Case Opstr
        Case "{", "}", "[", "]", "(", ")", ";", ":", "?", "&", "|", "/", "\", ",", "=", "!", "<", ">", "+", "-", "*", "^"
            JS_IsOperator = True
        Case Else
            JS_IsOperator = False
    End Select
End Function

Private Function JS_IsStatement(ByVal Word As String, ByRef inBi As Boolean) As String
    Dim isRet As String
    isRet = ""
    Select Case Trim(Word)
        Case "const"
            isRet = "const"
        Case "block"
            isRet = "block"
        Case "null"
            isRet = "null"
        Case "true"
            isRet = "true"
        Case "false"
            isRet = "false"
        Case "break"
            isRet = "break"
        Case "catch"
            isRet = "catch"
        Case "continue"
            isRet = "continue"
        Case "debugger"
            isRet = "debugger"
        Case "do"
            isRet = "do"
        Case "while"
            isRet = "while"
        Case "for"
            isRet = "for"
        Case "each"
            isRet = "each"
        Case "export"
            isRet = "export"
        Case "in"
            isRet = "in"
        Case "of"
            isRet = "of"
        Case "function"
            isRet = "function"
        Case "if"
            isRet = "if"
        Case "else"
            isRet = "else"
        Case "return"
            isRet = "return"
        Case "switch"
            isRet = "switch"
        Case "throw"
            isRet = "throw"
        Case "try"
            isRet = "try"
        Case "var"
            isRet = "var"
        Case "while"
            isRet = "while"
        Case "with"
            isRet = "with"
        Case "else"
            isRet = "else"
        Case "import"
            isRet = "import"
        Case "label"
            isRet = "label"
        Case "let"
            isRet = "let"
        Case "new"
            isRet = "new"
        Case "case"
            isRet = "case"
        Case "default"
            isRet = "default"
    End Select
    JS_IsStatement = isRet
End Function


Private Sub JS_CreateScriptLine(ByRef LineText As Strand, ByRef doneLine As Strand, Optional ByRef DoSwap As Boolean)

    Dim nextVal As New Strand
    
    Dim tChar As String
    Dim tChar2 As String
    Dim tChar3 As String
    Dim inComment As Boolean
    Dim inQuote As Integer
    Dim inBi As Boolean
    Dim sRet As String
    Dim tPos As Long
    
    tPos = 0
    inComment = DoSwap
    inQuote = 0

    Dim tLen As Long
    Dim LineArray() As Byte
    tLen = LineText.Size
    ReDim LineArray(0 To tLen) As Byte
    LineArray = StrConv(LineText.GetString, vbFromUnicode)
    tLen = UBound(LineArray(), 1)
    
    Do
        
        tChar = Chr(LineArray(tPos))
        nextVal.Concat tChar
    
        If Not inComment And inQuote = 0 Then
            
            If tChar = "/" Then
            
                If tPos < tLen Then
                    tChar2 = Chr(LineArray(tPos + 1))
                Else
                    tChar2 = ""
                End If
                
                If tChar2 = "/" Or tChar2 = "*" Then
                    inComment = True
                    DoSwap = (tChar2 = "*")
                End If
            Else
                    
                If JS_IsOperator(tChar) Then
                    inBi = False
                    
                    doneLine.Concat "\cf" & num_Operator & " "
                    sRet = JS_IsEscapeChar(tChar)
                    If sRet = "" Then sRet = nextVal.GetString
                    doneLine.Concat sRet
                    nextVal.Reset
                Else
                    If JS_IsQuote(tChar) Then
                        If tChar = "'" Then
                            inQuote = 1
                            sRet = Left(nextVal.GetString, nextVal.Size - 1) + "\lquote "
                            nextVal.Concat sRet, True
                        End If
                        If tChar = Chr(34) Then inQuote = 2
                    Else
                    
                        If tPos < tLen Then
                            tChar2 = Chr(LineArray(tPos + 1))
                        Else
                            tChar2 = ""
                        End If
                        
                        If tChar = " " Or tChar = Chr(9) Or JS_IsQuote(tChar2) Or JS_IsOperator(tChar2) Then
                            sRet = JS_IsStatement(nextVal.GetString, inBi)
                            If sRet <> "" And Not inBi Then
                                inBi = True
                                doneLine.Concat "\cf" & num_Statement & " "
                                doneLine.Concat sRet & Replace(nextVal.GetString, sRet, "", 1, 1, vbTextCompare)
                                nextVal.Reset
                            Else
                                inBi = False
                                
                                If (JS_IsValue(nextVal.GetString)) Then
                                    doneLine.Concat "\cf" & num_Value & " "
                                    doneLine.Concat nextVal.GetString
                                    nextVal.Reset
                                Else
                                    doneLine.Concat "\cf" & num_Variable & " "
                                    doneLine.Concat nextVal.GetString
                                    nextVal.Reset
                                End If
                            End If
                           
                        End If
                    End If
                End If
            End If
        
        Else
            If DoSwap And inComment And tChar = "*" Then
                If tPos < tLen Then
                    tChar2 = Chr(LineArray(tPos + 1))
                Else
                    tChar2 = ""
                End If
                
                If tChar2 = "/" Then
                    inComment = False
                    
                    doneLine.Concat "\cf" & num_Comment & " "
                    doneLine.Concat nextVal.GetString & "/"
                    nextVal.Reset
                    tPos = tPos + 1
                    
                End If
                
            Else
                        
                If JS_IsQuote(tChar) And (inQuote > 0) Then
                
                    If tPos > 1 Then
                        tChar2 = Chr(LineArray(tPos - 1))
                    Else
                        tChar2 = ""
                    End If
                    
                    If tPos > 2 Then
                        tChar3 = Chr(LineArray(tPos - 2))
                    Else
                        tChar3 = ""
                    End If
                    
                    If (tChar = "'" And inQuote = 1) And ((Not tChar2 = "\") Or ((tChar3 = "\") And (tChar2 = "\"))) Then
                        inQuote = 0
                        sRet = Left(nextVal.GetString, nextVal.Size - 1) + "\rquote "
                        doneLine.Concat "\cf" & num_Value & " "
                        doneLine.Concat sRet
                        nextVal.Reset
                    Else
                            
                        If (tChar = """" And inQuote = 2) And ((Not tChar2 = "\") Or ((tChar3 = "\") And (tChar2 = "\"))) Then
                            inQuote = 0
                            doneLine.Concat "\cf" & num_Value & " "
                            doneLine.Concat nextVal.GetString
                            nextVal.Reset
                        Else
                            sRet = JS_IsEscapeChar(tChar)
                            If sRet <> "" Then
                                sRet = Left(nextVal.GetString, nextVal.Size - 1) + sRet
                                nextVal.Concat sRet, True
                            End If
                        End If
                    End If
                       
                Else
                    sRet = JS_IsEscapeChar(tChar)
                    If sRet <> "" Then
                        sRet = Left(nextVal.GetString, nextVal.Size - 1) + sRet
                        nextVal.Concat sRet, True
                    End If
                                
                End If

            End If
            
        End If

        tPos = tPos + 1

    Loop Until tPos > tLen
    
    If inComment = False Then DoSwap = False
    
    If nextVal.GetString <> "" Then
        If inComment Then
            doneLine.Concat "\cf" & num_Comment & " "
            doneLine.Concat nextVal.GetString
        Else
            Dim DoBi As Boolean
            DoBi = inBi
            
            sRet = JS_IsStatement(nextVal.GetString, inBi)

            If (sRet <> "" And Not inBi) And ((Not Len(sRet) <= 2) Or DoBi) Then
                doneLine.Concat "\cf" & num_Statement & " "
                doneLine.Concat sRet
            ElseIf (JS_IsValue(nextVal.GetString)) Then
                doneLine.Concat "\cf" & num_Value & " "
                doneLine.Concat nextVal.GetString
            Else
                doneLine.Concat "\cf" & num_Variable & " "
                doneLine.Concat nextVal.GetString

            End If
        End If
    End If
    
    nextVal.Reset
    Set nextVal = Nothing
End Sub


Private Sub VBS_CreateScriptLine(ByRef LineText As Strand, ByRef doneLine As Strand)

    Dim nextVal As New Strand
    
    Dim tChar As String
    Dim tChar2 As String
    Dim tPos As Long

    Dim inComment As Boolean
    Dim inQuote As Integer
    Dim inBi As Boolean
    Dim toBi As Boolean
    
    Dim sRet As String
    Dim sTmp As String
        
    Dim tLen As Long
    Dim LineArray() As Byte
    
    tLen = LineText.Size
    ReDim LineArray(0 To tLen) As Byte
    LineArray = StrConv(LineText.GetString, vbFromUnicode)
    tLen = UBound(LineArray(), 1)

    tPos = 0
    inComment = False
    inQuote = 0

    Do

        tChar = Chr(LineArray(tPos))
        nextVal.Concat tChar

        If (Not inComment) And inQuote = 0 Then

            If tChar = "'" Then
            
                inComment = True
                
                If nextVal.Size > 0 Then
                    sRet = Left(nextVal.GetString, nextVal.Size - 1)
                End If
                
                If Not (sRet = "") Then
                    doneLine.Concat sRet
                End If
                
                nextVal.Concat "'", True

            Else

                If VBS_IsOperator(tChar) Then
                    inBi = False
                    
                    doneLine.Concat "\cf" & num_Operator & " "
                    tChar = VBS_IsEscapeChar(tChar)
                    If tChar <> "" Then
                        doneLine.Concat tChar
                    Else
                        doneLine.Concat nextVal.GetString
                    End If
                    nextVal.Reset

                Else

                    If VBS_IsQuote(tChar) Then

                        If tChar = """" Then inQuote = 2

                    Else
                    
                        If tPos < tLen Then
                            tChar2 = Chr(LineArray(tPos + 1))
                        Else
                            tChar2 = ""
                        End If
                            
                        If (tChar = " " Or tChar = Chr(9)) Or VBS_IsQuote(tChar2) Or VBS_IsOperator(tChar2) Then

                            sRet = VBS_IsStatement(nextVal.GetString, inBi)
        
                            If ((sRet <> "") And (Not inBi)) Then
                                inBi = True
                                doneLine.Concat "\cf" & num_Statement & " "
                                doneLine.Concat sRet & Replace(nextVal.GetString, sRet, "", 1, 1, vbTextCompare)
                                nextVal.Reset
                            Else
                                inBi = False
                                If (VBS_IsValue(nextVal.GetString)) Then
                                    doneLine.Concat "\cf" & num_Value & " "
                                    doneLine.Concat nextVal.GetString
                                    nextVal.Reset
                                Else
                                    doneLine.Concat "\cf" & num_Variable & " "
                                    doneLine.Concat nextVal.GetString
                                    nextVal.Reset
                                End If
                            End If

                        End If

                    End If
                End If
            End If

        Else

            If VBS_IsQuote(tChar) And inQuote > 0 Then
                If tChar = """" And inQuote = 2 Then
                    inQuote = 0
                    doneLine.Concat "\cf" & num_Value & " "
                    doneLine.Concat nextVal.GetString
                    nextVal.Reset
                Else
                    sRet = VBS_IsEscapeChar(tChar)
                    If sRet <> "" Then
                        sRet = Left(nextVal.GetString, nextVal.Size - 1) & sRet
                        nextVal.Concat sRet, True
                    End If
                End If
            Else
                sRet = VBS_IsEscapeChar(tChar)
                If sRet <> "" Then
                    sRet = Left(nextVal.GetString, nextVal.Size - 1) & sRet
                    nextVal.Concat sRet, True
                End If
            End If

        End If

        tPos = tPos + 1

    Loop Until tPos > tLen
     
    If inComment Then
        doneLine.Concat "\cf" & num_Comment & " "
        doneLine.Concat nextVal.GetString
    ElseIf nextVal.GetString <> "" Then
        Dim DoBi As Boolean
        DoBi = inBi

        tChar = VBS_IsStatement(nextVal.GetString, inBi)
        
        If ((tChar <> "" And Not inBi) And DoBi) Or ((tChar <> "") And _
        (Not Right(Left(txtScript.Text, txtScript.SelStart), Len(nextVal.GetString)) _
                        = Right(LCase(LineText.GetString), Len(nextVal.GetString)))) Then

            doneLine.Concat "\cf" & num_Statement & " "
            doneLine.Concat tChar
        Else
            If (VBS_IsValue(nextVal.GetString)) Then
                doneLine.Concat "\cf" & num_Value & " "
                doneLine.Concat nextVal.GetString
            Else
                doneLine.Concat "\cf" & num_Variable & " "
                doneLine.Concat nextVal.GetString
            End If
        End If

    End If
    
    nextVal.Reset
    Set nextVal = Nothing
    
End Sub

Private Function SetTextRTF(ByVal newText As String, ByRef nRange As RangeType)
        
    If Not Cancel Then
            
        DisableRedraw
        If IIf((nRange.StartPos = -1) And (nRange.StopPos = 0), True, False) Then
            txtScript.TextRTF = newText
        Else
            
            SendMessageStruct txtScript.hwnd, EM_EXSETSEL, 0, nRange
            txtScript.SelRTF = newText
        End If

        EnableRedraw

    End If

End Function

Private Function VBS_IsQuote(ByVal QStr As String) As Boolean
    Select Case QStr
        Case Chr(34)
            VBS_IsQuote = True
        Case Else
            VBS_IsQuote = False
    End Select
End Function

Private Function VBS_IsValue(ByVal QStr As String) As Boolean
    If IsNumeric(QStr) Then
        VBS_IsValue = True
    ElseIf UCase(Left(QStr, 1)) = "H" And IsNumeric(Mid(QStr, 2)) Then
        VBS_IsValue = True
    ElseIf QStr = "_" Then
        VBS_IsValue = True
    Else
        VBS_IsValue = False
    End If
End Function

Private Function VBS_IsEscapeChar(ByVal Char As String) As String

    Dim retChar As String
    retChar = ""
    Select Case Char
        Case "\"
            retChar = "\\"
        Case "{"
            retChar = "\'7b"
        Case "}"
            retChar = "\'7d"
    End Select
    VBS_IsEscapeChar = retChar

End Function

Private Function VBS_IsOperator(ByVal Opstr As String) As Boolean
    Select Case Opstr
        Case "{", "}", "[", "]", "(", ")", ";", "&", "|", "/", "\", ",", "=", "!", "<", ">", "+", "-", "*", "^"
            VBS_IsOperator = True
        Case Else
            VBS_IsOperator = False
    End Select
End Function

Private Function VBS_IsStatement(ByVal Word As String, ByRef inBi As Boolean) As String
    
    Dim isRet As String
    isRet = ""
    Select Case LCase(Trim(Word))
        Case "type", "nothing", "const", "do", "for", "next", "each", "function", "if", "error", "compare", "explicit", "property", "preserve", "let", "get", "set", "select", "case", "set", "sub", "with", "while", "until", "declare", "integer", "string", "long", "double", "binary", "boolean", "not", "true", "false", "resume", "goto"
            isRet = Trim(Word)
            If inBi Then inBi = False
        Case "to", "call", "class", "dim", "loop", "end", "erase", "exit", "in", "then", "elseif", "else", "on", "option", "private", "public", "friend", "randomize", "redim", "rem", "wend", "as", "bit", "and", "eqv", "imp", "is", "mod", "or", "xor", "Nor", "byval", "byref", "new", "clng", "csng", "cbool", "cint", "cstr"
            isRet = Trim(Word)
    End Select

    VBS_IsStatement = isRet

End Function



