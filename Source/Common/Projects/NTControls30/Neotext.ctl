VERSION 5.00
Begin VB.UserControl Neotext 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4155
   ClipControls    =   0   'False
   MouseIcon       =   "Neotext.ctx":0000
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   3330
   ScaleWidth      =   4155
   ToolboxBitmap   =   "Neotext.ctx":0152
   Begin NTControls30.ScrollBar ScrollBar2 
      Height          =   345
      Left            =   285
      Top             =   2655
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   556
      Orientation     =   1
      AutoRedraw      =   0   'False
   End
   Begin NTControls30.ScrollBar ScrollBar1 
      Height          =   2655
      Left            =   3510
      Top             =   270
      Width           =   330
      _ExtentX        =   953
      _ExtentY        =   4683
      Orientation     =   0
      AutoRedraw      =   0   'False
   End
   Begin VB.Timer Timer1 
      Left            =   810
      Top             =   1020
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

Public Event ColorText(ByVal ViewOffset As Long, ByVal ViewWidth As Long)

Private dragStart As Long
Private keySpeed As Long
Private hasFocus As Boolean
Private insertMode As Boolean
'Private firstRun As Boolean
Private colorOpen As Boolean
Private ircColors As Boolean

Private pOffsetX As Long
Private pOffsetY As Long

Private pEnabled As Boolean
Private pLocked As Boolean
Private pHideSelection As Boolean
Private pScrollToCaret As Boolean
Private pMultiLine As Boolean
Private pLineNumbers As Boolean
Private pForecolors() As ColorRange
Private pBackcolors() As ColorRange
Private pForecolor As OLE_COLOR
Private pBackcolor As OLE_COLOR
Private pScrollBars As vbScrollBars
Private pTabSpace As String
Private pCodePage As Long
Private pPageBreaks() As Long
Private pColorRecords() As ColorRange

Private pOldProc As Long

Private aText() As NTNodes10.Strands 'codepage wise of pText
Private tText As NTNodes10.Strands 'temporary (or visible) text
Private dText As NTNodes10.Strands 'temporary (dragging) text
Private pBackBuffer As Backbuffer

Private xCancel As Long 'hold the cancel stack, for every set true, must also occur the set false

Private xUndoLimit As Long
Private xUndoActs() As UndoType
Private xUndoStage As Long
Private xUndoBuffer As Long

Private pLastSel As RangeType
Private pSel As RangeType 'where the current selection is held at all states or set

Public Function FindText(ByVal Text As String, Optional ByVal Offset As Long = 0, Optional ByVal Width As Long = -1) As Long
    Dim idx As Long
    Dim cnt As Long
    Dim cnt2 As Long
    Dim Max As Long
    FindText = -1
    If Text <> "" Then
        If Width = -1 Then Width = pText.Length - Offset
        'Max = pText.Pass(Asc(Left(Text, 1)), Offset, Width)
        'If Max > 0 Then
            cnt = 1
            Do
                idx = pText.poll(Asc(Left(Text, 1)), cnt, Offset, Width) + 1
                If Offset + idx <= Offset + Width Then
                    For cnt2 = 0 To (Len(Text) - 2)
                        If Offset + idx + cnt2 < Offset + Width Then
                            If Not (pText.Peek(Offset + idx + cnt2) = Asc(Mid(Text, cnt2 + 2, 1))) Then
                                idx = -idx
                                Exit For
                            End If
                        Else
                            idx = -idx
                            Exit For
                        End If
                    Next
                    If idx >= 0 And idx <= Offset + Width Then
                        FindText = Offset + idx - 1
                        Exit Function
                    End If
                    cnt = cnt + 1
                End If
            Loop While idx < 0 'And cnt <= Max
            
        'End If
    End If
End Function

Public Property Get ColorPalette() As String
    Dim txt As String
    Dim tmp As Long
    Dim cnt As Long
    txt = txt & " "
    For cnt = 0 To 15
        tmp = CLng(InvertNum(cnt + 1, 16))
        txt = txt & Chr(3) & IIf(tmp <= 9, "0" & CStr(tmp), CStr(tmp)) & "," & IIf(cnt <= 9, "0" & CStr(cnt), CStr(cnt)) & IIf(cnt <= 9, CStr(cnt), CStr(cnt))
    Next
    txt = txt & vbLf
    tmp = 98
    For cnt = 16 To 87
        If cnt <= 51 Then
            tmp = 98
        Else
            tmp = 89
        End If

        txt = txt & Chr(3) & IIf(tmp <= 9, "0" & CStr(tmp), CStr(tmp)) & "," & CStr(cnt) & CStr(cnt)
        If (cnt - 15) Mod 12 = 0 Then
            txt = txt & vbLf
        End If
    Next

    For cnt = 88 To 98
        
        txt = txt & Chr(3) & CStr(CLng(InvertNum(cnt - 87, 18)) + 98 - 17) & "," & CStr(cnt + 1) & CStr(cnt)
        If (cnt - 15) Mod 12 = 0 Then
            txt = txt & vbLf
        End If
    Next
    txt = txt & Chr(3) & "16,9999"

    ColorPalette = txt
End Property
Friend Property Get pText() As NTNodes10.Strands
    Set pText = aText(pCodePage)
End Property
Friend Property Set pText(ByRef RHS As NTNodes10.Strands)
    Set aText(pCodePage) = RHS
End Property

Public Property Get CodePage() As Long ' _
Gets the isolated code page number that the control is currently displaying and/or editing the text of.
    CodePage = (pCodePage + 1)
End Property
Public Property Let CodePage(ByVal RHS As Long) ' _
Sets the isolated code page number that the control is currently displaying and/or editing the text of.
    If (RHS - 1) > UBound(aText) Then
        Dim cnt As Long
        ReDim Preserve aText(0 To (RHS - 1)) As Strands
        For cnt = LBound(aText) To UBound(aText)
            If aText(cnt) Is Nothing Then Set aText(cnt) = New Strands
        Next
    End If
    pCodePage = (RHS - 1)
    UserControl_Paint
End Property

Public Property Get Seperator(ByVal Number As Long) As Long ' _
Gets the number of lines between the seperator Number and the one just before it.
    Seperator = pPageBreaks(Number - 1)
End Property
Public Property Let Seperator(ByVal Number As Long, ByVal RHS As Long) ' _
Sets the number of lines between the seperator Number and the one just before it.
    If Number >= UBound(pPageBreaks) Then
        ReDim Preserve pPageBreaks(0 To Number - 1) As Long
    End If
    pPageBreaks(Number - 1) = RHS
End Property

Public Property Get UndoLimit() As Long ' _
Gets the total number of undo entries that the control will keep track of in undo cache.
    UndoLimit = xUndoLimit
End Property
Public Property Let UndoLimit(ByVal RHS As Long) ' _
Sets the total number of undo entries that the control will keep track of in undo cache.  Setting this property resets the undo cache.
    If RHS > -2 And xUndoLimit <> RHS Then
        xUndoLimit = RHS
        ResetUndoRedo
    End If
End Property

Private Sub ResetUndoRedo()
    If UndoSize(xUndoActs) > 0 Then
        Dim cnt As Long
        For cnt = LBound(xUndoActs) To UBound(xUndoActs)
            Set xUndoActs(cnt).AfterTextData = Nothing
            Set xUndoActs(cnt).PriorTextData = Nothing
        Next
    End If
        
    ReDim Preserve xUndoActs(0 To 0) As UndoType
    xUndoStage = 0
    xUndoBuffer = xUndoLimit
    Set xUndoActs(0).PriorTextData = New NTNodes10.Strands
    Set xUndoActs(0).AfterTextData = New NTNodes10.Strands
    xUndoActs(0).PriorSelRange.StartPos = pSel.StartPos
    xUndoActs(0).PriorSelRange.StopPos = pSel.StopPos
    
End Sub

Private Sub AddUndo()

   If (xUndoLimit <> 0) Then
    
        Dim cnt As Long
        
        If ((UBound(xUndoActs) < xUndoBuffer) Or (xUndoLimit = -1)) Or (Not (xUndoStage = UBound(xUndoActs))) Then
        
            If UndoSize(xUndoActs) >= xUndoStage + 1 Then
                For cnt = (xUndoStage + 1) To UBound(xUndoActs)
                    Set xUndoActs(cnt).AfterTextData = Nothing
                    Set xUndoActs(cnt).PriorTextData = Nothing
                Next
            End If
            
            ReDim Preserve xUndoActs(0 To xUndoStage + 1) As UndoType
            Set xUndoActs(xUndoStage + 1).AfterTextData = New NTNodes10.Strands
            Set xUndoActs(xUndoStage + 1).PriorTextData = New NTNodes10.Strands
            xUndoStage = xUndoStage + 1
        ElseIf (UBound(xUndoActs) = xUndoBuffer) Or (xUndoStage = UBound(xUndoActs)) Then
            
            Set xUndoActs(LBound(xUndoActs)).AfterTextData = Nothing
            Set xUndoActs(LBound(xUndoActs)).PriorTextData = Nothing
            For cnt = (LBound(xUndoActs) + 1) To UBound(xUndoActs)
                xUndoActs(cnt - 1) = xUndoActs(cnt)
            Next
            Set xUndoActs(xUndoStage).AfterTextData = New NTNodes10.Strands
            Set xUndoActs(xUndoStage).PriorTextData = New NTNodes10.Strands
    
        End If

        xUndoActs(xUndoStage).CodePage = CodePage
        xUndoActs(xUndoStage).PriorTextData.Clone xUndoActs(0).PriorTextData
        xUndoActs(xUndoStage).AfterTextData.Clone xUndoActs(0).AfterTextData
        xUndoActs(xUndoStage).AfterSelRange = xUndoActs(0).AfterSelRange
        xUndoActs(xUndoStage).PriorSelRange = xUndoActs(0).PriorSelRange

'        Debug.Print "AddUndo Entry"
'        Debug.Print "xUndoActs(" & xUndoStage & ").CodePage=" & xUndoActs(xUndoStage).CodePage
'        Debug.Print "xUndoActs(" & xUndoStage & ").PriorTextData=" & Convert(xUndoActs(xUndoStage).PriorTextData.Partial)
'        Debug.Print "xUndoActs(" & xUndoStage & ").PriorSelRange.StartPos=" & xUndoActs(xUndoStage).PriorSelRange.StartPos
'        Debug.Print "xUndoActs(" & xUndoStage & ").PriorSelRange.StopPos=" & xUndoActs(xUndoStage).PriorSelRange.StopPos
'
'        Debug.Print "xUndoActs(" & xUndoStage & ").AfterTextData=" & Convert(xUndoActs(xUndoStage).AfterTextData.Partial)
'        Debug.Print "xUndoActs(" & xUndoStage & ").AfterSelRange.StartPos=" & xUndoActs(xUndoStage).AfterSelRange.StartPos
'        Debug.Print "xUndoActs(" & xUndoStage & ").AfterSelRange.StopPos=" & xUndoActs(xUndoStage).AfterSelRange.StopPos
    End If

End Sub

Public Function CanUndo() As Boolean
    CanUndo = ((UBound(xUndoActs) > 0) And (xUndoStage > 0)) And (Not Locked)
End Function
Public Function CanRedo() As Boolean
    CanRedo = ((xUndoStage < UBound(xUndoActs)) And (UBound(xUndoActs) > 0)) And (Not Locked)
End Function

Public Sub Undo()
   
    If CanUndo Then

        CodePage = xUndoActs(xUndoStage).CodePage
        
        With xUndoActs(xUndoStage)
            Dim tmp As Long
            Dim swapPrior As Boolean
            Dim swapAfter As Boolean
           
            xUndoActs(0).PriorSelRange = .AfterSelRange
            xUndoActs(0).AfterSelRange = .PriorSelRange
            xUndoActs(0).AfterTextData.Clone .AfterTextData
            xUndoActs(0).PriorTextData.Clone .PriorTextData
            
            If .AfterSelRange.StartPos > .AfterSelRange.StopPos Then
                swapAfter = True
                tmp = .AfterSelRange.StopPos
                .AfterSelRange.StopPos = .AfterSelRange.StartPos
                .AfterSelRange.StartPos = tmp
            End If
            
            If .PriorSelRange.StartPos > .PriorSelRange.StopPos Then
                swapPrior = True
                tmp = .PriorSelRange.StopPos
                .PriorSelRange.StopPos = .PriorSelRange.StartPos
                .PriorSelRange.StartPos = tmp
            End If
            
'            If Cancel Then
'                Debug.Print "Redo Entry"
'            Else
'                Debug.Print "Undo Entry"
'            End If
'
'            Debug.Print "xUndoActs(" & xUndoStage & ").CodePage=" & xUndoActs(xUndoStage).CodePage
'            Debug.Print "xUndoActs(" & xUndoStage & ").PriorTextData=" & Convert(xUndoActs(xUndoStage).PriorTextData.Partial)
'            Debug.Print "xUndoActs(" & xUndoStage & ").PriorSelRange.StartPos=" & xUndoActs(xUndoStage).PriorSelRange.StartPos
'            Debug.Print "xUndoActs(" & xUndoStage & ").PriorSelRange.StopPos=" & xUndoActs(xUndoStage).PriorSelRange.StopPos
'
'            Debug.Print "xUndoActs(" & xUndoStage & ").AfterTextData=" & Convert(xUndoActs(xUndoStage).AfterTextData.Partial)
'            Debug.Print "xUndoActs(" & xUndoStage & ").AfterSelRange.StartPos=" & xUndoActs(xUndoStage).AfterSelRange.StartPos
'            Debug.Print "xUndoActs(" & xUndoStage & ").AfterSelRange.StopPos=" & xUndoActs(xUndoStage).AfterSelRange.StopPos
          
            If .PriorTextData.Length > 0 Then
                If .AfterTextData.Length = 0 Then
                    pText.Pyramid .PriorTextData, IIf(.PriorSelRange.StartPos < .AfterSelRange.StartPos, .PriorSelRange.StartPos, .AfterSelRange.StartPos), _
                         IIf(.PriorSelRange.StopPos - .PriorSelRange.StartPos < .AfterSelRange.StopPos - .AfterSelRange.StartPos, _
                        .PriorSelRange.StopPos - .PriorSelRange.StartPos, .AfterSelRange.StopPos - .AfterSelRange.StartPos)
                Else
                    pText.Pinch IIf(.AfterSelRange.StartPos < .PriorSelRange.StartPos, .AfterSelRange.StartPos, .PriorSelRange.StartPos), .AfterTextData.Length
                    pText.Pyramid .PriorTextData, IIf(.AfterSelRange.StartPos < .PriorSelRange.StartPos, .AfterSelRange.StartPos, .PriorSelRange.StartPos), 0
                End If
            Else
                If .AfterSelRange.StopPos - .PriorSelRange.StartPos > 0 Then
                    pText.Pinch .PriorSelRange.StartPos, .AfterSelRange.StopPos - .PriorSelRange.StartPos
                Else
                    pText.Pinch .AfterSelRange.StartPos, .PriorSelRange.StopPos - .AfterSelRange.StartPos
                End If
            End If

            If swapAfter Then
                tmp = .AfterSelRange.StopPos
                .AfterSelRange.StopPos = .AfterSelRange.StartPos
                .AfterSelRange.StartPos = tmp
            End If
            
            If swapPrior Then
                tmp = .PriorSelRange.StopPos
                .PriorSelRange.StopPos = .PriorSelRange.StartPos
                .PriorSelRange.StartPos = tmp
            End If

            pSel = .PriorSelRange
                                 
            xUndoStage = xUndoStage - 1
            If Not Cancel Then
                RaiseEventChange False
            End If
        End With
        
    End If
End Sub

Public Sub Redo()
    
    If CanRedo Then
    
        xUndoStage = xUndoStage + 1
        Cancel = True
        
        Dim tmp1 As Strands
        Dim tmp2 As RangeType

        Set tmp1 = xUndoActs(xUndoStage).PriorTextData
        Set xUndoActs(xUndoStage).PriorTextData = xUndoActs(xUndoStage).AfterTextData
        Set xUndoActs(xUndoStage).AfterTextData = tmp1

        tmp2 = xUndoActs(xUndoStage).PriorSelRange
        xUndoActs(xUndoStage).PriorSelRange = xUndoActs(xUndoStage).AfterSelRange
        xUndoActs(xUndoStage).AfterSelRange = tmp2
        
        Undo
    
        xUndoStage = xUndoStage + 1
        
        Set tmp1 = xUndoActs(xUndoStage).PriorTextData
        Set xUndoActs(xUndoStage).PriorTextData = xUndoActs(xUndoStage).AfterTextData
        Set xUndoActs(xUndoStage).AfterTextData = tmp1

        tmp2 = xUndoActs(xUndoStage).PriorSelRange
        xUndoActs(xUndoStage).PriorSelRange = xUndoActs(xUndoStage).AfterSelRange
        xUndoActs(xUndoStage).AfterSelRange = tmp2

        Cancel = False
        RaiseEventChange False
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

Public Property Get TabSpace() As String ' _
Gets the character equivelent to a tab defined in spaces.
    TabSpace = pTabSpace
End Property
Public Property Let TabSpace(ByVal RHS As String) ' _
Sets the character equivelent to a tab defined in spaces.
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

Public Sub Cut() ' _
Preforms a removal of any selected text and puts it in the clipboard.

    If Not Locked And Enabled Then

        If SelLength > 0 Then
            Cancel = True
            
            xUndoActs(0).PriorTextData.Reset
            xUndoActs(0).AfterTextData.Reset
            
            If pSel.StartPos > pSel.StopPos Then
                Dim tmp As Long
                tmp = pSel.StartPos
                pSel.StartPos = pSel.StopPos
                pSel.StopPos = tmp
            End If
            
            DepleetColorRecords pSel.StartPos, pSel.StopPos - pSel.StartPos

            xUndoActs(0).PriorTextData.Concat Convert(Replace(SelText, vbCrLf, vbLf))
            
            Clipboard.SetText Replace(SelText, vbCrLf, vbLf)
            pText.Pinch pSel.StartPos, pSel.StopPos - pSel.StartPos
            
            pSel.StopPos = pSel.StartPos
            
            Cancel = False
            RaiseEventChange
        End If

    End If

End Sub
Public Sub Copy() ' _
Places any selected text into the clipboard.
    If Enabled Then
        If SelLength > 0 Then
            Clipboard.SetText SelText
        End If
    End If
End Sub

Public Sub Paste() ' _
Inserts into the text at the current selection any text data contained in the clipboard.
    If Not Locked And Enabled Then

        If Clipboard.GetFormat(ClipBoardConstants.vbCFText) Then
            Cancel = True

            xUndoActs(0).PriorTextData.Reset
            xUndoActs(0).AfterTextData.Reset
            
            If SelLength > 0 Then
                xUndoActs(0).PriorTextData.Concat Convert(Replace(SelText, vbCrLf, vbLf))
            End If
            
            Dim clipText() As Byte
            clipText = Convert(Replace(Clipboard.GetText(ClipBoardConstants.vbCFText), vbCrLf, vbLf))
                        
            Dim tText As New Strands
            tText.Concat clipText
            If pSel.StartPos > pSel.StopPos Then
                ExpandColorRecords pSel.StopPos, SelLength
                pText.Pyramid tText, pSel.StopPos, pSel.StartPos - pSel.StopPos
                pSel.StopPos = pSel.StopPos + (pSel.StartPos - pSel.StopPos) - ((pSel.StartPos - pSel.StopPos) - tText.Length)
                pSel.StartPos = pSel.StopPos
            Else
                ExpandColorRecords pSel.StartPos, SelLength
                pText.Pyramid tText, pSel.StartPos, pSel.StopPos - pSel.StartPos
                pSel.StartPos = pSel.StartPos + (pSel.StopPos - pSel.StartPos) - ((pSel.StopPos - pSel.StartPos) - tText.Length)
                pSel.StopPos = pSel.StartPos
            End If
            Set tText = Nothing
            
            xUndoActs(0).AfterTextData.Concat clipText
            
            Cancel = False
            RaiseEventChange
        End If

    End If
End Sub
Public Sub ClearAll() ' _
Clears all the text on the current code page.
    If Not Locked And Enabled Then
        Cancel = True

        xUndoActs(0).PriorTextData.Reset
        xUndoActs(0).AfterTextData.Reset
        
        If pText.Length > 0 Then
            xUndoActs(0).PriorTextData.Concat pText.Partial
        End If
        
        pSel.StartPos = 0
        pSel.StopPos = 0
        pText.Reset
        
        Cancel = False
        RaiseEventChange
              
    End If
End Sub
Friend Sub ClearSeperators()
    ReDim pPageBreaks(0 To 0) As Long
    pPageBreaks(0) = 0
End Sub
Private Sub ResetColors()
    ReDim pColorRecords(0 To 0) As ColorRange
    pColorRecords(0).Forecolor = pForecolor
    pColorRecords(0).BackColor = pBackcolor
End Sub
Public Sub Reset() ' _
Resets the control to the default state of properties.

    ResetColors
    
    ClearSeperators
    ResetUndoRedo

    Set tText = Nothing
    Set dText = Nothing
    Dim cnt As Long
    For cnt = LBound(aText) To UBound(aText)
        Set aText(cnt) = Nothing
    Next
    ReDim aText(0 To 0) As Strands
    Set aText(0) = New Strands

End Sub

Public Property Get LineNumbers() As Boolean ' _
Gets whehter or not the control draws line numbers on the left margin.
    LineNumbers = pLineNumbers
End Property
Public Property Let LineNumbers(ByVal RHS As Boolean) ' _
Sets whether or not the control draws line numbers on the left margin.
    If pLineNumbers <> RHS Then
        pLineNumbers = RHS
        UserControl_Paint
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
    End If
End Property

Public Property Get Backbuffer() As Backbuffer ' _
Gets the BackBuffer object associated with this control, in which it draws to first before displaying.
    Set Backbuffer = pBackBuffer
End Property
Public Property Set Backbuffer(ByRef RHS As Backbuffer) ' _
Sets the BackBuffer object associated with this control, in which it draws to first before displaying.
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
    DrawableRect = RECT(LineColumnWidth, 0, UsercontrolWidth + LineColumnWidth, UsercontrolHeight)
End Function
Private Static Property Get UsercontrolWidth(Optional ByVal Recalc As Boolean = False) As Long
    Static pUserControlWidth As Long
    If pUserControlWidth = 0 Or Recalc Then
        UsercontrolWidth = IIf(ScrollBar1.Visible, (UserControl.Width - ScrollBar1.Width), UserControl.Width) - LineColumnWidth
        If UsercontrolWidth < 0 Then UsercontrolWidth = 0
    Else
        UsercontrolWidth = pUserControlWidth
    End If
End Property

Private Static Property Get UsercontrolHeight(Optional ByVal Recalc As Boolean = False) As Long
    Static pUserControlheight As Long
    If pUserControlheight = 0 Or Recalc Then
        UsercontrolHeight = IIf(ScrollBar2.Visible, (UserControl.Height - ScrollBar2.Height), UserControl.Height)
        If UsercontrolHeight < 0 Then UsercontrolHeight = 0
    Else
        UsercontrolHeight = pUserControlheight
    End If
End Property

Private Property Get LineColumnWidth() As Long
    If pLineNumbers Then
        LineColumnWidth = Me.TextWidth("." & (LineCount + (((UsercontrolHeight \ TextHeight) + 1) \ 2)) & ".")
    Else
        LineColumnWidth = 0
    End If
End Property

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
        UserControl_Paint
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

            Dim tText As New Strands
            If pText.Length > 0 Then
            
                tText.Concat Convert(Replace(Replace(Convert(pText.Partial), vbCrLf, vbLf), vbLf, ""))
            End If
            
            If tText.Length > 0 Then
                pText.Clone tText
            Else
                pText.Reset
            End If
            Set tText = Nothing
            
            pOffsetY = 0
        End If
        UserControl_Paint
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
            tmp = pText.poll(Asc(vbLf), tmp2 + 1)
        End If
        tmp2 = pText.poll(Asc(vbLf), tmp2 + UsercontrolHeight \ TextHeight + 1)
        
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
            Dim Max As Long
            Dim tmp As Long
            For cnt = 0 To LineCount - 1
                tmp = Me.TextWidth(LineText(cnt))
                If tmp > Max Then Max = tmp
            Next
            pCanvasWidth = Max + (UsercontrolWidth / 2)
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
    End If
End Property

Private Sub CanvasValidate(Optional ByVal RecalcSizeOf As Boolean = True)

    If pOffsetX > 0 Or UsercontrolWidth(RecalcSizeOf) > GetCanvasWidth(RecalcSizeOf) Then pOffsetX = 0
    If pOffsetY > 0 Or UsercontrolHeight(RecalcSizeOf) > GetCanvasHeight(RecalcSizeOf) Then pOffsetY = 0

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
        ReDim pForecolors(0 To 0) As ColorRange
        ReDim pBackcolors(0 To 0) As ColorRange
        
        pForecolor = RHS
        UserControl_Paint
    End If
End Property

Public Property Get BackColor() As OLE_COLOR ' _
Gets the default background color of the text display when a specific color table coloring is not used.
Attribute BackColor.VB_Description = "Gets the default background color of the text display when a specific color table coloring is not used."
    BackColor = pBackcolor
End Property
Public Property Let BackColor(ByVal RHS As OLE_COLOR) ' _
Sets the default background color of the text display when a specific color table coloring is not used.
Attribute BackColor.VB_Description = "Sets the default background color of the text display when a specific color table coloring is not used."
    If pBackcolor <> RHS Then
        ReDim pForecolors(0 To 0) As ColorRange
        ReDim pBackcolors(0 To 0) As ColorRange
        pBackcolor = RHS
        UserControl.BackColor = RHS
        UserControl_Paint
    End If
End Property

Public Sub ColorText(ByVal Forecolor As Variant, Optional BackColor As Variant, Optional ByVal Offset As Long = 0, Optional ByVal Width As Long = -1) ' _
Changes the color of existing text in the control specified by the optional Offset and Width, when omitted, the entire text color is changed.
     If colorOpen Then

        Dim startClr As Long
        Dim stopClr As Long
        Dim cnt As Long
        Dim cnt2 As Long
        Dim tmpBack As Long
        Dim tmpFore As Long

        If Width = -1 Then Width = Length - Offset
        stopClr = LocateColorRecord(Offset + Width)
        tmpBack = pColorRecords(stopClr).BackColor
        tmpFore = pColorRecords(stopClr).Forecolor

        startClr = LocateColorRecord(Offset) + 1
        ReDim Preserve pColorRecords(0 To UBound(pColorRecords) + 1) As ColorRange
        For cnt = UBound(pColorRecords) - 1 To startClr Step -1
            pColorRecords(cnt + 1) = pColorRecords(cnt)
        Next

        ReDim Preserve pColorRecords(0 To UBound(pColorRecords) + 1) As ColorRange
        For cnt = UBound(pColorRecords) - 1 To startClr + 1 Step -1
            pColorRecords(cnt + 1) = pColorRecords(cnt)
        Next
        
        cnt = startClr
        
        Do While cnt <= stopClr
            For cnt2 = cnt To UBound(pColorRecords) - 1
                pColorRecords(cnt2) = pColorRecords(cnt2 + 1)
            Next
            ReDim Preserve pColorRecords(0 To UBound(pColorRecords) - 1) As ColorRange
            stopClr = stopClr - 1
        Loop
        
        stopClr = startClr + 1

        With pColorRecords(startClr)
            If Not IsMissing(Forecolor) Then
                .Forecolor = Forecolor
            Else
                .Forecolor = pForecolor
            End If
            If Not IsMissing(BackColor) Then
                .BackColor = BackColor
            Else
                .BackColor = pBackcolor
            End If
            .StartLoc = Offset
        End With

        If (Not startClr = stopClr) Then

            With pColorRecords(stopClr)
                .Forecolor = tmpFore
                .BackColor = tmpBack
                .StartLoc = Offset + Width
            End With

        End If

    Else
        Err.Raise 8, , "The ColorText function must be used with in the ColorText event and can not be used otherwise.  (IRC style color coding may be used when setting the Text property, outside of the ColorText event)"
    End If
End Sub

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

    If Len(RHS) > 0 Then
        tText.Concat Convert(Replace(RHS, vbCrLf, vbLf))
    End If
    
    If pSel.StartPos >= 0 And pText.Length > pSel.StopPos And pSel.StopPos - pSel.StartPos > 0 Then
        pText.Pinch pSel.StartPos, pSel.StopPos - pSel.StartPos
    End If
    
    If tText.Length > 0 Then
        pText.Pyramid tText, pSel.StartPos, 0
    End If
    
    Set tText = Nothing
    
    pSel.StopPos = pSel.StartPos + Len(Replace(RHS, vbCrLf, vbLf))
    Swap pSel.StartPos, pSel.StopPos
    InvalidateCursor
    MakeCaretVisible CaretLocation, True
    SetScrollBarsReverse
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
    MakeCaretVisible CaretLocation, True
    SetScrollBarsReverse
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
    MakeCaretVisible CaretLocation, True
    SetScrollBarsReverse
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
        If InStr(RHS, Chr(3)) > 0 Then ircColors = True
        ResetColors
        pText.Concat Convert(Replace(Replace(RHS, vbCrLf, vbLf), vbLf, IIf(pMultiLine, vbLf, "")))
        ResetUndoRedo
        RaiseEventChange False

    End If
    
End Property
Public Property Set Text(ByRef RHS As Strands) ' _
Sets the text contents of the control in the NTNodes10.Staands object type.
Attribute Text.VB_Description = "Sets the text contents of the control in the NTNodes10.Staands object type."
    Select Case TypeName(RHS)
        Case "Strands", "NTNodes10.Strands", "NTControls30.Strands"
            pText.Reset
            If RHS.Length > 0 Then
                If RHS.Pass(3) > 0 Then ircColors = True
                ResetColors
                pText.Concat Convert(Replace(Replace(Convert(RHS.Partial), vbCrLf, vbLf), vbLf, IIf(pMultiLine, vbLf, "")))
                ResetUndoRedo
                RaiseEventChange False
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
'    Dim lastFont As StdFont
'    Set lastFont = UserControl.Font
'    Dim widthMatch As Boolean
    
resetfont:

    Set UserControl.Font = newVal
    Set pBackBuffer.Font = UserControl.Font
    
'    If widthMatch Then
'        Err.Raise 8, , "This control only accepts fonts where the width of all of its characters are equal."
'    End If
'
'    Dim charNum As Integer
'
'    For charNum = Asc("a") To Asc("a") + 25
'        widthMatch = (TextWidth(Chr(charNum)) = TextWidth(Chr(charNum + 1)))
'        If Not widthMatch Then
'            widthMatch = True
'            Set newVal = lastFont
'            GoTo resetfont
'        End If
'    Next
'
'    For charNum = Asc("A") To Asc("a") + 25
'        widthMatch = (TextWidth(Chr(charNum)) = TextWidth(Chr(charNum + 1)))
'        If Not widthMatch Then
'            widthMatch = True
'            Set newVal = lastFont
'            GoTo resetfont
'        End If
'    Next
'
'    For charNum = Asc("0") To Asc("0") + 9
'        widthMatch = (TextWidth(Chr(charNum)) = TextWidth(Chr(charNum + 1)))
'        If Not widthMatch Then
'            widthMatch = True
'            Set newVal = lastFont
'            GoTo resetfont
'        End If
'    Next
'
'    widthMatch = widthMatch And (TextWidth("`") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth("~") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth("!") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth("@") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth("#") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth("$") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth("%") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth("^") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth("&") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth("*") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth("(") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth(")") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth("_") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth("-") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth("+") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth("=") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth("[") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth("{") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth("]") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth("}") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth("\") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth("|") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth(";") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth(":") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth("'") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth("""") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth(",") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth("<") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth(".") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth(">") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth("/") = TextWidth("a"))
'    widthMatch = widthMatch And (TextWidth("?") = TextWidth("a"))
'    If Not widthMatch Then
'        widthMatch = True
'        Set newVal = lastFont
'        GoTo resetfont
'    End If
'
'    UserControl_Paint
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
            dragStart = -3
            pScrollToCaret = False
        End If
    End If
End Sub

Private Sub ScrollBar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If dragStart = -3 Then
        dragStart = 0
        pScrollToCaret = True
    End If
End Sub

Private Sub ScrollBar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If pScrollToCaret Then
            dragStart = -3
            pScrollToCaret = False
        End If
    End If
End Sub

Private Sub ScrollBar2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If dragStart = -3 Then
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
    Dim ClrRec As Long
    Dim newloc As POINTAPI
    newloc = CaretLocation
    
    If ((Not cursorBlink) Or (Not hasFocus)) Or ((newloc.X <> lastLoc.X) Or (newloc.Y <> lastLoc.Y)) Then
        If insertMode Then
            ClrRec = LocateColorRecord(pSel.StartPos)
            ClipPrintText lastLoc.X, lastLoc.Y, InsertCharacter(Convert(pText.Partial(pSel.StartPos, 1))), pColorRecords(ClrRec).Forecolor, pColorRecords(ClrRec).BackColor, False
        Else
            ClipLineDraw lastLoc.X, lastLoc.Y, lastLoc.X, (lastLoc.Y + TextHeight), pBackcolor
        End If
    End If
    
    Static lastSel As RangeType
    If (lastSel.StartPos <> pSel.StartPos Or lastSel.StopPos <> pSel.StopPos) Then
        
        MakeCaretVisible newloc, True
        cursorBlink = False
    End If
    lastSel.StartPos = pSel.StartPos
    lastSel.StopPos = pSel.StopPos

    Paint
    
    If (((cursorBlink Or ((newloc.X <> lastLoc.X) Or (newloc.Y <> lastLoc.Y))) And hasFocus)) And Enabled Then
        If insertMode Then
            ClipPrintText newloc.X, newloc.Y, InsertCharacter(Convert(pText.Partial(pSel.StartPos, 1))), pBackcolor, pForecolor, True
        Else
            ClipLineDraw newloc.X, newloc.Y, newloc.X, (newloc.Y + TextHeight), pForecolor
        End If
        
    ElseIf ((Not cursorBlink) Or (Not hasFocus)) Or ((newloc.X <> lastLoc.X) Or (newloc.Y <> lastLoc.Y)) Then
        If insertMode Then
            ClrRec = LocateColorRecord(pSel.StartPos)
            ClipPrintText lastLoc.X, lastLoc.Y, InsertCharacter(Convert(pText.Partial(pSel.StartPos, 1))), pColorRecords(ClrRec).Forecolor, pColorRecords(ClrRec).BackColor, False
        Else
            ClipLineDraw lastLoc.X, lastLoc.Y, lastLoc.X, (lastLoc.Y + TextHeight), pBackcolor
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
Private Function MakeCaretVisible(ByRef Loc As POINTAPI, ByVal LargeJump As Boolean) As Boolean
    If Enabled Then
        If pScrollToCaret And (Not ClippingWouldDraw(DrawableRect, RECT(Loc.X, Loc.Y, Loc.X + TextWidth, Loc.Y + TextHeight), True)) Then
            If Loc.X < 1 Then
                If LargeJump Then
                    pOffsetX = ((pOffsetX + LineColumnWidth) + ((1 - Loc.X) + (UsercontrolWidth / 2)))
                Else
                    pOffsetX = (pOffsetX + ((1 - Loc.X) + ScrollBar2.SmallChange))
                End If
                If ScrollBar2.Visible And ScrollBar2.Value <> -pOffsetX Then ScrollBar2.Value = -pOffsetX
                MakeCaretVisible = True
            ElseIf Loc.X > UsercontrolWidth Or Loc.X + TextWidth > UsercontrolWidth Then
                If LargeJump Then
                    pOffsetX = (pOffsetX - ((Loc.X - UsercontrolWidth) + (UsercontrolWidth / 2)))
                Else
                    pOffsetX = (pOffsetX - ((Loc.X - UsercontrolWidth) + ScrollBar2.SmallChange))
                End If
                If ScrollBar2.Visible And ScrollBar2.Value <> -pOffsetX Then ScrollBar2.Value = -pOffsetX
                MakeCaretVisible = True
            End If
            If Loc.Y < 1 Then
                If LargeJump Then
                    pOffsetY = (pOffsetY + (((((1 - Loc.Y) + (UsercontrolHeight / 2)) \ TextHeight)) * TextHeight))
                Else
                    pOffsetY = (pOffsetY + (((((1 - Loc.Y) + ScrollBar1.SmallChange) \ TextHeight)) * TextHeight))
                End If
                If ScrollBar1.Visible And ScrollBar1.Value <> -pOffsetY Then ScrollBar1.Value = -pOffsetY
                MakeCaretVisible = True
            ElseIf Loc.Y > UsercontrolHeight Or Loc.Y + TextHeight > UsercontrolHeight Then
                If LargeJump Then
                    pOffsetY = (pOffsetY - (((((Loc.Y - UsercontrolHeight) + (UsercontrolHeight / 2)) \ TextHeight)) * TextHeight))
                Else
                    pOffsetY = (pOffsetY - (((((Loc.Y - UsercontrolHeight) + ScrollBar1.SmallChange) \ TextHeight)) * TextHeight))
                End If
                If ScrollBar1.Visible And ScrollBar1.Value <> -pOffsetY Then ScrollBar1.Value = -pOffsetY
                MakeCaretVisible = True
            End If
            
            If MakeCaretVisible = True Then
                CanvasValidate True
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

Private Function ClipPrintText(ByVal X1 As Single, ByVal Y1 As Single, ByVal StrText As String, Optional fColor As Variant, Optional bColor As Variant, Optional ByVal BoxFill As Boolean = False) As Long
    StrText = Replace(Replace(StrText, Chr(9), TabSpace), Chr(0), "")
    ClipPrintText = Me.TextWidth(StrText)
    If Not IsMissing(bColor) Then
        If bColor <> pBackcolor Then
            If ClipLineDraw(X1, Y1, (ClipPrintText + X1), (Me.TextHeight(StrText) + Y1), bColor, True) Then
                pBackBuffer.DrawText X1 / Screen.TwipsPerPixelX + 1, Y1 / Screen.TwipsPerPixelY, StrText, fColor
            'Else
            '    ClipPrintText = 0
            End If
        ElseIf ClippingWouldDraw(DrawableRect, RECT(X1, Y1, (Me.TextWidth(StrText) + X1), (Me.TextHeight(StrText) + Y1))) Then
            If ClipLineDraw(X1, Y1, (ClipPrintText + X1), (Me.TextHeight(StrText) + Y1), bColor, True) Then
                pBackBuffer.DrawText X1 / Screen.TwipsPerPixelX + 1, Y1 / Screen.TwipsPerPixelY, StrText, fColor
            'Else
            '    ClipPrintText = 0
            End If
            'pBackBuffer.DrawText X1 / Screen.TwipsPerPixelX + 1, Y1 / Screen.TwipsPerPixelY, StrText, fColor
        'Else
        '    ClipPrintText = 0
        End If
    ElseIf BoxFill Then
        If ClipLineDraw(X1, Y1, (ClipPrintText + X1), (Me.TextHeight(StrText) + Y1), bColor, True) Then
            pBackBuffer.DrawText X1 / Screen.TwipsPerPixelX + 1, Y1 / Screen.TwipsPerPixelY, StrText, fColor
        'Else
        '    ClipPrintText = 0
        End If
    ElseIf ClippingWouldDraw(DrawableRect, RECT(X1, Y1, (Me.TextWidth(StrText) + X1), (Me.TextHeight(StrText) + Y1))) Then
        If ClipLineDraw(X1, Y1, (ClipPrintText + X1), (Me.TextHeight(StrText) + Y1), bColor, True) Then
            pBackBuffer.DrawText X1 / Screen.TwipsPerPixelX + 1, Y1 / Screen.TwipsPerPixelY, StrText, fColor
        'Else
        '    ClipPrintText = 0
        End If
        'pBackBuffer.DrawText X1 / Screen.TwipsPerPixelX + 1, Y1 / Screen.TwipsPerPixelY, StrText, fColor
    'Else
    '    ClipPrintText = 0
    End If
End Function

Private Function GetRGBColor(ByVal IRCColorNum As Long, ByVal ForeElseBack As Boolean) As Long
    Select Case IRCColorNum
        Case 0 '- 00 - White.
            GetRGBColor = RGB(255, 255, 255)
        Case 1 '- 01 - Black.
            GetRGBColor = RGB(0, 0, 0)
        Case 2 '- 02 - Blue.
            GetRGBColor = RGB(0, 0, 127)
        Case 3 '- 03 - Green.
            GetRGBColor = RGB(0, 147, 0)
        Case 4 '- 04 - Red.
            GetRGBColor = RGB(255, 0, 0)
        Case 5 '- 05 - Brown.
            GetRGBColor = RGB(127, 0, 0)
        Case 6 '- 06 - Magenta.
            GetRGBColor = RGB(156, 0, 156)
        Case 7 '- 07 - Orange.
            GetRGBColor = RGB(252, 127, 0)
        Case 8 '- 08 - Yellow.
            GetRGBColor = RGB(255, 255, 0)
        Case 9 '- 09 - Light Green.
            GetRGBColor = RGB(0, 252, 0)
        Case 10 '- 10 - Cyan.
            GetRGBColor = RGB(0, 147, 147)
        Case 11 '- 11 - Light Cyan.
            GetRGBColor = RGB(0, 255, 255)
        Case 12 '- 12 - Light Blue.
            GetRGBColor = RGB(0, 0, 252)
        Case 13 '- 13 - Pink.
            GetRGBColor = RGB(255, 0, 255)
        Case 14 '- 14 - Grey.
            GetRGBColor = RGB(127, 127, 127)
        Case 15 '- 15 - Light Grey.
            GetRGBColor = RGB(210, 210, 210)

        Case 16
            GetRGBColor = RGB(Val("&H47"), Val("&H00"), Val("&H00"))
        Case 17
            GetRGBColor = RGB(Val("&H47"), Val("&H21"), Val("&H00"))
        Case 18
            GetRGBColor = RGB(Val("&H47"), Val("&H47"), Val("&H00"))
        Case 19
            GetRGBColor = RGB(Val("&H32"), Val("&H47"), Val("&H00"))
        Case 20
            GetRGBColor = RGB(Val("&H00"), Val("&H47"), Val("&H00"))
        Case 21
            GetRGBColor = RGB(Val("&H00"), Val("&H47"), Val("&H2c"))
        Case 22
            GetRGBColor = RGB(Val("&H00"), Val("&H47"), Val("&H47"))
        Case 23
            GetRGBColor = RGB(Val("&H00"), Val("&H27"), Val("&H47"))
        Case 24
            GetRGBColor = RGB(Val("&H00"), Val("&H00"), Val("&H47"))
        Case 25
            GetRGBColor = RGB(Val("&H2e"), Val("&H00"), Val("&H47"))
        Case 26
            GetRGBColor = RGB(Val("&H47"), Val("&H00"), Val("&H47"))
        Case 27
            GetRGBColor = RGB(Val("&H47"), Val("&H00"), Val("&H2a"))
        Case 28
            GetRGBColor = RGB(Val("&H74"), Val("&H00"), Val("&H00"))
        Case 29
            GetRGBColor = RGB(Val("&H74"), Val("&H3a"), Val("&H00"))
        Case 30
            GetRGBColor = RGB(Val("&H74"), Val("&H74"), Val("&H00"))
        Case 31
            GetRGBColor = RGB(Val("&H51"), Val("&H74"), Val("&H00"))
        Case 32
            GetRGBColor = RGB(Val("&H00"), Val("&H74"), Val("&H00"))
        Case 33
            GetRGBColor = RGB(Val("&H00"), Val("&H74"), Val("&H49"))
        Case 34
            GetRGBColor = RGB(Val("&H00"), Val("&H74"), Val("&H74"))
        Case 35
            GetRGBColor = RGB(Val("&H00"), Val("&H40"), Val("&H74"))
        Case 36
            GetRGBColor = RGB(Val("&H00"), Val("&H00"), Val("&H74"))
        Case 37
            GetRGBColor = RGB(Val("&H40"), Val("&H00"), Val("&H74"))
        Case 38
            GetRGBColor = RGB(Val("&H74"), Val("&H00"), Val("&H74"))
        Case 39
            GetRGBColor = RGB(Val("&H74"), Val("&H00"), Val("&H45"))
        Case 40
            GetRGBColor = RGB(Val("&Hb5"), Val("&H00"), Val("&H00"))
        Case 41
            GetRGBColor = RGB(Val("&Hb5"), Val("&H63"), Val("&H00"))
        Case 42
            GetRGBColor = RGB(Val("&Hb5"), Val("&Hb5"), Val("&H00"))
        Case 43
            GetRGBColor = RGB(Val("&H7d"), Val("&Hb5"), Val("&H00"))
        Case 44
            GetRGBColor = RGB(Val("&H00"), Val("&Hb5"), Val("&H00"))
        Case 45
            GetRGBColor = RGB(Val("&H00"), Val("&Hb5"), Val("&H71"))
        Case 46
            GetRGBColor = RGB(Val("&H00"), Val("&Hb5"), Val("&Hb5"))
        Case 47
            GetRGBColor = RGB(Val("&H00"), Val("&H63"), Val("&Hb5"))
        Case 48
            GetRGBColor = RGB(Val("&H00"), Val("&H00"), Val("&Hb5"))
        Case 49
            GetRGBColor = RGB(Val("&H75"), Val("&H00"), Val("&Hb5"))
        Case 50
            GetRGBColor = RGB(Val("&Hb5"), Val("&H00"), Val("&Hb5"))
        Case 51
            GetRGBColor = RGB(Val("&Hb5"), Val("&H00"), Val("&H6b"))
        Case 52
            GetRGBColor = RGB(Val("&Hff"), Val("&H00"), Val("&H00"))
        Case 53
            GetRGBColor = RGB(Val("&Hff"), Val("&H8c"), Val("&H00"))
        Case 54
            GetRGBColor = RGB(Val("&Hff"), Val("&Hff"), Val("&H00"))
        Case 55
            GetRGBColor = RGB(Val("&Hb2"), Val("&Hff"), Val("&H00"))
        Case 56
            GetRGBColor = RGB(Val("&H00"), Val("&Hff"), Val("&H00"))
        Case 57
            GetRGBColor = RGB(Val("&H00"), Val("&Hff"), Val("&Ha0"))
        Case 58
            GetRGBColor = RGB(Val("&H00"), Val("&Hff"), Val("&Hff"))
        Case 59
            GetRGBColor = RGB(Val("&H00"), Val("&H8c"), Val("&Hff"))
        Case 60
            GetRGBColor = RGB(Val("&H00"), Val("&H00"), Val("&Hff"))
        Case 61
            GetRGBColor = RGB(Val("&Ha5"), Val("&H00"), Val("&Hff"))
        Case 62
            GetRGBColor = RGB(Val("&Hff"), Val("&H00"), Val("&Hff"))
        Case 63
            GetRGBColor = RGB(Val("&Hff"), Val("&H00"), Val("&H98"))
        Case 64
            GetRGBColor = RGB(Val("&Hff"), Val("&H59"), Val("&H59"))
        Case 65
            GetRGBColor = RGB(Val("&Hff"), Val("&Hb4"), Val("&H59"))
        Case 66
            GetRGBColor = RGB(Val("&Hff"), Val("&Hff"), Val("&H71"))
        Case 67
            GetRGBColor = RGB(Val("&Hcf"), Val("&Hff"), Val("&H60"))
        Case 68
            GetRGBColor = RGB(Val("&H6f"), Val("&Hff"), Val("&H6f"))
        Case 69
            GetRGBColor = RGB(Val("&H65"), Val("&Hff"), Val("&Hc9"))
        Case 70
            GetRGBColor = RGB(Val("&H6d"), Val("&Hff"), Val("&Hff"))
        Case 71
            GetRGBColor = RGB(Val("&H59"), Val("&Hb4"), Val("&Hff"))
        Case 72
            GetRGBColor = RGB(Val("&H59"), Val("&H59"), Val("&Hff"))
        Case 73
            GetRGBColor = RGB(Val("&Hc4"), Val("&H59"), Val("&Hff"))
        Case 74
            GetRGBColor = RGB(Val("&Hff"), Val("&H66"), Val("&Hff"))
        Case 75
            GetRGBColor = RGB(Val("&Hff"), Val("&H59"), Val("&Hbc"))
        Case 76
            GetRGBColor = RGB(Val("&Hff"), Val("&H9c"), Val("&H9c"))
        Case 77
            GetRGBColor = RGB(Val("&Hff"), Val("&Hd3"), Val("&H9c"))
        Case 78
            GetRGBColor = RGB(Val("&Hff"), Val("&Hff"), Val("&H9c"))
        Case 79
            GetRGBColor = RGB(Val("&He2"), Val("&Hff"), Val("&H9c"))
        Case 80
            GetRGBColor = RGB(Val("&H9c"), Val("&Hff"), Val("&H9c"))
        Case 81
            GetRGBColor = RGB(Val("&H9c"), Val("&Hff"), Val("&Hdb"))
        Case 82
            GetRGBColor = RGB(Val("&H9c"), Val("&Hff"), Val("&Hff"))
        Case 83
            GetRGBColor = RGB(Val("&H9c"), Val("&Hd3"), Val("&Hff"))
        Case 84
            GetRGBColor = RGB(Val("&H9c"), Val("&H9c"), Val("&Hff"))
        Case 85
            GetRGBColor = RGB(Val("&Hdc"), Val("&H9c"), Val("&Hff"))
        Case 86
            GetRGBColor = RGB(Val("&Hff"), Val("&H9c"), Val("&Hff"))
        Case 87
            GetRGBColor = RGB(Val("&Hff"), Val("&H94"), Val("&Hd3"))
        Case 88
            GetRGBColor = RGB(Val("&H00"), Val("&H00"), Val("&H00"))
        Case 89
            GetRGBColor = RGB(Val("&H13"), Val("&H13"), Val("&H13"))
        Case 90
            GetRGBColor = RGB(Val("&H28"), Val("&H28"), Val("&H28"))
        Case 91
            GetRGBColor = RGB(Val("&H36"), Val("&H36"), Val("&H36"))
        Case 92
            GetRGBColor = RGB(Val("&H4d"), Val("&H4d"), Val("&H4d"))
        Case 93
            GetRGBColor = RGB(Val("&H65"), Val("&H65"), Val("&H65"))
        Case 94
            GetRGBColor = RGB(Val("&H81"), Val("&H81"), Val("&H81"))
        Case 95
            GetRGBColor = RGB(Val("&H9f"), Val("&H9f"), Val("&H9f"))
        Case 96
            GetRGBColor = RGB(Val("&Hbc"), Val("&Hbc"), Val("&Hbc"))
        Case 97
            GetRGBColor = RGB(Val("&He2"), Val("&He2"), Val("&He2"))
        Case 98
            GetRGBColor = RGB(Val("&Hff"), Val("&Hff"), Val("&Hff"))

        Case 99 '- 99 - Default Foreground/Background - Not universally supported.
            GetRGBColor = IIf(ForeElseBack, pForecolor, pBackcolor)
        
    End Select
End Function

Private Function GetIRCColor(ByVal RGBColorNum As Long, ByVal ForeElseBack As Boolean) As String
    Select Case RGBColorNum
        Case RGB(255, 255, 255) '- 00 - White.
            GetIRCColor = "00"
        Case RGB(0, 0, 0) '- 01 - Black.
            GetIRCColor = "01"
        Case RGB(0, 0, 127)  '- 02 - Blue.
            GetIRCColor = "02"
        Case RGB(0, 147, 0) '- 03 - Green.
            GetIRCColor = "03"
        Case RGB(255, 0, 0) '- 04 - Red.
            GetIRCColor = "04"
        Case RGB(127, 0, 0)  '- 05 - Brown.
            GetIRCColor = "05"
        Case RGB(156, 0, 156)  '- 06 - Magenta.
            GetIRCColor = "06"
        Case RGB(252, 127, 0) '- 07 - Orange.
            GetIRCColor = "07"
        Case RGB(255, 255, 0)  '- 08 - Yellow.
            GetIRCColor = "08"
        Case RGB(0, 252, 0) '- 09 - Light Green.
            GetIRCColor = "09"
        Case RGB(0, 147, 147) '- 10 - Cyan.
            GetIRCColor = "10"
        Case RGB(0, 255, 255) '- 11 - Light Cyan.
            GetIRCColor = "11"
        Case RGB(0, 0, 252) '- 12 - Light Blue.
            GetIRCColor = "12"
        Case RGB(255, 0, 255)  '- 13 - Pink.
            GetIRCColor = "13"
        Case RGB(127, 127, 127) '- 14 - Grey.
            GetIRCColor = "14"
        Case RGB(210, 210, 210) '- 15 - Light Grey.
            GetIRCColor = "15"
            
        Case RGB(Val("&H47"), Val("&H00"), Val("&H00"))
            GetIRCColor = "16"
        Case RGB(Val("&H47"), Val("&H21"), Val("&H00"))
            GetIRCColor = "17"
        Case RGB(Val("&H47"), Val("&H47"), Val("&H00"))
            GetIRCColor = "18"
        Case RGB(Val("&H32"), Val("&H47"), Val("&H00"))
            GetIRCColor = "19"
        Case RGB(Val("&H00"), Val("&H47"), Val("&H00"))
            GetIRCColor = "20"
        Case RGB(Val("&H00"), Val("&H47"), Val("&H2c"))
            GetIRCColor = "21"
        Case RGB(Val("&H00"), Val("&H47"), Val("&H47"))
            GetIRCColor = "22"
        Case RGB(Val("&H00"), Val("&H27"), Val("&H47"))
            GetIRCColor = "23"
        Case RGB(Val("&H00"), Val("&H00"), Val("&H47"))
            GetIRCColor = "24"
        Case RGB(Val("&H2e"), Val("&H00"), Val("&H47"))
            GetIRCColor = "25"
        Case RGB(Val("&H47"), Val("&H00"), Val("&H47"))
            GetIRCColor = "26"
        Case RGB(Val("&H47"), Val("&H00"), Val("&H2a"))
            GetIRCColor = "27"
        Case RGB(Val("&H74"), Val("&H00"), Val("&H00"))
            GetIRCColor = "28"
        Case RGB(Val("&H74"), Val("&H3a"), Val("&H00"))
            GetIRCColor = "29"
        Case RGB(Val("&H74"), Val("&H74"), Val("&H00"))
            GetIRCColor = "30"
        Case RGB(Val("&H51"), Val("&H74"), Val("&H00"))
            GetIRCColor = "31"
        Case RGB(Val("&H00"), Val("&H74"), Val("&H00"))
            GetIRCColor = "32"
        Case RGB(Val("&H00"), Val("&H74"), Val("&H49"))
            GetIRCColor = "33"
        Case RGB(Val("&H00"), Val("&H74"), Val("&H74"))
            GetIRCColor = "34"
        Case RGB(Val("&H00"), Val("&H40"), Val("&H74"))
            GetIRCColor = "35"
        Case RGB(Val("&H00"), Val("&H00"), Val("&H74"))
            GetIRCColor = "36"
        Case RGB(Val("&H40"), Val("&H00"), Val("&H74"))
            GetIRCColor = "37"
        Case RGB(Val("&H74"), Val("&H00"), Val("&H74"))
            GetIRCColor = "38"
        Case RGB(Val("&H74"), Val("&H00"), Val("&H45"))
            GetIRCColor = "39"
        Case RGB(Val("&Hb5"), Val("&H00"), Val("&H00"))
            GetIRCColor = "40"
        Case RGB(Val("&Hb5"), Val("&H63"), Val("&H00"))
            GetIRCColor = "41"
        Case RGB(Val("&Hb5"), Val("&Hb5"), Val("&H00"))
            GetIRCColor = "42"
        Case RGB(Val("&H7d"), Val("&Hb5"), Val("&H00"))
            GetIRCColor = "43"
        Case RGB(Val("&H00"), Val("&Hb5"), Val("&H00"))
            GetIRCColor = "44"
        Case RGB(Val("&H00"), Val("&Hb5"), Val("&H71"))
            GetIRCColor = "45"
        Case RGB(Val("&H00"), Val("&Hb5"), Val("&Hb5"))
            GetIRCColor = "46"
        Case RGB(Val("&H00"), Val("&H63"), Val("&Hb5"))
            GetIRCColor = "47"
        Case RGB(Val("&H00"), Val("&H00"), Val("&Hb5"))
            GetIRCColor = "48"
        Case RGB(Val("&H75"), Val("&H00"), Val("&Hb5"))
            GetIRCColor = "48"
        Case RGB(Val("&Hb5"), Val("&H00"), Val("&Hb5"))
            GetIRCColor = "50"
        Case RGB(Val("&Hb5"), Val("&H00"), Val("&H6b"))
            GetIRCColor = "51"
        Case RGB(Val("&Hff"), Val("&H00"), Val("&H00"))
            GetIRCColor = "52"
        Case RGB(Val("&Hff"), Val("&H8c"), Val("&H00"))
            GetIRCColor = "53"
        Case RGB(Val("&Hff"), Val("&Hff"), Val("&H00"))
            GetIRCColor = "54"
        Case RGB(Val("&Hb2"), Val("&Hff"), Val("&H00"))
            GetIRCColor = "55"
        Case RGB(Val("&H00"), Val("&Hff"), Val("&H00"))
            GetIRCColor = "56"
        Case RGB(Val("&H00"), Val("&Hff"), Val("&Ha0"))
            GetIRCColor = "57"
        Case RGB(Val("&H00"), Val("&Hff"), Val("&Hff"))
            GetIRCColor = "58"
        Case RGB(Val("&H00"), Val("&H8c"), Val("&Hff"))
            GetIRCColor = "59"
        Case RGB(Val("&H00"), Val("&H00"), Val("&Hff"))
            GetIRCColor = "60"
        Case RGB(Val("&Ha5"), Val("&H00"), Val("&Hff"))
            GetIRCColor = "61"
        Case RGB(Val("&Hff"), Val("&H00"), Val("&Hff"))
            GetIRCColor = "62"
        Case RGB(Val("&Hff"), Val("&H00"), Val("&H98"))
            GetIRCColor = "63"
        Case RGB(Val("&Hff"), Val("&H59"), Val("&H59"))
            GetIRCColor = "64"
        Case RGB(Val("&Hff"), Val("&Hb4"), Val("&H59"))
            GetIRCColor = "65"
        Case RGB(Val("&Hff"), Val("&Hff"), Val("&H71"))
            GetIRCColor = "66"
        Case RGB(Val("&Hcf"), Val("&Hff"), Val("&H60"))
            GetIRCColor = "67"
        Case RGB(Val("&H6f"), Val("&Hff"), Val("&H6f"))
            GetIRCColor = "68"
        Case RGB(Val("&H65"), Val("&Hff"), Val("&Hc9"))
            GetIRCColor = "69"
        Case RGB(Val("&H6d"), Val("&Hff"), Val("&Hff"))
            GetIRCColor = "70"
        Case RGB(Val("&H59"), Val("&Hb4"), Val("&Hff"))
            GetIRCColor = "71"
        Case RGB(Val("&H59"), Val("&H59"), Val("&Hff"))
            GetIRCColor = "72"
        Case RGB(Val("&Hc4"), Val("&H59"), Val("&Hff"))
            GetIRCColor = "73"
        Case RGB(Val("&Hff"), Val("&H66"), Val("&Hff"))
            GetIRCColor = "74"
        Case RGB(Val("&Hff"), Val("&H59"), Val("&Hbc"))
            GetIRCColor = "75"
        Case RGB(Val("&Hff"), Val("&H9c"), Val("&H9c"))
            GetIRCColor = "76"
        Case RGB(Val("&Hff"), Val("&Hd3"), Val("&H9c"))
            GetIRCColor = "77"
        Case RGB(Val("&Hff"), Val("&Hff"), Val("&H9c"))
            GetIRCColor = "78"
        Case RGB(Val("&He2"), Val("&Hff"), Val("&H9c"))
            GetIRCColor = "79"
        Case RGB(Val("&H9c"), Val("&Hff"), Val("&H9c"))
            GetIRCColor = "80"
        Case RGB(Val("&H9c"), Val("&Hff"), Val("&Hdb"))
            GetIRCColor = "81"
        Case RGB(Val("&H9c"), Val("&Hff"), Val("&Hff"))
            GetIRCColor = "82"
        Case RGB(Val("&H9c"), Val("&Hd3"), Val("&Hff"))
            GetIRCColor = "83"
        Case RGB(Val("&H9c"), Val("&H9c"), Val("&Hff"))
            GetIRCColor = "84"
        Case RGB(Val("&Hdc"), Val("&H9c"), Val("&Hff"))
            GetIRCColor = "85"
        Case RGB(Val("&Hff"), Val("&H9c"), Val("&Hff"))
            GetIRCColor = "86"
        Case RGB(Val("&Hff"), Val("&H94"), Val("&Hd3"))
            GetIRCColor = "87"
        Case RGB(Val("&H00"), Val("&H00"), Val("&H00"))
            GetIRCColor = "88"
        Case RGB(Val("&H13"), Val("&H13"), Val("&H13"))
            GetIRCColor = "89"
        Case RGB(Val("&H28"), Val("&H28"), Val("&H28"))
            GetIRCColor = "90"
        Case RGB(Val("&H36"), Val("&H36"), Val("&H36"))
            GetIRCColor = "91"
        Case RGB(Val("&H4d"), Val("&H4d"), Val("&H4d"))
            GetIRCColor = "92"
        Case RGB(Val("&H65"), Val("&H65"), Val("&H65"))
            GetIRCColor = "93"
        Case RGB(Val("&H81"), Val("&H81"), Val("&H81"))
            GetIRCColor = "94"
        Case RGB(Val("&H9f"), Val("&H9f"), Val("&H9f"))
            GetIRCColor = "95"
        Case RGB(Val("&Hbc"), Val("&Hbc"), Val("&Hbc"))
            GetIRCColor = "96"
        Case RGB(Val("&He2"), Val("&He2"), Val("&He2"))
            GetIRCColor = "97"
        Case RGB(Val("&Hff"), Val("&Hff"), Val("&Hff"))
            GetIRCColor = "98"
            
        Case IIf(ForeElseBack, pForecolor, pBackcolor) '- 99 - Default Foreground/Background - Not universally supported.
            GetIRCColor = "99"
                        
    End Select
End Function


Private Sub DepleetColorRecords(ByVal CursorLoc As Long, ByVal Width As Long)
    Dim Index As Long
    Dim cnt As Long
    Index = LocateColorRecord(CursorLoc)
    Do While Index <= UBound(pColorRecords)
        If CursorLoc <= pColorRecords(Index).StartLoc And CursorLoc + Width >= pColorRecords(Index).StartLoc + WidthOfColorRecord(Index) Then
            For cnt = Index To UBound(pColorRecords) - 1
                pColorRecords(cnt) = pColorRecords(cnt + 1)
            Next
            ReDim Preserve pColorRecords(0 To UBound(pColorRecords) - 1) As ColorRange
        Else
            If pColorRecords(Index).StartLoc > CursorLoc Then
                pColorRecords(Index).StartLoc = pColorRecords(Index).StartLoc - Width
            End If
            Index = Index + 1
        End If
    Loop
End Sub

Private Sub ExpandColorRecords(ByVal CursorLoc As Long, ByVal Width As Long)
    Dim Index As Long
    Index = LocateColorRecord(CursorLoc)
    Do While Index <= UBound(pColorRecords)
        If pColorRecords(Index).StartLoc > CursorLoc Then
            pColorRecords(Index).StartLoc = pColorRecords(Index).StartLoc + Width
        End If
        Index = Index + 1
    Loop
End Sub

Private Sub CleanColorRecords(ByVal CursorLoc As Long, ByVal Width As Long)
    'remove any color record whose width is 0 or less, with-in parameters ranges
    Dim Index As Long
    Index = LocateColorRecord(CursorLoc)
    Do While Index <= UBound(pColorRecords)
        'If pColorRecords(Index).StartLoc < CursorLoc + Width Then
        If pColorRecords(Index).StartLoc >= CursorLoc And pColorRecords(Index).StartLoc < CursorLoc + Width Then
            If WidthOfColorRecord(Index) = 0 Then
                DelColorRecord Index
            Else
                Index = Index + 1
            End If
        Else
            Exit Do
        End If
    Loop
End Sub

Private Function AddColorRecord(ByVal CursorLoc As Long, Optional ByVal Width As Long = 0) As Long
    Dim Index As Long
    Index = LocateColorRecord(CursorLoc)
    ReDim Preserve pColorRecords(0 To UBound(pColorRecords) + 1) As ColorRange
    If (UBound(pColorRecords) - 1) - Index > 0 Then
        Dim cnt As Long
        For cnt = UBound(pColorRecords) - 1 To Index Step -1
            pColorRecords(cnt).StartLoc = pColorRecords(cnt).StartLoc + Width
            pColorRecords(cnt + 1) = pColorRecords(cnt)
        Next
    Else
        Index = UBound(pColorRecords)
    End If
    pColorRecords(Index).StartLoc = CursorLoc
    AddColorRecord = Index
End Function

Private Sub DelColorRecord(ByVal Index As Long)
    Dim cnt As Long
    Dim Width As Long
    Width = WidthOfColorRecord(Index)
    For cnt = Index To UBound(pColorRecords) - 1
        pColorRecords(cnt) = pColorRecords(cnt + 1)
        pColorRecords(cnt).StartLoc = pColorRecords(cnt).StartLoc - Width
    Next
    ReDim Preserve pColorRecords(0 To UBound(pColorRecords) - 1) As ColorRange
End Sub

Private Function LocateColorRecord(ByVal CursorLoc As Long) As Long
    If UBound(pColorRecords) > 0 Then
        Dim cnt As Long
        For cnt = 0 To UBound(pColorRecords) - 1
            If CursorLoc >= pColorRecords(cnt).StartLoc And CursorLoc < pColorRecords(cnt + 1).StartLoc Then
                'seeking the inbetween marker, exit if found
                LocateColorRecord = cnt
                Exit Function
            ElseIf pColorRecords(cnt).StartLoc = pColorRecords(cnt + 1).StartLoc Then
                LocateColorRecord = cnt + 1
            End If
        Next
        
        'didn't find inbetween so check the last record
        If CursorLoc >= pColorRecords(UBound(pColorRecords)).StartLoc Then
            LocateColorRecord = UBound(pColorRecords)
            Exit Function
        End If
        
    End If
End Function

Private Function WidthOfColorRecord(ByVal Index As Long) As Long
    'colorrecord 0 is special, it contains the defaults and spans the entire contents,
    'when colorrecords are used, it is not used beyond the color record at index 0
    If Index >= 0 And Index < UBound(pColorRecords) Then
        WidthOfColorRecord = (pColorRecords(Index + 1).StartLoc - pColorRecords(Index).StartLoc)
    ElseIf Index = UBound(pColorRecords) Then
        WidthOfColorRecord = (pText.Length - pColorRecords(Index).StartLoc)
    ElseIf Index = 0 Then
        WidthOfColorRecord = pText.Length
    End If
End Function


Private Function CheckNumeric(ByRef StrText() As Byte, ByVal Loc As Long) As Boolean
    If Loc <= UBound(StrText) Then
        If IsNumeric(Chr(StrText(Loc))) Then
            CheckNumeric = True
        End If
    End If
End Function
Private Function CheckComma(ByRef StrText() As Byte, ByVal Loc As Long) As Boolean
    If Loc <= UBound(StrText) Then
        If Chr(StrText(Loc)) = "," Then
            CheckComma = True
        End If
    End If
End Function

Private Function SubClipPrintTextBlock(ByRef X1 As Single, ByRef Y1 As Single, ByRef StrText As String, Optional fColor As Variant, Optional bColor As Variant, Optional ByVal BoxFill As Boolean = False) As Boolean
    If StrText <> "" Then

        If Not IsMissing(fColor) Then
            X1 = X1 + ClipPrintText(X1, Y1, StrText, fColor, bColor, (bColor <> pBackcolor))
            
        Else
            X1 = X1 + ClipPrintText(X1, Y1, StrText, pBackBuffer.Forecolor, pBackBuffer.BackColor, (pBackBuffer.BackColor <> pBackcolor))
        End If

        StrText = ""
    End If
End Function

Private Function ClipPrintTextBlock(ByRef X1 As Single, ByRef Y1 As Single, ByRef StrText() As Byte, ByVal Offset As Long, Optional fColor As Variant, Optional bColor As Variant, Optional ByVal BoxFill As Boolean = False) As Boolean
    
    Dim cnt As Long
    Dim ClrRec As Long
    Dim line As Long
    Dim rmvCnt As Long
    Dim newClrRec As Boolean
    Dim ircText() As Byte
    ReDim ircText(LBound(StrText) To UBound(StrText)) As Byte
    Dim nextPrint As String
    
    cnt = LBound(StrText)
    
    pColorRecords(0).BackColor = pBackcolor
    pColorRecords(0).Forecolor = pForecolor
    
    ClrRec = LocateColorRecord(cnt + Offset)
    
    If Not IsMissing(fColor) Then
        pBackBuffer.Forecolor = fColor
    Else
        pBackBuffer.Forecolor = pColorRecords(ClrRec).Forecolor
    End If
    If Not IsMissing(bColor) Then
        pBackBuffer.BackColor = bColor
    Else
        pBackBuffer.BackColor = pColorRecords(ClrRec).BackColor
    End If
    
    line = LineFirstVisible
    
    Do While cnt <= UBound(StrText)

        Select Case StrText(cnt)
            Case 0, 13
                SubClipPrintTextBlock X1, Y1, nextPrint, fColor, bColor, BoxFill
                If Not newClrRec Then newClrRec = True
                
                rmvCnt = rmvCnt + 1
                
            Case 3 'convert IRC color codes to permanent color records in our system as they are seen
                    'we wont be handling them in the background for the overall controls hidden health
                    'a property maybe made that can return the text with IRC style color coding put back
                    
                SubClipPrintTextBlock X1, Y1, nextPrint, fColor, bColor, BoxFill
                If Not newClrRec Then newClrRec = True
                
                ClrRec = AddColorRecord((cnt - rmvCnt) + Offset)

                cnt = cnt + 1 'pass the chr(3)
                rmvCnt = rmvCnt + 1

                
                If CheckNumeric(StrText, cnt) Then
                    cnt = cnt + 1 'pass first digit
                    rmvCnt = rmvCnt + 1
                    'at least forecolor
                    If CheckNumeric(StrText, cnt) Then
                        'two digit forecolor
                        cnt = cnt + 1 'pass second digit
                        rmvCnt = rmvCnt + 1
                        pColorRecords(ClrRec).Forecolor = GetRGBColor(CLng(CStr(Chr(StrText(cnt - 2)) & Chr(StrText(cnt - 1)))), True)
                        
                    Else
                        'one digit forecolor
                        pColorRecords(ClrRec).Forecolor = GetRGBColor(CLng(CStr(Chr(StrText(cnt - 1)))), True)
                    End If
                
                    If CheckComma(StrText, cnt) Then
                        'may have back color
                        cnt = cnt + 1 'pass the comma
                        rmvCnt = rmvCnt + 1
                        
                        If CheckNumeric(StrText, cnt) Then
                            cnt = cnt + 1 'pass first digit
                            rmvCnt = rmvCnt + 1
                            'at least forecolor
                            If CheckNumeric(StrText, cnt) Then
                                'two digit forecolor
                                cnt = cnt + 1 'pass second digit
                                rmvCnt = rmvCnt + 1
                                pColorRecords(ClrRec).BackColor = GetRGBColor(CLng(CStr(Chr(StrText(cnt - 2)) & Chr(StrText(cnt - 1)))), False)
    
                            Else
                                'one digit forecolor
                                pColorRecords(ClrRec).BackColor = GetRGBColor(CLng(CStr(Chr(StrText(cnt - 1)))), False)
                            End If
                        End If

                    End If
                    
                End If
                
                cnt = cnt - 1 'skip a beat comming next in the loop
                
                pBackBuffer.Forecolor = pColorRecords(ClrRec).Forecolor
                pBackBuffer.BackColor = pColorRecords(ClrRec).BackColor
                
            Case 10
                ircText(cnt - rmvCnt) = StrText(cnt)
                
                SubClipPrintTextBlock X1, Y1, nextPrint, fColor, bColor, BoxFill

                ClrRec = LocateColorRecord((cnt - rmvCnt) + Offset + 1)
                If Not IsMissing(fColor) Then
                    pBackBuffer.Forecolor = fColor
               Else
                    pBackBuffer.Forecolor = pColorRecords(ClrRec).Forecolor
                End If
                If Not IsMissing(bColor) Then
                    pBackBuffer.BackColor = bColor
                Else
                    pBackBuffer.BackColor = pColorRecords(ClrRec).BackColor
                End If
    
                X1 = pOffsetX + LineColumnWidth
                Y1 = Y1 + Me.TextHeight()

                line = line + 1

            Case Else

                ircText(cnt - rmvCnt) = StrText(cnt)
                
                If ClrRec <= UBound(pColorRecords) Then
                    If (cnt - rmvCnt) + Offset >= pColorRecords(ClrRec).StartLoc Then
                        SubClipPrintTextBlock X1, Y1, nextPrint, fColor, bColor, BoxFill

                        pBackBuffer.BackColor = pColorRecords(ClrRec).BackColor
                        pBackBuffer.Forecolor = pColorRecords(ClrRec).Forecolor
                        
                        ClrRec = ClrRec + 1
                        
                    End If
                End If
                    
                nextPrint = nextPrint & IIf(StrText(cnt) = 9, TabSpace, Chr(StrText(cnt)))
                
        End Select
        
        cnt = cnt + 1
    Loop
    
    SubClipPrintTextBlock X1, Y1, nextPrint, fColor, bColor, BoxFill

    If newClrRec Then  ' we had IRC style coloring so add the final record, remove the control characters and clean the block
        
        ReDim Preserve ircText(LBound(ircText) To cnt - rmvCnt) As Byte

        pText.Pyramid Strands(ircText), Offset, cnt
        
        ClipPrintTextBlock = True

    End If

End Function

Private Function ClipLineDraw(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, Optional Color As Variant, Optional ByVal BoxFill As Boolean = False) As Boolean
    If ClippingWouldDraw(DrawableRect, RECT(X1, Y1, X2, Y2)) Then
        If BoxFill Then
            pBackBuffer.DrawLine X1 / Screen.TwipsPerPixelX, Y1 / Screen.TwipsPerPixelY, X2 / Screen.TwipsPerPixelX, Y2 / Screen.TwipsPerPixelY, Color, bf
        Else
            pBackBuffer.DrawLine X1 / Screen.TwipsPerPixelX, Y1 / Screen.TwipsPerPixelY, X2 / Screen.TwipsPerPixelX, Y2 / Screen.TwipsPerPixelY, Color
        End If
        ClipLineDraw = True
    End If
End Function

Private Sub UserControl_Click()
    dragStart = 0
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    Dim lpos As Long
    Dim lend As Long
    Dim ltmp1 As Long
    Dim ltmp2 As Long
    Dim ltmp3 As Long
    Dim usechar As Byte
    usechar = Asc(" ")
    
    ltmp1 = pText.Pass(Asc(" "), 0, SelStart)
    ltmp2 = pText.Pass(Asc(vbTab), 0, SelStart)
    ltmp3 = pText.Pass(Asc(vbLf), 0, SelStart)
    
    If ltmp1 > ltmp2 And ltmp1 > ltmp3 Then
        usechar = Asc(" ")
    ElseIf ltmp2 > ltmp1 And ltmp2 > ltmp3 Then
        usechar = Asc(vbTab)
    ElseIf ltmp3 > ltmp1 And ltmp3 > ltmp2 Then
        usechar = Asc(vbLf)
    End If
    
    lpos = pText.Pass(usechar, 0, SelStart)

    If lpos > 0 Then
        lpos = pText.poll(usechar, lpos, 0, SelStart) + 1

        ltmp1 = pText.poll(Asc(" "), 1, lpos + 1, pText.Length - (lpos + 1))
        ltmp2 = pText.poll(Asc(vbTab), 1, lpos + 1, pText.Length - (lpos + 1))
        ltmp3 = pText.poll(Asc(vbLf), 1, lpos + 1, pText.Length - (lpos + 1))
        
        If ltmp1 < ltmp2 And ltmp1 < ltmp3 And ltmp1 > 0 Then
            lend = ltmp1
        ElseIf ltmp2 < ltmp1 And ltmp2 < ltmp3 And ltmp2 > 0 Then
            lend = ltmp2
        ElseIf ltmp3 < ltmp1 And ltmp3 < ltmp2 And ltmp3 > 0 Then
            lend = ltmp3
        End If

        SelStart = lpos
        SelLength = ((lpos + 1) + lend) - lpos
    Else

        ltmp1 = pText.poll(Asc(" "), 1, 0, pText.Length)
        ltmp2 = pText.poll(Asc(vbTab), 1, 0, pText.Length)
        ltmp3 = pText.poll(Asc(vbLf), 1, 0, pText.Length)
        lend = pText.Length
        
        If ltmp1 < ltmp2 And ltmp1 < ltmp3 And ltmp1 > 0 Then
            lend = ltmp1
        ElseIf ltmp2 < ltmp1 And ltmp2 < ltmp3 And ltmp2 > 0 Then
            lend = ltmp2
        ElseIf ltmp3 < ltmp1 And ltmp3 < ltmp2 And ltmp3 > 0 Then
            lend = ltmp3
        End If
        
        SelStart = 0
        SelLength = lend
    End If

    RaiseEvent DblClick
End Sub

Private Sub UserControl_GotFocus()
    hasFocus = True
End Sub

Private Sub RefreshEditMenu()
    mnuUndo.Enabled = CanUndo
    mnuRedo.Enabled = CanRedo
    mnuCut.Enabled = Enabled And Not Locked And (SelLength > 0)
    mnuCopy.Enabled = (SelLength > 0)
    mnuPaste.Enabled = Enabled And Not Locked
    mnuDelete.Enabled = Enabled And Not Locked
    mnuSelectAll.Enabled = Enabled
    mnuClear.Enabled = Enabled And (Length > 0)
    mnuIndent.Enabled = (Enabled And Not Locked And (CountWord(SelText, vbLf) > 1))
    mnuUnindent.Enabled = (Enabled And Not Locked And (CountWord(SelText, vbLf) > 1))
End Sub

Private Sub UserControl_Initialize()
    Cancel = True
    
    ReDim pPageBreaks(0 To 0) As Long
    ReDim aText(0 To 0) As Strands
    Set aText(0) = New Strands

    Set tText = New Strands
    
    xUndoLimit = 150
    
    ResetUndoRedo
        
    SystemParametersInfo SPI_GETKEYBOARDSPEED, 0, keySpeed, 0
    Timer1.Interval = keySpeed * 10

    Set pBackBuffer = New Backbuffer
    pBackBuffer.hWnd = UserControl.hWnd
    pBackBuffer.Forecolor = ConvertColor(SystemColorConstants.vbWindowText)
    pForecolor = pBackBuffer.Forecolor
    pBackBuffer.BackColor = ConvertColor(SystemColorConstants.vbWindowBackground)
    pBackcolor = pBackBuffer.BackColor
    Set pBackBuffer.Font = UserControl.Font

    ScrollBar1.Backbuffer.hdc = pBackBuffer.hdc
    ScrollBar2.Backbuffer.hdc = pBackBuffer.hdc
    pTabSpace = "    "
    pLineNumbers = True

    ResetColors
    Cancel = False
    
    Hook Me
End Sub

Private Sub UserControl_InitProperties()
    pForecolor = GetSysColor(COLOR_WINDOWTEXT)
    pBackcolor = GetSysColor(COLOR_WINDOW)
    UserControl.BackColor = GetSysColor(COLOR_WINDOW)
    pScrollToCaret = True
    pHideSelection = True
    UserControl.Font.name = "Lucida Console"
    Set pBackBuffer.Font = UserControl.Font
    pMultiLine = True
    pScrollBars = vbScrollBars.Both
    pEnabled = True
    pTabSpace = "    "
    xUndoLimit = 150
    pLineNumbers = True
    ResetUndoRedo
    ResetColors
    
End Sub

Public Function LineOffset(ByVal LineIndex As Long) As Long ' _
Returns the offset amount of characters upto a line, specified by the zero based LineIndex. Example, LineOffset(0)=0.
Attribute LineOffset.VB_Description = "Returns the offset amount of characters upto a line, specified by the zero based LineIndex. Example, LineOffset(0)=0."
    If pText.Length > 0 Then
        LineOffset = pText.Offset(LineIndex + 1)
    End If
End Function

Public Function LineLength(ByVal LineIndex As Long) As Long ' _
Returns the length of characters with-in a line, specifiied by the zero based LineIndex.
Attribute LineLength.VB_Description = "Returns the length of characters with-in a line, specifiied by the zero based LineIndex."
    If pText.Length > 0 Then
        LineLength = pText.Offset(LineIndex + 2) - pText.Offset(LineIndex + 1)
    End If
End Function

Public Function LineText(ByVal LineIndex As Long) As String ' _
Returns the text with-in a line, specified by the zero based LineIndex.
Attribute LineText.VB_Description = "Returns the text with-in a line, specified by the zero based LineIndex."
    Dim lpos As Long
    lpos = LineLength(LineIndex)
    If lpos > 0 Then
        LineIndex = LineOffset(LineIndex)
        If LineIndex > 0 Then
            LineText = Convert(pText.Partial(LineIndex, lpos))
        Else
            LineText = Convert(pText.Partial(0, lpos))
        End If
    End If
End Function

Public Function LineIndex(Optional ByVal CharIndex As Long = -1) As Long ' _
Returns the zero based index line number of which the given zero based character index falls upon with-in Text.
Attribute LineIndex.VB_Description = "Returns the zero based index line number of which the given zero based character index falls upon with-in Text."
    If CharIndex = -1 Then CharIndex = pSel.StartPos
    If CharIndex > 0 Then
        LineIndex = pText.Pass(Asc(vbLf), 0, CharIndex)
    End If
End Function

Public Function LineCount() As Long ' _
Returns the numerical count of how many lines, delimited by line feeds, that exists with-in Text.
Attribute LineCount.VB_Description = "Returns the numerical count of how many lines, delimited by line feeds, that exists with-in Text."
    If pText.Length > 0 Then
        LineCount = pText.count
    End If
End Function

Private Function CaretLocation(Optional ByVal AtCharPos As Long = -1) As POINTAPI
    
    If pSel.StartPos <= 0 Then pSel.StartPos = 0
    If AtCharPos = -1 Then AtCharPos = pSel.StartPos
    If AtCharPos > 0 And AtCharPos <= pText.Length Then
        Dim cnt As Long
        cnt = pText.Pass(Asc(vbLf), 0, AtCharPos)
        If cnt >= 0 Then
            CaretLocation.Y = (TextHeight * cnt) + pOffsetY '+ (LineFirstVisible * TextHeight)
            'CaretLocation.X = Me.TextWidth * ((pText.Length - LineOffset(cnt)) - (pText.Length - AtCharPos))
            Dim part As String
            part = Left(LineText(cnt), ((pText.Length - LineOffset(cnt)) - (pText.Length - AtCharPos)))
            CaretLocation.X = Me.TextWidth(part)
        Else
            CaretLocation.Y = pOffsetY '- (LineFirstVisible * TextHeight)
        End If
    Else
        CaretLocation.Y = pOffsetY
    End If
    CaretLocation.X = CaretLocation.X + pOffsetX + LineColumnWidth

End Function

Public Function CaretFromPoint(ByVal X As Single, ByVal Y As Single) As Long ' _
Retruns the zero based character index position with-in Text based upon given pixel corrdinates, X and Y.
Attribute CaretFromPoint.VB_Description = "Retruns the zero based character index position with-in Text based upon given pixel corrdinates, X and Y."
    Dim lText As String
    X = (X - pOffsetX) - LineColumnWidth
    Y = ((Y - pOffsetY) \ TextHeight)
    If Y < LineCount() Then
        lText = Replace(LineText(Y), vbLf, "")
        If X >= TextWidth(lText) Then
            CaretFromPoint = LineOffset(Y) + Len(lText)
        Else
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
                
                xUndoActs(0).PriorTextData.Reset
                xUndoActs(0).AfterTextData.Reset

                If pSel.StartPos > pSel.StopPos Then
                    Swap pSel.StartPos, pSel.StopPos
                End If
                
                If pSel.StartPos <> pSel.StopPos Then
                    xUndoActs(0).PriorTextData.Concat Convert(Replace(SelText, vbCrLf, vbLf))
                    pText.Pinch pSel.StartPos, pSel.StopPos - pSel.StartPos
                    
                End If

                Set tText = New Strands
                tText.Concat Convert(vbLf)
                pText.Pyramid tText, pSel.StartPos, 0
                Set tText = Nothing
                
                xUndoActs(0).AfterTextData.Concat Convert(vbLf)
                    
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
                
                xUndoActs(0).PriorTextData.Reset
                xUndoActs(0).AfterTextData.Reset
                
                If pSel.StartPos = pSel.StopPos Then
                
                    If KeyCode = 8 And pSel.StartPos - 1 >= 0 And pSel.StartPos <= pText.Length Then  'backspace
                    
                        xUndoActs(0).PriorTextData.Concat Convert(pText.Partial(pSel.StartPos - 1, 1))
                        DepleetColorRecords pSel.StartPos - 1, 1
                        pText.Pinch pSel.StartPos - 1, 1
                        pSel.StartPos = pSel.StartPos - 1
                        
                    ElseIf KeyCode = 46 And pSel.StartPos >= 0 And pSel.StartPos < pText.Length Then 'delete

                        xUndoActs(0).PriorTextData.Concat Convert(pText.Partial(pSel.StartPos, 1))
                        DepleetColorRecords pSel.StartPos, 1
                        pText.Pinch pSel.StartPos, 1

                    End If
                Else 'delete or backspace
                    
                    xUndoActs(0).PriorTextData.Concat Convert(Replace(SelText, vbCrLf, vbLf))

                    If pSel.StartPos > pSel.StopPos Then
                        Swap pSel.StartPos, pSel.StopPos
                    End If
                    
                    DepleetColorRecords pSel.StartPos, pSel.StopPos - pSel.StartPos

                    If pSel.StartPos > 0 And pSel.StopPos < pText.Length Then
                        pText.Pinch pSel.StartPos, pSel.StopPos - pSel.StartPos
                    End If
                    

                End If
                pSel.StopPos = pSel.StartPos
                
                RaiseEventChange
                
            Case 9 'tab
                If pLocked Then Exit Sub
                
                If pSel.StartPos = pSel.StopPos Then
                    xUndoActs(0).PriorTextData.Reset
                    xUndoActs(0).AfterTextData.Reset
                    xUndoActs(0).AfterTextData.Concat Convert(Chr(KeyCode))
                    
                    ExpandColorRecords SelStart, 1
                    
                    RaiseEventChange
                Else
                
'                    lIndex = LineOffset(LineIndex(pSel.StartPos))
'                    temp = LineOffset(LineIndex(pSel.StopPos))
'                    If temp > lIndex Then

                        Dim tmpSel As Long
                        If Shift = 0 Then
                            
                            tmpSel = SelLength
                            Indenting SelStart, SelLength, Chr(9), True
                            ExpandColorRecords pSel.StartPos, pSel.StopPos - (tmpSel - SelLength)

                        ElseIf Shift = 1 Then
                            tmpSel = SelLength
                            Indenting SelStart, SelLength, Chr(9) & Chr(8), True
                            DepleetColorRecords pSel.StartPos, pSel.StopPos - (pSel.StartPos + tmpSel)
                        End If
                        
                        KeyCode = 0
  
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
                pSel.StartPos = LineOffset(pSel.StartPos) + LineLength(pSel.StartPos) - 1
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
    Dim hasLF As Boolean
    Dim tmpSel As RangeType
    Dim txtLines As String
    
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

    xUndoActs(0).PriorTextData.Reset
    xUndoActs(0).AfterTextData.Reset
                
    xUndoActs(0).PriorTextData.Concat Convert(Replace(SelText, vbCrLf, vbLf))
    txtLines = SelText
    
    For temp = LineIndex(tmpSel.StartPos) To LineIndex(tmpSel.StopPos)
    
        hasLF = InStr(txtLines, vbLf) > 0
        txt = RemoveNextArg(txtLines, vbLf)
        
        If hasLF Then tmpSel.StopPos = tmpSel.StopPos - 1
        
        txt = Replace(txt, vbLf, "")
        If Len(txt) > 0 Then
            If InStr(CharStr, Chr(8)) = 0 Then
                txt = CharStr & txt
                tmpSel.StopPos = tmpSel.StopPos + Len(CharStr)
            Else
                If Left(txt, Len(Replace(CharStr, Chr(8), ""))) = Replace(CharStr, Chr(8), "") Then
                    txt = Mid(txt, Len(Replace(CharStr, Chr(8), "")) + 1)
                    tmpSel.StopPos = tmpSel.StopPos - Len(CharStr)
                End If
            End If
        End If

        newtxt = newtxt & txt & IIf(hasLF, vbLf, "")
        
    Next
    SelText = newtxt
    
    pSel.StartPos = tmpSel.StartPos
    pSel.StopPos = tmpSel.StopPos
    
    xUndoActs(0).AfterTextData.Concat Convert(newtxt)

End Sub
Friend Sub KeyPageUp(ByRef Shift As Integer)
    pSel.StartPos = LineOffset(LineIndex - (UsercontrolHeight \ TextHeight))
    If Shift = 0 Then pSel.StopPos = pSel.StartPos
End Sub
Friend Sub KeyPageDown(ByRef Shift As Integer)
    Dim lIndex As Long
    lIndex = LineIndex + (UsercontrolHeight \ TextHeight)
    If lIndex > LineCount - 1 Then lIndex = LineCount - 1
    pSel.StartPos = LineOffset(lIndex)
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

            xUndoActs(0).PriorTextData.Reset
            xUndoActs(0).AfterTextData.Reset
            If pSel.StartPos > pSel.StopPos Then
                Swap pSel.StartPos, pSel.StopPos
            End If
                
            Dim tText As Strands
            If insertMode Then
                
                If pSel.StartPos + 1 < pText.Length Then
                    Set tText = New Strands
                    tText.Concat Convert(Chr(KeyAscii))
                    If pSel.StartPos = pSel.StopPos Then
                        xUndoActs(0).PriorTextData.Concat pText.Partial(pSel.StartPos, 1)
                        pText.Pyramid tText, pSel.StartPos, 1
                    Else
                        xUndoActs(0).PriorTextData.Concat pText.Partial(pSel.StartPos, pSel.StopPos - pSel.StartPos)
                        pText.Pyramid tText, pSel.StartPos, pSel.StopPos - pSel.StartPos
                    End If
                    Set tText = Nothing
                Else
                    pText.post CByte(KeyAscii)
                End If
            
            ElseIf pSel.StartPos < pText.Length Then

                Set tText = New Strands

                ExpandColorRecords pSel.StartPos, 1

                tText.Concat Convert(Chr(KeyAscii))
                If pSel.StopPos < pText.Length Then
                    If pSel.StopPos - pSel.StartPos > 0 Then
                        xUndoActs(0).PriorTextData.Concat pText.Partial(pSel.StartPos, pSel.StopPos - pSel.StartPos)
                    End If

                    pText.Pyramid tText, pSel.StartPos, pSel.StopPos - pSel.StartPos
                End If

                Set tText = Nothing
            Else
                ExpandColorRecords pText.Length, 1
                pText.post CByte(KeyAscii)
            End If
            
            xUndoActs(0).AfterTextData.post CByte(KeyAscii)
            
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

        Dim lpos As Long
        lpos = CaretFromPoint(X, Y)

        If (SelLength > 0 And lpos > SelStart And lpos < SelStart + SelLength) And dragStart = 0 Then
            If pSel.StartPos > pSel.StopPos Then
                dragStart = lpos - pSel.StopPos
            Else
                dragStart = lpos - pSel.StartPos
            End If
        Else
            pSel.StartPos = lpos
            If Shift = 0 Then pSel.StopPos = pSel.StartPos
        End If
    
        InvalidateCursor

    Else
        dragStart = 0
    End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)

    If Button = 1 And hasFocus Then

        Dim lpos As Long

        lpos = CaretFromPoint(X, Y)
    
        If lpos < pSel.StopPos And (dragStart = -1 Or dragStart = 0) Then
            pSel.StopPos = lpos
            dragStart = -1
        ElseIf (dragStart = -2 Or dragStart = 0) Then
            pSel.StartPos = lpos
            dragStart = -2
        ElseIf (dragStart > 0) Then
            If dText Is Nothing Then

                UserControl.OLEDrag
            Else
                pSel.StartPos = lpos
                pSel.StopPos = lpos
                UserControl_Paint
            End If
            
        End If

        Dim Loc As POINTAPI
        Loc = CaretLocation

        Dim newloc As POINTAPI
        newloc.X = Loc.X
        newloc.Y = Loc.Y

        If X < 0 Then
            If (-X < (UsercontrolWidth / 2)) Then 'slow
                newloc.X = newloc.X - TextWidth
            Else
                newloc.X = newloc.X - (TextWidth * 4)
            End If

        ElseIf X > UsercontrolWidth + LineColumnWidth Then
            If (X - (UsercontrolWidth + LineColumnWidth) < (UsercontrolWidth / 2)) Then 'slow
                newloc.X = newloc.X + TextWidth
            Else
                newloc.X = newloc.X + (TextWidth * 4)
            End If
        End If

        If Y < 0 Then
            If (-Y < (UsercontrolWidth / 2)) Then 'slow
                newloc.Y = newloc.Y - TextHeight
            Else
                newloc.Y = newloc.Y - (TextHeight * 4)
            End If
        ElseIf Y > UsercontrolHeight Then
            If (Y - UsercontrolHeight < (UsercontrolWidth / 2)) Then 'slow
                newloc.Y = newloc.Y + TextHeight
            Else
                newloc.Y = newloc.Y + (TextHeight * 4)
            End If
        End If

        If Loc.X <> newloc.X Or Loc.Y <> newloc.Y Then
            MakeCaretVisible newloc, False
        End If

    Else
        dragStart = 0

    End If
    
    If dragStart > 0 Then
        UserControl.MousePointer = 99
    Else
        If pLineNumbers And X < LineColumnWidth Then
            UserControl.MousePointer = 1
        Else
            UserControl.MousePointer = IIf(Enabled, 3, 1)
        End If
    End If
End Sub

Friend Sub InvalidateCursor()
    Timer1.Enabled = False
    If Not Cancel Then Timer1_Timer
End Sub

Friend Sub RaiseEventChange(Optional ByVal KeepUndo As Boolean = True)
  
    RaiseEventSelChange KeepUndo
    
    If Not Cancel Then
        Cancel = True
        
        If KeepUndo Then
            AddUndo
        End If
        
        RaiseEvent Change
        
        Cancel = False
        
        If Not KeepUndo Then
            InvalidateCursor
        Else
            UserControl_Paint
        End If
                
    End If
    
End Sub
Private Function RaiseEventSelChange(Optional ByVal KeepUndo As Boolean = False) As Boolean
    
    If KeepUndo Then
        xUndoActs(0).PriorSelRange.StartPos = xUndoActs(0).AfterSelRange.StartPos
        xUndoActs(0).PriorSelRange.StopPos = xUndoActs(0).AfterSelRange.StopPos
        xUndoActs(0).AfterSelRange.StartPos = pSel.StartPos
        xUndoActs(0).AfterSelRange.StopPos = pSel.StopPos
    End If
    
    If (pLastSel.StartPos <> pSel.StartPos Or pLastSel.StopPos <> pSel.StopPos) Then
        
        CanvasValidate
        SetScrollBars

        RaiseEvent SelChange
        
        RaiseEventSelChange = True
        pLastSel.StartPos = pSel.StartPos
        pLastSel.StopPos = pSel.StopPos
    
    End If
    
End Function
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If dragStart > 0 Then

        If Not dText Is Nothing Then
            Cancel = True
            
            xUndoActs(0).PriorTextData.Reset
            xUndoActs(0).AfterTextData.Reset
        
            If pSel.StartPos > pSel.StopPos Then
                ExpandColorRecords pSel.StopPos, dText.Length
                pText.Pyramid dText, pSel.StopPos, pSel.StartPos - pSel.StopPos
                pSel.StartPos = pSel.StopPos + dText.Length
            Else
                ExpandColorRecords pSel.StartPos, dText.Length
                pText.Pyramid dText, pSel.StartPos, pSel.StopPos - pSel.StartPos
                pSel.StopPos = pSel.StartPos + dText.Length
                
            End If

            xUndoActs(0).AfterTextData.Concat dText.Partial
            
            Set dText = Nothing
            Cancel = False
            
            RaiseEventChange
            
        Else
            Dim lpos As Long
            lpos = CaretFromPoint(X, Y)
            pSel.StartPos = lpos
            If Shift = 0 Then pSel.StopPos = pSel.StartPos
        End If

        dragStart = 0
        
    End If
            
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
    If Data.GetFormat(ClipBoardConstants.vbCFText) Then
        Unhook Me

        UserControl_MouseUp 1, 0, X, Y

        Hook Me
    End If
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Dim pt As POINTAPI
    
    Dim lpos As Long
    lpos = CaretFromPoint(X, Y)
    If dragStart < 1 Then
        If pSel.StartPos > pSel.StopPos Then
            dragStart = lpos - pSel.StopPos
        Else
            dragStart = lpos - pSel.StartPos
        End If
    End If
   
    pSel.StartPos = lpos
    pSel.StopPos = lpos

    If Not Data.GetFormat(ClipBoardConstants.vbCFText) Then Effect = vbDropEffectNone
    
    If Data.GetFormat(ClipBoardConstants.vbCFText) Then
        If dText Is Nothing Then
            Set dText = New NTNodes10.Strands
            dText.Concat Convert(Replace(Data.GetData(ClipBoardConstants.vbCFText), vbCrLf, vbLf))
        End If
        
    End If
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    Data.SetData SelText, ClipBoardConstants.vbCFText
    Set dText = New NTNodes10.Strands
    dText.Concat Convert(Replace(SelText, vbCrLf, vbLf))
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_Paint()
    If AutoRedraw Then
        Refresh
        PaintBuffer
    End If

    RaiseEvent Paint
End Sub
Private Function Strands(ByRef Text() As Byte) As Strands
    Set Strands = New Strands
    Strands.Concat Text
End Function
Public Static Sub Refresh()

    If (Not Cancel) Or colorOpen Then
        If Not colorOpen Then Cancel = True
        
        Set Font = UserControl.Font
        
        CanvasValidate False
    
        Dim curX As Single
        Dim curY As Single
        
        Dim cnt As Long
        Dim bpos As Long
        Dim epos As Long
        Dim reClean As Boolean
        
        Static lastOffsets As RangeType
        Static lastFirstLine As Long
        Static lastColumnWidth As Long
        Static lastPolls As RangeType
        
        Dim tmpSel As RangeType
        lastColumnWidth = LineColumnWidth
        
       ' If lastFirstLine <> LineFirstVisible Or lastOffsets.StartPos <> pOffsetX Or lastOffsets.StopPos <> pOffsetY Or (Not firstRun) Then
            
            lastFirstLine = LineFirstVisible
            
            lastOffsets.StartPos = pOffsetX
            lastOffsets.StopPos = pOffsetY
    
            bpos = pText.poll(Asc(vbLf), lastFirstLine)
            If bpos > 0 Then bpos = bpos + 1
            epos = pText.poll(Asc(vbLf), lastFirstLine + (UsercontrolHeight \ TextHeight) + 1)
            
            lastPolls.StartPos = bpos
            lastPolls.StopPos = epos
            
            CleanColorRecords bpos, epos
'        Else
'            bpos = lastPolls.StartPos
'            epos = lastPolls.StopPos
'        End If
            
        curX = pOffsetX + lastColumnWidth
        curY = pOffsetY + (lastFirstLine * TextHeight)
        
        If pSel.StartPos <> pSel.StopPos Then
            If pSel.StartPos > pSel.StopPos Then
                tmpSel.StartPos = pSel.StopPos
                tmpSel.StopPos = pSel.StartPos
            Else
                tmpSel.StartPos = pSel.StartPos
                tmpSel.StopPos = pSel.StopPos
            End If
        End If
        
        If ScrollBar1.Visible Then ScrollBar1.Refresh
        If ScrollBar2.Visible Then ScrollBar2.Refresh
        
        Dim tmpBack As Variant
        tmpBack = pBackBuffer.BackColor
        pBackBuffer.BackColor = pBackcolor
        pBackBuffer.DrawCls UsercontrolWidth + lastColumnWidth, UsercontrolHeight
        pBackBuffer.BackColor = tmpBack
        
        If pText.Length > 0 And epos - bpos > 0 Then
            
            If Enabled Then
                    
                If ((pSel.StartPos <> pSel.StopPos) And (hasFocus Xor ((Not hasFocus) And (Not pHideSelection)))) _
                    And (Not ((tmpSel.StopPos <= bpos) Or (tmpSel.StartPos >= epos))) Then
    
                    Static lastPos As RangeType
                    If lastPos.StartPos <> bpos Or lastPos.StopPos <> epos - bpos Then
                        tText.Reset
                        tText.Concat pText.Partial(bpos, epos - bpos)
                    End If
                
                    lastPos.StartPos = bpos
                    lastPos.StopPos = epos - bpos
    
                    If ((tmpSel.StartPos <= bpos) And (tmpSel.StopPos >= epos)) Then
                        If epos - bpos > 0 Then reClean = ClipPrintTextBlock(curX, curY, tText.Partial(0, epos - bpos), LineOffset(lastFirstLine), GetSysColor(COLOR_WINDOW), GetSysColor(COLOR_HIGHLIGHT), True)
                    ElseIf ((tmpSel.StartPos > bpos) And (tmpSel.StopPos < epos)) Then
                        If (tmpSel.StartPos - bpos) > 0 Then reClean = ClipPrintTextBlock(curX, curY, tText.Partial(0, (tmpSel.StartPos - bpos)), LineOffset(lastFirstLine), , , False)
                        If ((tmpSel.StopPos - bpos) - (tmpSel.StartPos - bpos)) > 0 Then reClean = reClean Or ClipPrintTextBlock(curX, curY, tText.Partial(tmpSel.StartPos - bpos, ((tmpSel.StopPos - bpos) - (tmpSel.StartPos - bpos))), LineOffset(lastFirstLine) + (tmpSel.StartPos - bpos), GetSysColor(COLOR_WINDOW), GetSysColor(COLOR_HIGHLIGHT), True)
                        If tText.Length - (tmpSel.StopPos - bpos) > 0 Then reClean = reClean Or ClipPrintTextBlock(curX, curY, tText.Partial(tmpSel.StopPos - bpos), LineOffset(lastFirstLine) + (tmpSel.StopPos - bpos), , , False)
                    ElseIf ((tmpSel.StartPos > bpos) And (tmpSel.StopPos >= epos)) Then
                        If (tmpSel.StartPos - bpos) > 0 Then reClean = ClipPrintTextBlock(curX, curY, tText.Partial(0, (tmpSel.StartPos - bpos)), LineOffset(lastFirstLine), , , False)
                        If tText.Length - (tmpSel.StartPos - bpos) > 0 Then reClean = reClean Or ClipPrintTextBlock(curX, curY, tText.Partial((tmpSel.StartPos - bpos)), LineOffset(lastFirstLine) + (tmpSel.StartPos - bpos), GetSysColor(COLOR_WINDOW), GetSysColor(COLOR_HIGHLIGHT), True)
                    ElseIf ((tmpSel.StartPos <= bpos) And (tmpSel.StopPos < epos)) Then
                        If ((epos - bpos) - (epos - tmpSel.StopPos)) > 0 Then reClean = ClipPrintTextBlock(curX, curY, tText.Partial(0, ((epos - bpos) - (epos - tmpSel.StopPos))), LineOffset(lastFirstLine), GetSysColor(COLOR_WINDOW), GetSysColor(COLOR_HIGHLIGHT), True)
                        If tText.Length - ((epos - bpos) - (epos - tmpSel.StopPos)) > 0 Then reClean = reClean Or ClipPrintTextBlock(curX, curY, tText.Partial(((epos - bpos) - (epos - tmpSel.StopPos))), LineOffset(lastFirstLine) + ((epos - bpos) - (epos - tmpSel.StopPos)), , , False)
                    End If
    
                Else
                    If epos - bpos > 0 Then reClean = ClipPrintTextBlock(curX, curY, pText.Partial(bpos, epos - bpos), bpos, , , False)
                End If
                    
            Else
                If epos - bpos > 0 Then reClean = ClipPrintTextBlock(curX, curY, pText.Partial(bpos, epos - bpos), bpos, GetSysColor(COLOR_GRAYTEXT), GetSysColor(COLOR_WINDOW), False)
            End If
         

            If Not colorOpen Then
                colorOpen = True
            Else
                colorOpen = False
            End If
            
            If colorOpen Then
                RaiseEvent ColorText(bpos, epos - bpos)
                colorOpen = False
            Else
                colorOpen = True
            End If
            
            If reClean And pText.Length > 0 Then 'irc codes happened
    
                CleanColorRecords bpos, epos
            End If
            
'        Else
'            firstRun = False
        End If
        
        If pLineNumbers Then
        
            pBackBuffer.DrawLine 0, 0, lastColumnWidth / Screen.TwipsPerPixelX, UsercontrolHeight, GetSysColor(COLOR_SCROLLBAR), bf
            For cnt = lastFirstLine To ((UsercontrolHeight \ TextHeight) + 1) + lastFirstLine
                pBackBuffer.DrawText (lastColumnWidth - (TextWidth((cnt + 1) & "."))) / Screen.TwipsPerPixelX, ((((cnt - lastFirstLine) * TextHeight) / Screen.TwipsPerPixelY) + 1), (cnt + 1) & ".", GetSysColor(COLOR_GRAYTEXT)
            Next
        End If
    
        If pText.Length > 0 Then
            If UBound(pPageBreaks) > 0 Or pPageBreaks(0) > 0 Then
                Dim lineMark As Long
                Dim tmpCnt As Long
                tmpCnt = LineCount
                cnt = 0
                lineMark = 0
                Do While cnt <= UBound(pPageBreaks) And lineMark <= (lastFirstLine + (UsercontrolHeight \ TextHeight) + 1) And lineMark < LineCount
                    lineMark = lineMark + pPageBreaks(cnt)
                    If lineMark > 0 Then
                        pBackBuffer.DrawLine lastColumnWidth / Screen.TwipsPerPixelX, ((lineMark - lastFirstLine) * TextHeight) / Screen.TwipsPerPixelY, UsercontrolWidth, (((lineMark - lastFirstLine) * TextHeight) / Screen.TwipsPerPixelY) + 1, GetSysColor(COLOR_SCROLLBAR), bf
                    End If
                    cnt = cnt + 1
                Loop
            End If
        End If
        
        If ScrollBar1.Visible And ScrollBar2.Visible Then
            DrawFrameControl pBackBuffer.hdc, RECT((UserControl.Width / Screen.TwipsPerPixelX) - (ScrollBar1.Width / Screen.TwipsPerPixelX), (UserControl.Height / Screen.TwipsPerPixelY) - (ScrollBar2.Height / Screen.TwipsPerPixelY), (UserControl.Width / Screen.TwipsPerPixelX), (UserControl.Height / Screen.TwipsPerPixelY)), DFC_SCROLL, DFCS_SCROLLSIZEGRIP
        End If
        
        If Not colorOpen Then Cancel = False
    End If
    
    RaiseEventSelChange True
    
End Sub

Friend Sub PaintBuffer()
    If Not Cancel Then
        ScrollBar1.Backbuffer.Paint (ScrollBar1.Left / Screen.TwipsPerPixelX), (ScrollBar1.Top / Screen.TwipsPerPixelY), ((ScrollBar1.Left + ScrollBar1.Width) / Screen.TwipsPerPixelX), ((ScrollBar1.Top + ScrollBar1.Height) / Screen.TwipsPerPixelY)
        ScrollBar2.Backbuffer.Paint (ScrollBar2.Left / Screen.TwipsPerPixelX), (ScrollBar2.Top / Screen.TwipsPerPixelY), ((ScrollBar2.Left + ScrollBar2.Width) / Screen.TwipsPerPixelX), ((ScrollBar2.Top + ScrollBar2.Height) / Screen.TwipsPerPixelY)

        pBackBuffer.Paint 0, 0, ((UsercontrolWidth + LineColumnWidth) / Screen.TwipsPerPixelX), (UsercontrolHeight / Screen.TwipsPerPixelY)
    End If

End Sub

Public Sub Paint()
    Refresh
    PaintBuffer
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Cancel = True
    
    BackColor = PropBag.ReadProperty("Backcolor", GetSysColor(COLOR_WINDOW))
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
    UndoLimit = PropBag.ReadProperty("UndoLimit", 150)
    LineNumbers = PropBag.ReadProperty("LineNumbers", True)
    CodePage = PropBag.ReadProperty("CodePage", 1)
    
    Cancel = False
End Sub

Private Sub UserControl_Resize()

    SetScrollBars
    CanvasValidate True
    UserControl_Paint
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

Friend Sub SetScrollBarsReverse()
    ScrollBar2.Value = -OffsetX
    ScrollBar1.Value = -OffsetY
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
        ScrollBar1.Max = (CanvasHeight - UsercontrolHeight)
        ScrollBar1.SmallChange = TextHeight
        ScrollBar1.LargeChange = ScrollBar1.SmallChange * 4
        If ScrollBar1.Visible Then

            If ScrollBar1.Top <> 0 Then ScrollBar1.Top = 0
            If ScrollBar1.Width <> Screen.TwipsPerPixelX ^ 2 Then ScrollBar1.Width = Screen.TwipsPerPixelX ^ 2
            If ScrollBar1.Left <> UsercontrolWidth + LineColumnWidth Then ScrollBar1.Left = UsercontrolWidth + LineColumnWidth
            If ScrollBar1.Height <> UsercontrolHeight Then ScrollBar1.Height = UsercontrolHeight
        End If
        ScrollBar2.Max = (CanvasWidth - UsercontrolWidth)
        ScrollBar2.SmallChange = TextWidth
        ScrollBar2.LargeChange = ScrollBar2.SmallChange * 4
        If ScrollBar2.Visible Then

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
    Reset
    Set xUndoActs(0).AfterTextData = Nothing
    Set xUndoActs(0).PriorTextData = Nothing
    Set aText(0) = Nothing
    Erase xUndoActs
    Erase aText
    Erase pPageBreaks
    Erase pColorRecords
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
    PropBag.WriteProperty "UndoLimit", UndoLimit, 150
    PropBag.WriteProperty "LineNumbers", LineNumbers, True
    PropBag.WriteProperty "CodePage", CodePage, 1
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



