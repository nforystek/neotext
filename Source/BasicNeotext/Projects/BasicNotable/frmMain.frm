VERSION 5.00
Object = "{C98B112F-745F-4542-B5B3-DDFADF1F6E2F}#1355.0#0"; "NTControls22.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Notable"
   ClientHeight    =   7230
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10035
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   10035
   StartUpPosition =   3  'Windows Default
   Begin NTControls22.CodeEdit txtMain 
      Height          =   2895
      Left            =   2820
      TabIndex        =   0
      Top             =   1455
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   5106
      FontSize        =   9
      ColorDream1     =   8388736
      ColorDream2     =   8388608
      ColorDream3     =   8421376
      ColorDream4     =   32768
      ColorDream5     =   32896
      ColorDream6     =   16512
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   600
   End
   Begin VB.Timer Stopper 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2070
      Top             =   330
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2925
      Top             =   195
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuDash7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuDash6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^{F4}
      End
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
      Begin VB.Menu mnuDash2 
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
      Begin VB.Menu mnuDash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuDash23872 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCOmment 
         Caption         =   "C&omment"
      End
      Begin VB.Menu mnuUncomment 
         Caption         =   "U&ncomment"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuFindReplace 
         Caption         =   "&Find/Replace"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuGoto 
         Caption         =   "&Goto Line..."
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuSpellCheck 
         Caption         =   "&Spell Check"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuTimeDate 
         Caption         =   "Time/&Date"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuRunBatch 
         Caption         =   "&Run Batch..."
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuAssemble 
         Caption         =   "&Assembly..."
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuWordWrap 
         Caption         =   "&Word Wrap"
      End
      Begin VB.Menu mnuSetFont 
         Caption         =   "Set &Font..."
      End
      Begin VB.Menu mnuSetColor 
         Caption         =   "Set &Colors"
         Begin VB.Menu mnuBackgroundColor 
            Caption         =   "&Background..."
         End
         Begin VB.Menu mnuForegroundColor 
            Caption         =   "&Foreground..."
         End
         Begin VB.Menu mnuDash3287 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBackAndInkColors 
            Caption         =   "&Batch Ink..."
         End
         Begin VB.Menu mnuVBScriptColors 
            Caption         =   "&VB Script..."
         End
         Begin VB.Menu mnuJScriptColors 
            Caption         =   "&JavaScript..."
         End
         Begin VB.Menu mnuNSISScript 
            Caption         =   "&NSIS Script..."
         End
         Begin VB.Menu mnuAssembly 
            Caption         =   "&Assembly..."
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private pTextChanged As Boolean
Private pTextFileName As String
Private OpenInitDir As String

Private CancelChange As Boolean
Private pBatchSignal As Integer

'**************************************************************
'API for StayOnTop and NotStayOnTop
'**************************************************************
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private elapsedTime As Single
Private diffLitedLine As Long
Private highLitedLine As Long
Private jumpClr(1 To 3) As OLE_COLOR

Private pColors(1 To 30) As OLE_COLOR

Public Property Get ColorProperty(ByVal Index As ColorProperties) As OLE_COLOR
    ColorProperty = pColors(Index)
End Property
Public Property Let ColorProperty(ByVal Index As ColorProperties, ByVal newVal As OLE_COLOR)
    pColors(Index) = newVal
    Select Case Index
        Case ColorProperties.ColorBackground
            txtMain.BackColor = newVal
        Case ColorProperties.ColorForeGround
            txtMain.ForeColor = newVal
            
        Case ColorProperties.BatchInkComment, ColorProperties.NSISScriptComment, ColorProperties.AssemblyComment
            txtMain.ColorDream1 = newVal
        Case ColorProperties.BatchInkCommands, ColorProperties.NSISScriptCommands, ColorProperties.AssemblyCommand
            txtMain.ColorDream2 = newVal
        Case ColorProperties.BatchInkFinished, ColorProperties.NSISScriptEqualJump, ColorProperties.AssemblyNotation
            txtMain.ColorDream3 = newVal
        Case ColorProperties.BatchInkCurrently, ColorProperties.NSISScriptElseJump, ColorProperties.AssemblyRegister
            txtMain.ColorDream4 = newVal
        Case ColorProperties.BatchInkIncomming, ColorProperties.NSISScriptAboveJump, ColorProperties.AssemblyParameter
            txtMain.ColorDream5 = newVal
        Case ColorProperties.AssemblyError
            txtMain.ColorDream6 = newVal
            
        Case ColorProperties.JScriptComment
            If GetFileExt(pTextFileName) = ".js" Then txtMain.ColorComment = newVal
        Case ColorProperties.JScriptStatements
            If GetFileExt(pTextFileName) = ".js" Then txtMain.ColorStatement = newVal
        Case ColorProperties.JScriptOperators
            If GetFileExt(pTextFileName) = ".js" Then txtMain.ColorOperator = newVal
        Case ColorProperties.JScriptVariables
            If GetFileExt(pTextFileName) = ".js" Then txtMain.ColorVariable = newVal
        Case ColorProperties.JScriptValues
            If GetFileExt(pTextFileName) = ".js" Then txtMain.ColorValue = newVal
        Case ColorProperties.JScriptError
            If GetFileExt(pTextFileName) = ".js" Then txtMain.ColorError = newVal

        Case ColorProperties.VBScriptComment
            If GetFileExt(pTextFileName) = ".vbs" Then txtMain.ColorComment = newVal
        Case ColorProperties.VBScriptStatements
            If GetFileExt(pTextFileName) = ".vbs" Then txtMain.ColorStatement = newVal
        Case ColorProperties.VBScriptOperators
            If GetFileExt(pTextFileName) = ".vbs" Then txtMain.ColorOperator = newVal
        Case ColorProperties.VBScriptVariables
            If GetFileExt(pTextFileName) = ".vbs" Then txtMain.ColorVariable = newVal
        Case ColorProperties.VBScriptValues
            If GetFileExt(pTextFileName) = ".vbs" Then txtMain.ColorValue = newVal
        Case ColorProperties.VBScriptError
            If GetFileExt(pTextFileName) = ".vbs" Then txtMain.ColorError = newVal

    End Select
   ' txtMain.Redraw
End Property

Public Property Get BatchSignal() As Integer
    BatchSignal = pBatchSignal
End Property
Public Property Let BatchSignal(ByVal newVal As Integer)
    pBatchSignal = newVal
End Property

Public Property Let Running(ByVal newVal As Boolean)
    If Not (txtMain.Locked = newVal) Then txtMain.Locked = newVal
    If newVal Then
        txtMain.BackColor = fMain.BackColor
    Else
        txtMain.BackColor = ColorProperty(ColorBackground)
    End If
    
    SetStyleing
    mnuTools_Click
    mnuOptions_Click
    mnuFile_Click
    mnuEdit_Click

End Property

Public Static Sub StartRun()
    If BatchSignal = 0 Or BatchSignal = 2 Then
        BatchSignal = 0
        Stopper.Tag = "False"
        Running = True
        On Error Resume Next
        Load frmBatch
        If Err Then Resume
        On Error GoTo 0
        If Stopper.Tag = "False" Then
            Stopper.Enabled = True
        End If
    End If
End Sub

Public Static Sub StopRun()
    If BatchSignal = 0 Then
        mnuRunBatch.Enabled = False
        BatchSignal = 1
        Stopper.Tag = "True"
        Stopper.Enabled = True
    End If
    mnuRunBatch.Caption = "&Run Batch..."
    mnuRunBatch.Enabled = True
End Sub

Private Sub mnuAssemble_Click()
MsgBox "Shorthand COM assembly" & vbCrLf & _
"(save as .asm, F5 to compile)" & vbCrLf & vbCrLf & _
"ADD <reg>, <val>" & vbCrLf & _
"CMP <reg>, <val>" & vbCrLf & _
"MOV <reg>, <val(s)>" & vbCrLf & _
"INC <reg>" & vbCrLf & vbCrLf & _
"POP <peg>" & vbCrLf & _
"PSH <peg>" & vbCrLf & vbCrLf & _
"JME <loc>" & vbCrLf & _
"JMP <loc>" & vbCrLf & _
"INT <int>" & vbCrLf & _
"RET" & vbCrLf & vbCrLf & _
"DAT <txt>, <eof>" & vbCrLf & vbCrLf & _
"<reg>" & vbCrLf & _
"    AX, AL, AH" & vbCrLf & _
"    BX , BL, BH" & vbCrLf & _
"    CX , CL, CH" & vbCrLf & _
"    DX , DL, DH" & vbCrLf & vbCrLf & _
"<peg>" & vbCrLf & _
"    AX , BX, CX, DX" & vbCrLf

End Sub

Private Sub mnuAssembly_Click()
    Load frmScriptColors
    frmScriptColors.InitializeToScript "Assembly"
    frmScriptColors.Show 1, Me
End Sub

Private Sub mnuComment_Click()
    txtMain.Comment True
End Sub

Private Sub mnuUncomment_Click()
    txtMain.Comment False
End Sub

Private Sub Stopper_Timer()
    If Not BatchSignal = 1 Then
        Running = False
        BatchSignal = 2
        Stopper.Enabled = False
    End If
End Sub

Public Property Get Running() As Boolean
    Running = (txtMain.Locked)
End Property

Private Sub mnuBackgroundColor_Click()
    On Error Resume Next
    
    CommonDialog1.Color = ColorProperty(ColorProperties.ColorBackground)
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlCCFullOpen Or cdlCCRGBInit
    CommonDialog1.ShowColor
    
    If Err = 0 Then
        ColorProperty(ColorProperties.ColorBackground) = CommonDialog1.Color
    Else
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Sub mnuForegroundColor_Click()
    On Error Resume Next
    
    CommonDialog1.Color = ColorProperty(ColorProperties.ColorForeGround)
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlCCFullOpen Or cdlCCRGBInit
    CommonDialog1.ShowColor
    
    If Err = 0 Then
        ColorProperty(ColorProperties.ColorForeGround) = CommonDialog1.Color
    Else
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Sub mnuBackAndInkColors_Click()
    Load frmScriptColors
    frmScriptColors.InitializeToScript "BatchInk"
    frmScriptColors.Show 1, Me
End Sub

Private Sub mnuJScriptColors_Click()
    Load frmScriptColors
    frmScriptColors.InitializeToScript "JScript"
    frmScriptColors.Show 1, Me
End Sub

Private Sub mnuVBScriptColors_Click()
    Load frmScriptColors
    frmScriptColors.InitializeToScript "VBScript"
    frmScriptColors.Show 1, Me
End Sub


Private Sub mnuNSISScript_Click()
    Load frmScriptColors
    frmScriptColors.InitializeToScript "NSISScript"
    frmScriptColors.Show 1, Me
End Sub

    
Public Sub StayOnTop(ByRef myForm, Optional ByVal HoldPos As Boolean = False)
    If HoldPos Then
        SetWindowPos myForm.hwnd, HWND_TOPMOST, myForm.Left / Screen.TwipsPerPixelX, _
        myForm.Top / Screen.TwipsPerPixelY, myForm.Width / Screen.TwipsPerPixelX, _
        myForm.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW
    Else
        SetWindowPos myForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW
    End If
End Sub

Public Sub NotStayOnTop(ByRef myForm, Optional ByVal HoldPos As Boolean = False)
    If HoldPos Then
        SetWindowPos myForm.hwnd, HWND_NOTOPMOST, myForm.Left / Screen.TwipsPerPixelX, _
        myForm.Top / Screen.TwipsPerPixelY, myForm.Width / Screen.TwipsPerPixelX, _
        myForm.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW
    Else
        SetWindowPos myForm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW
    End If
End Sub

Public Property Get TextFileName() As String
    TextFileName = pTextFileName
End Property
Private Property Let TextFileName(ByVal newVal As String)
    pTextFileName = newVal
    SetCaption
End Property
Private Sub SetCaption()
    Dim star As String
    If TextChanged Then
        star = "*"
    Else
        star = ""
    End If
    Me.Caption = "[" & GetFileName(pTextFileName) & star & "] - Notable Ink"
End Sub
Private Property Get TextChanged() As Boolean
    TextChanged = pTextChanged
End Property
Private Property Let TextChanged(ByVal newVal As Boolean)
    pTextChanged = newVal
    SetCaption
End Property

Public Function GetFileExt(ByVal URL As String, Optional ByVal LowerCase As Boolean = True, Optional ByVal RemoveDot As Boolean = False) As String
    If InStrRev(URL, ".") > 0 Then
        If LowerCase Then
            GetFileExt = Trim(LCase(Mid(URL, (InStrRev(URL, ".") + -CInt(RemoveDot)))))
        Else
            GetFileExt = Mid(URL, (InStrRev(URL, ".") + -CInt(RemoveDot)))
        End If
    Else
        GetFileExt = vbNullString
    End If
End Function
Friend Property Get IsScript() As Boolean
    IsScript = ((GetFileExt(TextFileName, True, True) = "ink") Or (GetFileExt(TextFileName, True, True) = "bat"))
End Property

Public Sub SetStyleing()
    CancelChange = True
    txtMain.BackColor = ColorProperty(ColorBackground)
    txtMain.ForeColor = ColorProperty(ColorForeGround)
    Select Case GetLanguage
        Case "VBScript"
            txtMain.Language = "VBScript"
            txtMain.ColorComment = ColorProperty(VBScriptComment)
            txtMain.ColorStatement = ColorProperty(VBScriptStatements)
            txtMain.ColorOperator = ColorProperty(VBScriptOperators)
            txtMain.ColorVariable = ColorProperty(VBScriptVariables)
            txtMain.ColorValue = ColorProperty(VBScriptValues)
            txtMain.ColorError = ColorProperty(VBScriptError)
        Case "JScript"
            txtMain.Language = "JScript"
            txtMain.ColorComment = ColorProperty(JScriptComment)
            txtMain.ColorStatement = ColorProperty(JScriptStatements)
            txtMain.ColorOperator = ColorProperty(JScriptOperators)
            txtMain.ColorVariable = ColorProperty(JScriptVariables)
            txtMain.ColorValue = ColorProperty(JScriptValues)
            txtMain.ColorError = ColorProperty(JScriptError)

            
        Case "Defined", "BatchInk", "NSISScript", "Assembly"
            txtMain.Language = "Defined"
            
            If IsScript Then
                txtMain.ColorDream1 = ColorProperty(BatchInkComment)
                txtMain.ColorDream2 = ColorProperty(BatchInkCommands)
                txtMain.ColorDream3 = ColorProperty(BatchInkFinished)
                txtMain.ColorDream4 = ColorProperty(BatchInkCurrently)
                txtMain.ColorDream5 = ColorProperty(BatchInkIncomming)
            ElseIf GetFileExt(TextFileName, True, True) = "asm" Then
                txtMain.ColorDream1 = ColorProperty(AssemblyComment)
                txtMain.ColorDream2 = ColorProperty(AssemblyCommand)
                txtMain.ColorDream3 = ColorProperty(AssemblyNotation)
                txtMain.ColorDream4 = ColorProperty(AssemblyRegister)
                txtMain.ColorDream5 = ColorProperty(AssemblyParameter)
                txtMain.ColorDream6 = ColorProperty(AssemblyError)
            Else
                txtMain.ColorDream1 = ColorProperty(NSISScriptComment)
                txtMain.ColorDream2 = ColorProperty(NSISScriptCommands)
                txtMain.ColorDream3 = ColorProperty(NSISScriptEqualJump)
                txtMain.ColorDream4 = ColorProperty(NSISScriptElseJump)
                txtMain.ColorDream5 = ColorProperty(NSISScriptAboveJump)
            End If

            ColorBatchFile
            
        Case Else
            txtMain.Language = "PlainText"
    End Select
    CancelChange = False
End Sub

Private Function GetLanguage() As String
    Select Case GetFileExt(pTextFileName, True, True)
        Case "vbs", "vbscript", "vb", "bas", "cls", "frm", "ctl"
            GetLanguage = "VBScript"
        Case "js", "jscript", "java"
            GetLanguage = "JScript"
        Case "ink", "bat", "nsh", "nsi", "asm"
            GetLanguage = "Defined"
        Case Else
            GetLanguage = "Unknown"
    End Select
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

Private Sub Form_Load()

    Stopper.Tag = "False"
    CancelChange = True

    TextFileName = ""
    
    LoadINI
    If Me.Top < 0 Then Me.Top = 0
    If Me.Left < 0 Then Me.Left = 0
        
    If Command <> "" Then
        If LCase(Left(Trim(LCase(Command)), 5)) = "/run " Or _
            LCase(Left(Trim(LCase(Command)), 6)) = "/exec " Then
            If FileExists(Mid(Replace(Command, """", ""), 6)) Then
                
                If LCase(Left(Trim(LCase(Command)), 5)) = "/run " Then
                    Me.Show
                    DoEvents
                    Timer1.Enabled = True
                End If
                OpenDocument Mid(Replace(Command, """", ""), 6)
                mnuRunBatch_Click

            Else
                If MsgBox("File not found - [" & Mid(Replace(Command, """", ""), 6) & "]", vbInformation + vbOKCancel, App.Title) = vbCancel Then Unload Me
            End If
        ElseIf LCase(Left(Trim(LCase(Command)), 6)) = "/open " Then
            If FileExists(Mid(Replace(Command, """", ""), 7)) Then
                OpenDocument Mid(Replace(Command, """", ""), 7)
            Else
                If MsgBox("File not found - [" & Mid(Replace(Command, """", ""), 7) & "]", vbInformation + vbOKCancel, App.Title) = vbCancel Then Unload Me
            End If
        ElseIf LCase(Left(Trim(LCase(Command)), 9)) = "/runexit " Or _
             LCase(Left(Trim(LCase(Command)), 9)) = "/runhide " Then
            If FileExists(Mid(Replace(Command, """", ""), 10)) Then
                If LCase(Left(Trim(LCase(Command)), 9)) = "/runexit " Then
                    Me.Show
                    DoEvents
                    Timer1.Enabled = True
                End If
                
                OpenDocument Mid(Replace(Command, """", ""), 10)
                mnuRunBatch_Click
                If LCase(Left(Trim(LCase(Command)), 9)) = "/runhide " Or (Stopper.Tag = "False") Then
                     Unload Me
                End If
                
               
            Else
                MsgBox "File not found - [" & Mid(Replace(Command, """", ""), 10) & "]", vbInformation + vbOKOnly, App.Title
                Unload Me
            End If
        ElseIf FileExists(Replace(Command, """", "")) Then
            OpenDocument Replace(Command, """", "")
        Else
            MsgBox "File not found - [" & Replace(Command, """", "") & "]", vbInformation + vbOKOnly, App.Title
        End If
    End If

    pTextChanged = False
    CancelChange = False
    SetCaption
    
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    CancelChange = True
    Timer1.Enabled = False
    If UnloadMode = 0 Then
        Cancel = Not CloseDocument
        If Not Cancel And Running Then
            StopRun
        Else
            Timer1.Enabled = True
        End If
    End If
    CancelChange = False
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    txtMain.Top = 0
    txtMain.Left = 0
    txtMain.Width = Me.ScaleWidth
    txtMain.Height = Me.ScaleHeight

    If Err Then Err.Clear
    On Error GoTo 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
    CancelChange = True
    Timer1.Enabled = False
    
    SaveINI
    If PathExists(App.path & "\" & App.EXEName & ".bat", True) Then
        On Error Resume Next
        Kill App.path & "\" & App.EXEName & ".bat"
        If Err Then Err.Clear
        On Error GoTo 0
    End If
    
    Set fMain = Nothing
    Unload Me
    If Not IsDebugger Then
        TerminateProcess GetCurrentProcess, 0
    End If
    
End Sub

Private Sub OpenDocument(ByVal FileName As String)
    TextFileName = FileName
    txtMain.FileName = FileName
    OpenInitDir = GetFilePath(FileName)
    
    SetStyleing
    TextChanged = False
End Sub
Private Function CloseDocument() As Boolean
    If TextChanged Then
        Select Case MsgBox("Do you want to save changes to - [" & GetFileName(TextFileName) & "]?", vbQuestion + vbYesNoCancel, "Notable Ink")
            Case vbYes
                mnuSave_Click
                CloseDocument = True
            Case vbNo
                CloseDocument = True
            Case vbCancel
                CloseDocument = False
        End Select
    Else
        CloseDocument = True
    End If
End Function

Private Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuFile_Click()
    mnuSave.Enabled = (Not txtMain.Locked) And (Not Running)
    mnuNew.Enabled = Not Running
    mnuOpen.Enabled = Not Running
    mnuSaveAs.Enabled = Not Running
    mnuPrint.Enabled = Not Running
    mnuAbout.Enabled = Not Running
End Sub

Private Sub mnuFindReplace_Click()
    frmFind.Show
End Sub

Private Sub mnuGoto_Click()
    If Not mnuWordWrap.Checked Then
        frmGoto.Show
    End If
End Sub

Private Sub mnuNew_Click()
    If CloseDocument Then
        txtMain.Text = ""
        TextFileName = ""
        txtMain.Locked = False
        TextChanged = False
        SetStyleing
    End If
End Sub

Private Sub mnuOpen_Click()
    On Error Resume Next
    
    CommonDialog1.InitDir = OpenInitDir
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNLongNames Or cdlOFNPathMustExist Or cdlOFNFileMustExist
    CommonDialog1.Filter = "Text and Script Files|*.txt;*.bat;*.nsi;*.nsh;*.asm;*.ink;*.js;*.jscript;*.java;*.vbs;*.vbscript;*.vb;*.bas;*.cls;*.frm;*.ctl|All Files (*.*)|*.*"
    CommonDialog1.FilterIndex = 1
    
    CommonDialog1.ShowOpen
    
    If Err = 0 Then
        If CloseDocument Then
            txtMain.Locked = ((CommonDialog1.Flags And cdlOFNReadOnly) = cdlOFNReadOnly)
            OpenDocument CommonDialog1.FileName
            Form_Resize
        End If
    Else
        Err.Clear
    End If
    On Error GoTo 0
End Sub


Private Sub mnuOptions_Click()
    mnuWordWrap.Enabled = (Not IsScript)
    mnuSetFont.Enabled = (Not Running)
    mnuSetColor.Enabled = (Not Running)
End Sub

Private Sub mnuPrint_Click()
On Error GoTo ErrHandler

    Dim cnt As Integer
    Dim AllPages As Boolean
    Dim Selection As Boolean
    
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlPDHidePrintToFile Or cdlPDUseDevModeCopies Or cdlPDNoPageNums
    
    CommonDialog1.ShowPrinter
    
    AllPages = (CommonDialog1.Flags = cdlPDAllPages)
    
    Printer.Orientation = CommonDialog1.Orientation
    Printer.Copies = CommonDialog1.Copies
    Printer.FontBold = txtMain.Font.Bold
    Printer.FontItalic = txtMain.Font.Italic
    Printer.FontUnderline = txtMain.Font.Underline
    Printer.FontStrikethru = txtMain.Font.Strikethru
    Printer.FontSize = txtMain.Font.size
    Printer.FontName = txtMain.Font.Name
    
    Dim PrintText As String
    If AllPages Then
        PrintText = txtMain.Text
    Else
        PrintText = txtMain.SelText
    End If
        
    Printer.NewPage
    Printer.Print PrintText
    Printer.EndDoc
        
ErrHandler:
    If Err Then Err.Clear
    On Error GoTo 0
End Sub

Private Sub mnuRunBatch_Click()
    If GetFileExt(TextFileName, True, True) = "asm" Then
        ComAsmRun TextFileName, txtMain.Text
    ElseIf GetFileExt(TextFileName, True, True) = "nsi" Then
        RunProcess "C:\Program Files\NSIS\makensisw.exe", """" & TextFileName & """", vbNormal, False
        
    ElseIf Not IsScript Then
        MsgBox "Only batch, ink, and asm files may be executed and" & vbCrLf & _
            "must have an extention of .BAT, .INK or .ASM" & vbCrLf & _
            "Press F1 for more information on Assembly.", vbInformation, "Notable Ink"
    Else
        If Running Then
            StopRun
        Else
            Running = True
            StartRun
        End If
    End If
End Sub

Private Sub mnuSpellCheck_Click()
    If Not IsScript Then
        txtMain.Text = SpellCheckText(txtMain.Text)
    End If
End Sub

Private Function SpellCheckText(ByVal Text As String) As String
        
    Load frmSpellCheck
    frmSpellCheck.Start Text
    Do Until frmSpellCheck.Finished
        DoEvents
    Loop
    SpellCheckText = frmSpellCheck.CheckText.Text
    Unload frmSpellCheck

End Function

Private Sub mnuTools_Click()
    mnuGoto.Enabled = (Not mnuWordWrap.Checked) And (Not Running)
    mnuSpellCheck.Enabled = (Not IsScript) And (Not Running)
    mnuWordWrap.Checked = txtMain.WordWrap And (Not Running)
    mnuRunBatch.Caption = IIf(Running, "&Stop Batch...", "&Run Batch...")
    mnuFindReplace.Enabled = (Not Running)
    mnuGoto.Enabled = (Not txtMain.WordWrap) And (Not Running)
    mnuTimeDate.Enabled = (Not Running)
End Sub

Private Sub mnuUndo_Click()
    If txtMain.CanUndo Then
        CancelChange = True
        txtMain.Undo
        CancelChange = False
    End If
End Sub
Private Sub mnuRedo_Click()
    If txtMain.CanRedo Then
        CancelChange = True
        txtMain.Redo
        CancelChange = False
    End If
End Sub

Private Sub mnuCopy_Click()
    If EnableCopy Then
        txtMain.Copy
    End If
End Sub
Private Sub mnuCut_Click()
    If EnableCut Then
        txtMain.Cut
    End If
End Sub
Private Sub mnuPaste_Click()
    If EnablePaste Then
        txtMain.Paste
    End If
End Sub
Private Sub mnuDelete_Click()
    If EnableDelete Then
        If txtMain.SelLength = 0 Then
            txtMain.SelLength = 1
        End If
        txtMain.SelText = ""
    End If
End Sub
Private Sub mnuSelectAll_Click()
    If EnableSelectAll Then
        txtMain.SelStart = 0
        txtMain.SelLength = Len(txtMain.Text)
    End If
End Sub

Private Function EnableCopy() As Boolean
    EnableCopy = (txtMain.SelLength > 0)
End Function
Private Function EnableCut() As Boolean
    EnableCut = (txtMain.SelLength > 0)
End Function
Private Function EnablePaste() As Boolean
    EnablePaste = (Len(Clipboard.GetText(vbCFText)) > 0)
End Function
Private Function EnableDelete() As Boolean
    EnableDelete = (txtMain.SelLength > 0)
End Function
Private Function EnableSelectAll() As Boolean
    EnableSelectAll = (Len(txtMain.Text) > 0)
End Function
Private Sub mnuEdit_Click()
    On Error Resume Next
    
    mnuCopy.Enabled = EnableCopy And (Not Running)
    mnuCut.Enabled = EnableCut And (Not txtMain.Locked) And (Not Running)
    mnuPaste.Enabled = EnablePaste And (Not txtMain.Locked) And (Not Running)
    mnuDelete.Enabled = EnableDelete And (Not txtMain.Locked) And (Not Running)
    mnuSelectAll.Enabled = EnableSelectAll And (Not Running)
    
    mnuUndo.Enabled = txtMain.CanUndo And (Not txtMain.Locked) And (Not Running)
    mnuRedo.Enabled = txtMain.CanRedo And (Not txtMain.Locked) And (Not Running)

    mnuDash23872.Visible = IsScript
    mnuCOmment.Visible = IsScript
    mnuUncomment.Visible = IsScript
    mnuCOmment.Enabled = Not Running
    mnuUncomment.Enabled = Not Running
    
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuSave_Click()
    If Not txtMain.Locked Then
        If Not FileExists(TextFileName) Then
            mnuSaveAs_Click
        Else
            WriteTextFile TextFileName, txtMain.Text
            TextChanged = False
        End If
    End If
End Sub
Private Sub mnuSaveAs_Click()
    On Error Resume Next
    
    If FileExists(TextFileName) Then CommonDialog1.InitDir = GetFilePath(TextFileName)
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
    CommonDialog1.Filter = "Text and Script Files|*.txt;*.bat;*.asm;*.ink;*.nsi;*.nsh;*.js;*.jscript;*.java;*.vbs;*.vbscript;*.vb;*.bas;*.cls;*.frm;*.ctl|All Files (*.*)|*.*"
    CommonDialog1.FilterIndex = 1
    
    CommonDialog1.ShowSave
    
    If Err = 0 Then
        TextFileName = CommonDialog1.FileName
        WriteTextFile TextFileName, txtMain.Text
        TextChanged = False
        Form_Resize
    Else
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Sub mnuSetFont_Click()
    On Error GoTo exitfontset
    
    CommonDialog1.FontName = txtMain.Font.Name
    CommonDialog1.FontItalic = txtMain.Font.Italic
    CommonDialog1.FontBold = txtMain.Font.Bold
    CommonDialog1.FontStrikethru = txtMain.Font.StrikeThrough
    CommonDialog1.FontUnderline = txtMain.Font.Underline
    CommonDialog1.FontSize = txtMain.Font.size
    CommonDialog1.Color = txtMain.ForeColor
    
    
    CommonDialog1.CancelError = True
    
    CommonDialog1.Flags = cdlCFScreenFonts Or cdlCFEffects
    CommonDialog1.ShowFont

    txtMain.Font.Name = CommonDialog1.FontName
    txtMain.Font.Italic = CommonDialog1.FontItalic
    txtMain.Font.Bold = CommonDialog1.FontBold
    txtMain.Font.Underline = CommonDialog1.FontUnderline
    txtMain.Font.size = CommonDialog1.FontSize
    txtMain.ForeColor = CommonDialog1.Color
    Set txtMain.Font = txtMain.Font
    'txtMain.Redraw
    
    On Error GoTo 0
    
    Exit Sub
exitfontset:
    Err.Clear
    
    On Error GoTo 0
End Sub

Private Sub mnuTimeDate_Click()
    Dim ndt As String
    ndt = Now
    txtMain.Text = Left(txtMain.Text, txtMain.SelStart) & ndt & Mid(txtMain.Text, txtMain.SelStart + txtMain.SelLength + 1)
    txtMain.SelStart = txtMain.SelStart + Len(ndt) + 1
End Sub

Private Sub mnuWordWrap_Click()
    CancelChange = True
    txtMain.WordWrap = Not txtMain.WordWrap
    mnuWordWrap.Checked = txtMain.WordWrap
    CancelChange = False
End Sub

Private Sub WriteTextFile(ByVal FilePath As String, ByVal FileContents As String)
On Error GoTo errcatch
    Dim fso, f
    Const ForWriting = 2, ForAppending = 8
  
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(FilePath, ForWriting, True)
    f.Write FileContents
    f.Close
    
    Set f = Nothing
    Set fso = Nothing
    
    On Error GoTo 0
Exit Sub
errcatch:
    MsgBox "Unable to write file: " & Err.Description, vbInformation, "Notable Ink"
    Err.Clear
    On Error GoTo 0
End Sub
Private Function ReadTextFile(ByVal FilePath As String) As String
On Error GoTo errcatch

    Dim fso, f
    Const ForReading = 1
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(FilePath) Then
        Set f = fso.OpenTextFile(FilePath, ForReading)
        ReadTextFile = f.ReadAll
        f.Close
    End If
    
    Set f = Nothing
    Set fso = Nothing
    
    On Error GoTo 0
Exit Function
errcatch:
    MsgBox "Unable to read file: " & Err.Description, vbInformation, "Notable Ink"
    Err.Clear
    On Error GoTo 0
End Function

Private Function GetFilePath(ByVal FullFileName As String) As String
    'Returns everything but the last name in FullFileName, opposite of GetFileName
    If InStrRev(FullFileName, "\") > 0 Then
        FullFileName = Left(FullFileName, InStrRev(FullFileName, "\") - 1)
    ElseIf InStrRev(FullFileName, "/") > 0 Then
        FullFileName = Left(FullFileName, InStrRev(FullFileName, "/") - 1)
    End If
    GetFilePath = FullFileName
End Function

Private Function GetFileName(ByVal FullFileName As String) As String
    'Returns just the last name in FullFileName, opposite of GetFilePath
    If FullFileName = "" Then
        FullFileName = "Untitled"
    Else
        If InStrRev(FullFileName, "\") > 0 Then
            FullFileName = Mid(FullFileName, InStrRev(FullFileName, "\") + 1)
        ElseIf InStrRev(FullFileName, "/") > 0 Then
            FullFileName = Mid(FullFileName, InStrRev(FullFileName, "/") + 1)
        End If
    End If
    GetFileName = FullFileName
End Function

Private Function FileExists(ByVal FilePath As String) As Boolean
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(FilePath)
    Set fso = Nothing
End Function

Public Sub GotoLine(ByVal Number As Long)
    txtMain.SelectRow Number, False
    
End Sub

Private Function LoadINI()

    Me.Width = GetSetting("BasicNeotext", "Notable Ink", "Width", Me.Width)
    Me.Height = GetSetting("BasicNeotext", "Notable Ink", "Height", Me.Height)
    Me.Top = GetSetting("BasicNeotext", "Notable Ink", "Top", Me.Top)
    Me.Left = GetSetting("BasicNeotext", "Notable Ink", "Left", Me.Left)
    txtMain.WordWrap = GetSetting("BasicNeotext", "Notable Ink", "WordWrap", txtMain.WordWrap)
    txtMain.Font.Name = GetSetting("BasicNeotext", "Notable Ink", "FontName", txtMain.Font.Name)
    txtMain.Font.size = GetSetting("BasicNeotext", "Notable Ink", "FontSize", txtMain.Font.size)
    txtMain.Font.Italic = GetSetting("BasicNeotext", "Notable Ink", "FontItalic", CBool(txtMain.Font.Italic))
    txtMain.Font.Bold = GetSetting("BasicNeotext", "Notable Ink", "FontBold", CBool(txtMain.Font.Bold))
    txtMain.Font.StrikeThrough = GetSetting("BasicNeotext", "Notable Ink", "FontStrikethru", CBool(txtMain.Font.StrikeThrough))
    txtMain.Font.Underline = GetSetting("BasicNeotext", "Notable Ink", "FontUnderline", CBool(txtMain.Font.Underline))
    
    ColorProperty(BatchInkComment) = GetSetting("BasicNeotext", "Notable Ink", "BatchInkComment", &H8000&)
    ColorProperty(BatchInkCommands) = GetSetting("BasicNeotext", "Notable Ink", "BatchInkCommands", &H0&)
    ColorProperty(BatchInkFinished) = GetSetting("BasicNeotext", "Notable Ink", "BatchInkFinished", &H0&)
    ColorProperty(BatchInkCurrently) = GetSetting("BasicNeotext", "Notable Ink", "BatchInkCurrently", &HFFFFFF)
    ColorProperty(BatchInkIncomming) = GetSetting("BasicNeotext", "Notable Ink", "BatchInkIncomming", &H404040)
  
    ColorProperty(JScriptComment) = GetSetting("BasicNeotext", "Notable Ink", "JScriptComment", &H8000&)
    ColorProperty(JScriptStatements) = GetSetting("BasicNeotext", "Notable Ink", "JScriptStatements", &H800000)
    ColorProperty(JScriptOperators) = GetSetting("BasicNeotext", "Notable Ink", "JScriptOperators", &HFF&)
    ColorProperty(JScriptVariables) = GetSetting("BasicNeotext", "Notable Ink", "JScriptVariables", &H0&)
    ColorProperty(JScriptValues) = GetSetting("BasicNeotext", "Notable Ink", "JScriptValues", &H808080)
    ColorProperty(JScriptError) = GetSetting("BasicNeotext", "Notable Ink", "JScriptError", &HFF&)
  
    ColorProperty(VBScriptComment) = GetSetting("BasicNeotext", "Notable Ink", "VBScriptComment", &H8000&)
    ColorProperty(VBScriptStatements) = GetSetting("BasicNeotext", "Notable Ink", "VBScriptStatements", &H80&)
    ColorProperty(VBScriptOperators) = GetSetting("BasicNeotext", "Notable Ink", "VBScriptOperators", &H808080)
    ColorProperty(VBScriptVariables) = GetSetting("BasicNeotext", "Notable Ink", "VBScriptVariables", &H800000)
    ColorProperty(VBScriptValues) = GetSetting("BasicNeotext", "Notable Ink", "VBScriptValues", &H808080)
    ColorProperty(VBScriptError) = GetSetting("BasicNeotext", "Notable Ink", "VBScriptError", &HFF&)

    ColorProperty(NSISScriptComment) = GetSetting("BasicNeotext", "Notable Ink", "NSISScriptComment", &H8000&)
    ColorProperty(NSISScriptCommands) = GetSetting("BasicNeotext", "Notable Ink", "NSISScriptCommands", &H0&)
    ColorProperty(NSISScriptEqualJump) = GetSetting("BasicNeotext", "Notable Ink", "NSISScriptEqualJump", &H808080)
    ColorProperty(NSISScriptElseJump) = GetSetting("BasicNeotext", "Notable Ink", "NSISScriptElseJump", &H8000&)
    ColorProperty(NSISScriptAboveJump) = GetSetting("BasicNeotext", "Notable Ink", "NSISScriptAboveJump", &H80&)

    ColorProperty(AssemblyComment) = GetSetting("BasicNeotext", "Notable Ink", "AssemblyComment", &H8000&)
    ColorProperty(AssemblyCommand) = GetSetting("BasicNeotext", "Notable Ink", "AssemblyCommand", &H0&)
    ColorProperty(AssemblyNotation) = GetSetting("BasicNeotext", "Notable Ink", "AssemblyNotation", &H808080)
    ColorProperty(AssemblyRegister) = GetSetting("BasicNeotext", "Notable Ink", "AssemblyRegister", &H8000&)
    ColorProperty(AssemblyParameter) = GetSetting("BasicNeotext", "Notable Ink", "AssemblyParameter", &H80&)
    ColorProperty(AssemblyError) = GetSetting("BasicNeotext", "Notable Ink", "AssemblyError", &HFF&)
    
    ColorProperty(ColorBackground) = GetSetting("BasicNeotext", "Notable Ink", "Background", &HFFFFFF)
    ColorProperty(ColorForeGround) = GetSetting("BasicNeotext", "Notable Ink", "ForeGround", &H0&)

End Function

Private Function SaveINI()
    
    If Me.WindowState = 0 Then
        SaveSetting "BasicNeotext", "Notable Ink", "Width", Me.Width
        SaveSetting "BasicNeotext", "Notable Ink", "Height", Me.Height
        SaveSetting "BasicNeotext", "Notable Ink", "Top", Me.Top
        SaveSetting "BasicNeotext", "Notable Ink", "Left", Me.Left
    End If
    SaveSetting "BasicNeotext", "Notable Ink", "WordWrap", txtMain.WordWrap
    SaveSetting "BasicNeotext", "Notable Ink", "FontName", txtMain.Font.Name
    SaveSetting "BasicNeotext", "Notable Ink", "FontSize", txtMain.Font.size
    SaveSetting "BasicNeotext", "Notable Ink", "FontItalic", CBool(txtMain.Font.Italic)
    SaveSetting "BasicNeotext", "Notable Ink", "FontBold", CBool(txtMain.Font.Bold)
    SaveSetting "BasicNeotext", "Notable Ink", "FontStrikethru", CBool(txtMain.Font.StrikeThrough)
    SaveSetting "BasicNeotext", "Notable Ink", "FontUnderline", CBool(txtMain.Font.Underline)
    
    SaveSetting "BasicNeotext", "Notable Ink", "BackGround", ColorProperty(ColorBackground)
    SaveSetting "BasicNeotext", "Notable Ink", "ForeGround", ColorProperty(ColorForeGround)
    
    SaveSetting "BasicNeotext", "Notable Ink", "BatchInkComment", ColorProperty(BatchInkComment)
    SaveSetting "BasicNeotext", "Notable Ink", "BatchInkCommands", ColorProperty(BatchInkCommands)
    SaveSetting "BasicNeotext", "Notable Ink", "BatchInkFinished", ColorProperty(BatchInkFinished)
    SaveSetting "BasicNeotext", "Notable Ink", "BatchInkCurrently", ColorProperty(BatchInkCurrently)
    SaveSetting "BasicNeotext", "Notable Ink", "BatchInkIncomming", ColorProperty(BatchInkIncomming)
    
    SaveSetting "BasicNeotext", "Notable Ink", "JScriptComment", ColorProperty(JScriptComment)
    SaveSetting "BasicNeotext", "Notable Ink", "JScriptStatements", ColorProperty(JScriptStatements)
    SaveSetting "BasicNeotext", "Notable Ink", "JScriptOperators", ColorProperty(JScriptOperators)
    SaveSetting "BasicNeotext", "Notable Ink", "JScriptVariables", ColorProperty(JScriptVariables)
    SaveSetting "BasicNeotext", "Notable Ink", "JScriptValues", ColorProperty(JScriptValues)
    SaveSetting "BasicNeotext", "Notable Ink", "JScriptError", ColorProperty(JScriptError)
    
    SaveSetting "BasicNeotext", "Notable Ink", "VBScriptComment", ColorProperty(VBScriptComment)
    SaveSetting "BasicNeotext", "Notable Ink", "VBScriptStatements", ColorProperty(VBScriptStatements)
    SaveSetting "BasicNeotext", "Notable Ink", "VBScriptOperators", ColorProperty(VBScriptOperators)
    SaveSetting "BasicNeotext", "Notable Ink", "VBScriptVariables", ColorProperty(VBScriptVariables)
    SaveSetting "BasicNeotext", "Notable Ink", "VBScriptValues", ColorProperty(VBScriptValues)
    SaveSetting "BasicNeotext", "Notable Ink", "VBScriptError", ColorProperty(VBScriptError)

    SaveSetting "BasicNeotext", "Notable Ink", "NSISScriptComment", ColorProperty(NSISScriptComment)
    SaveSetting "BasicNeotext", "Notable Ink", "NSISScriptCommands", ColorProperty(NSISScriptCommands)
    SaveSetting "BasicNeotext", "Notable Ink", "NSISScriptEqualJump", ColorProperty(NSISScriptEqualJump)
    SaveSetting "BasicNeotext", "Notable Ink", "NSISScriptElseJump", ColorProperty(NSISScriptElseJump)
    SaveSetting "BasicNeotext", "Notable Ink", "NSISScriptAboveJump", ColorProperty(NSISScriptAboveJump)
    

    SaveSetting "BasicNeotext", "Notable Ink", "AssemblyComment", ColorProperty(AssemblyComment)
    SaveSetting "BasicNeotext", "Notable Ink", "AssemblyCommand", ColorProperty(AssemblyCommand)
    SaveSetting "BasicNeotext", "Notable Ink", "AssemblyNotation", ColorProperty(AssemblyNotation)
    SaveSetting "BasicNeotext", "Notable Ink", "AssemblyRegister", ColorProperty(AssemblyRegister)
    SaveSetting "BasicNeotext", "Notable Ink", "AssemblyParameter", ColorProperty(AssemblyParameter)
    SaveSetting "BasicNeotext", "Notable Ink", "AssemblyError", ColorProperty(AssemblyError)
End Function

Private Function NextArg(ByVal TheParams As String, ByVal TheSeperator As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare, Optional ByVal TrimResult As Boolean = True) As String
    If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
        NextArg = Left(TheParams, InStr(1, TheParams, TheSeperator, Compare) - 1)
    Else
        NextArg = TheParams
    End If
    If TrimResult Then NextArg = Trim(NextArg)
End Function

Private Function RemoveArg(ByVal TheParams As String, ByVal TheSeperator As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare, Optional ByVal TrimResult As Boolean = True) As String
    If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
        RemoveArg = Mid(TheParams, InStr(1, TheParams, TheSeperator, Compare) + Len(TheSeperator), Len(TheParams) - Len(TheSeperator))
    Else
        RemoveArg = ""
    End If
    If TrimResult Then RemoveArg = Trim(RemoveArg)
End Function

Private Function RemoveNextArg(ByRef TheParams As Variant, ByVal TheSeperator As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare, Optional ByVal TrimResult As Boolean = True) As String
    If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
        RemoveNextArg = Left(TheParams, InStr(1, TheParams, TheSeperator, Compare) - 1)
        TheParams = Mid(TheParams, InStr(1, TheParams, TheSeperator, Compare) + Len(TheSeperator))
    Else
        RemoveNextArg = TheParams
        TheParams = ""
    End If
    If TrimResult Then
        RemoveNextArg = Trim(RemoveNextArg)
        TheParams = Trim(TheParams)
    End If
    
End Function


Private Sub ColorBatchFile()
    If GetLanguage = "Defined" And Not Running Then
    
        Dim line As Long
        Dim txt As String
        Dim dirty As Boolean
        Dim Start As String


        txt = txtMain.Text
        Start = txt
    
        Do Until (txt = "") Or (txtMain.Text <> Start)
            line = line + 1
            If Left(LCase(Trim(Replace(RemoveNextArg(txt, vbCrLf), vbTab, ""))), 3) = "rem" Then
                txtMain.LineDefines(line) = Dream1Index
            ElseIf Left(LCase(Trim(Replace(RemoveNextArg(txt, vbCrLf), vbTab, ""))), 1) = ";" Then
            Else
                txtMain.LineDefines(line) = Dream1Index
            End If

        Loop

        txtMain.ColorDream1 = ColoringIndexes.CommentIndex
        txtMain.ColorDream2 = ColoringIndexes.StatementIndex
            
        If GetFileExt(TextFileName, True, True) = "asm" Then
            txtMain.ColorDream1 = ColorProperty(AssemblyComment)
            txtMain.ColorDream2 = ColorProperty(AssemblyCommand)
            txtMain.ColorDream3 = ColorProperty(AssemblyNotation)
            txtMain.ColorDream4 = ColorProperty(AssemblyRegister)
            txtMain.ColorDream5 = ColorProperty(AssemblyParameter)
            txtMain.ColorDream6 = ColorProperty(AssemblyError)
        ElseIf GetFileExt(TextFileName, True, True) = "nsi" Or GetFileExt(TextFileName, True, True) = "nsh" Then
            txtMain.ColorDream1 = ColoringIndexes.CommentIndex
            txtMain.ColorDream2 = ColoringIndexes.StatementIndex
        Else
            txtMain.ColorDream1 = ColorProperty(BatchInkComment)
            txtMain.ColorDream2 = ColorProperty(BatchInkCommands)
        End If
        txtMain.BackColor = txtMain.BackColor

    End If
 
End Sub

Private Sub Timer1_Timer()
    If Not frmMain.Visible Then
        Timer1.Interval = 0
        End
    Else
        If elapsedTime >= 0 And (Not CancelChange) Then
            If Timer - elapsedTime > 0.2 Then
                elapsedTime = -1
                If (((GetFileExt(TextFileName, True, True) = "nsi") Or (GetFileExt(TextFileName, True, True) = "nsh"))) Then
            
                    Dim cursorLine As String
                    Dim oncolor As Integer
                    Dim LastPos As Long
                    Dim nextJump As String
            
                    cursorLine = StrReverse(NextArg(StrReverse(Left(txtMain.Text, txtMain.SelStart)), vbLf & vbCr, , False)) + NextArg(Mid(txtMain.Text, txtMain.SelStart + 1), vbCrLf, , False)
                    If InStr(cursorLine, "+") = 0 And InStr(cursorLine, "-") = 0 Then
                        LastPos = txtMain.SelStart
                        LastPos = InStrRev(Left(txtMain.Text, LastPos), "+")
                        If LastPos <= 0 Then
                            LastPos = txtMain.SelStart
                            LastPos = InStrRev(Left(txtMain.Text, LastPos), "-")
                            If LastPos <= 0 Then
                                LastPos = txtMain.SelStart
                            ElseIf (InStr(Mid(txtMain.Text, LastPos + 1), vbLf) > InStr(Mid(txtMain.Text, LastPos + 1), "+")) Or _
                                (InStr(Mid(txtMain.Text, LastPos + 1), vbLf) > InStr(Mid(txtMain.Text, LastPos + 1), "-")) Then
                                LastPos = txtMain.SelStart
                            End If
                        ElseIf (InStr(Mid(txtMain.Text, LastPos + 1), vbLf) > InStr(Mid(txtMain.Text, LastPos + 1), "+")) Or _
                            (InStr(Mid(txtMain.Text, LastPos + 1), vbLf) > InStr(Mid(txtMain.Text, LastPos + 1), "-")) Then
                            LastPos = txtMain.SelStart
                        End If
                    Else
                        LastPos = txtMain.SelStart
                    End If
            
                    If LastPos > 0 Then
                        highLitedLine = (CountWord(Left(txtMain.Text, LastPos), vbLf) + 1)
                        diffLitedLine = (CountWord(Left(txtMain.Text, txtMain.SelStart), vbLf) + 1) - (CountWord(Left(txtMain.Text, LastPos), vbLf) + 1)
                        If highLitedLine > 0 Then
                            cursorLine = StrReverse(NextArg(StrReverse(Left(txtMain.Text, LastPos)), vbLf & vbCr, , False)) + NextArg(Mid(txtMain.Text, LastPos + 1), vbCrLf, , False)
                            oncolor = 1
    
                            jumpClr(1) = -1
                            jumpClr(2) = -1
                            jumpClr(3) = -1
                            Dim lNum As Long
                            Dim lAt As Long
                            lNum = txtMain.LineNumber
            
                            Do While (InStr(cursorLine, "+") > 0 Or InStr(cursorLine, "-") > 0) And (oncolor <= 3) And (Not CancelChange)
                                If (InStr(cursorLine, "+") > 0) Then
                                    RemoveNextArg cursorLine, "+", , False
                                    If IsNumeric(Left(cursorLine, 1)) Then
                                        nextJump = NextArg(cursorLine, " ")
                                        If IsNumeric(nextJump) Then
                                            If nextJump = "0" Then nextJump = "1"
                                            lAt = lNum + CLng(nextJump) + 1 - diffLitedLine
                                            txtMain.LineDefines(lAt) = (Dream2Index + oncolor)
                                            jumpClr(oncolor) = lAt
                                            oncolor = oncolor + 1
                                        End If
                                    End If
                                ElseIf (InStr(cursorLine, "-") > 0) Then
                                    RemoveNextArg cursorLine, "-", , False
                                    If IsNumeric(Left(cursorLine, 1)) Then
                                        nextJump = NextArg(cursorLine, " ")
                                        If IsNumeric(nextJump) Then
                                            If nextJump = "0" Then nextJump = "1"
                                            lAt = txtMain.LineNumber + CLng(nextJump) + 1 - diffLitedLine
                                            txtMain.LineDefines(lAt) = (Dream2Index + oncolor)
                                            jumpClr(oncolor) = lAt
                                            oncolor = oncolor + 1
                                        End If
                                    End If
                                End If
            
                            Loop
                            txtMain.LineDefines(-1) = -1
                            
                        End If
            
                    End If
    
                    txtMain.ColorDream1 = txtMain.ColorDream1
                    txtMain.ColorDream2 = txtMain.ColorDream2
                    txtMain.ColorDream3 = txtMain.ColorDream3
                    txtMain.ColorDream4 = txtMain.ColorDream4
                    txtMain.ColorDream5 = txtMain.ColorDream5
                    txtMain.ColorDream6 = txtMain.ColorDream6
                    
                End If
            End If
        End If
    End If
End Sub

Private Sub txtMain_Change()

    elapsedTime = Timer
    pTextChanged = True
    SetCaption
   
End Sub

Private Sub txtMain_QueryDefine(ByVal linenum As Long, ByRef LineText As String, ColorIndex As ColoringIndexes)
    
    If (Not Running) And (Not CancelChange) Then
        If (GetFileExt(TextFileName, True, True) = "asm") Then
        
            
        
            'ColorIndex = Dream1Index
            
            
        ElseIf (GetFileExt(TextFileName, True, True) = "bat") Then
            ColorIndex = IIf(Left(LCase(Trim(Replace(LineText, vbTab, ""))), 3) = "rem", Dream1Index, Dream2Index)
        ElseIf (GetFileExt(TextFileName, True, True) = "nsi") Or (GetFileExt(TextFileName, True, True) = "nsh") Then
            If Left(LCase(Trim(Replace(LineText, vbTab, ""))), 3) = ";" Then
                ColorIndex = Dream1Index
            ElseIf highLitedLine <> linenum Then
    
                If linenum = jumpClr(1) Then
                    ColorIndex = Dream3Index
                ElseIf linenum = jumpClr(2) Then
                    ColorIndex = Dream4Index
                ElseIf linenum = jumpClr(3) Then
                    ColorIndex = Dream5Index
                Else
                    ColorIndex = Dream2Index
                End If

            End If
        End If
    End If

End Sub

Private Sub txtMain_SelChange()
    elapsedTime = Timer

End Sub
