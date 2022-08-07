VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C98B112F-745F-4542-B5B3-DDFADF1F6E2F}#1059.0#0"; "NTControls22.ocx"
Begin VB.Form frmMain 
   Caption         =   "RemindMe"
   ClientHeight    =   8115
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11535
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8115
   ScaleWidth      =   11535
   Begin NTControls22.CodeEdit CodeEdit1 
      Height          =   2085
      Left            =   720
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3210
      Width           =   6090
      _ExtentX        =   10742
      _ExtentY        =   3678
      FontSize        =   9.75
      BackColor       =   16777215
      ColorDream1     =   8388736
      ColorDream2     =   8388608
      ColorDream3     =   8421376
      ColorDream4     =   32768
      ColorDream5     =   32896
      ColorDream6     =   16512
   End
   Begin VB.TextBox txtDebug 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   285
      Left            =   2595
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6105
      Width           =   5175
   End
   Begin VB.PictureBox hSizer 
      BorderStyle     =   0  'None
      Height          =   100
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   105
      ScaleWidth      =   7425
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2880
      Width           =   7425
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   -15
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5415
      Top             =   990
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   "item"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstObjects 
      Height          =   2025
      Left            =   60
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   465
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   3572
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   5821
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Enabled"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Procedure"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Type"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Schedule"
         Object.Width           =   5821
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   794
      ButtonWidth     =   820
      ButtonHeight    =   794
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ToolImgsOut"
      DisabledImageList=   "ToolImgsOut"
      HotImageList    =   "ToolImgsOver"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "servicestop"
            Object.ToolTipText     =   "Stop the RemindMe Service"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "servicestart"
            Object.ToolTipText     =   "Starts the RemindMe Service"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "servicepause"
            Object.ToolTipText     =   "The RemindMe Service is Paused"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "add"
            Object.ToolTipText     =   "Add Operation"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "edit"
            Object.ToolTipText     =   "Edit Operation"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "delete"
            Object.ToolTipText     =   "Delete Operation"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "runselected"
            Object.ToolTipText     =   "Run Only Selected Operations"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "stop"
            Object.ToolTipText     =   "Stop All Operations"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Object.ToolTipText     =   "Update Code"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "enabled"
            Object.ToolTipText     =   "Enable or disable a selected operation."
            ImageIndex      =   14
         EndProperty
      EndProperty
      MousePointer    =   1
      Begin MSComctlLib.ImageList ToolImgsOver 
         Left            =   6795
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   24
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0894
               Key             =   "up"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0C3B
               Key             =   "servicepause"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":12CD
               Key             =   "servicestop"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":16AA
               Key             =   "servicestart"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1A86
               Key             =   "delete"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1E56
               Key             =   "down"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":21FA
               Key             =   "edit"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":25DC
               Key             =   "open"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2996
               Key             =   "runall"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2D68
               Key             =   "runselected"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":313E
               Key             =   "save"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3507
               Key             =   "stop"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":38E0
               Key             =   "add"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3CBA
               Key             =   "enabled"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ToolImgsOut 
         Left            =   6045
         Top             =   15
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   24
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":434C
               Key             =   "up"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":46EF
               Key             =   "servicepause"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4D81
               Key             =   "servicestop"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5141
               Key             =   "servicestart"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5503
               Key             =   "delete"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":58C6
               Key             =   "down"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5C69
               Key             =   "edit"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":603C
               Key             =   "open"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":63EE
               Key             =   "runall"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":67AA
               Key             =   "runselected"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6B6A
               Key             =   "save"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6F26
               Key             =   "stop"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":72F3
               Key             =   "add"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":76C7
               Key             =   "enabled"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuOperation 
      Caption         =   "&Operation"
      Begin VB.Menu mnuWizard 
         Caption         =   "&Wizard"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuConfig 
         Caption         =   "&Config"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remo&ve"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuDash6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDisable 
         Caption         =   "Disa&ble"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuDash237 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRunSelected 
         Caption         =   "&Run"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuScript 
      Caption         =   "&Script"
      Begin VB.Menu mnuViewScript 
         Caption         =   "View &Script"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuDash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUpdate 
         Caption         =   "&Update"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVBScript 
         Caption         =   "&VBScript"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuJavaScript 
         Caption         =   "&JScript"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Redo"
      End
      Begin VB.Menu mnuDash8 
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
      Begin VB.Menu mnuDash9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "De&lete"
      End
      Begin VB.Menu mnuDash2187 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComment 
         Caption         =   "C&omment"
      End
      Begin VB.Menu mnuUncomment 
         Caption         =   "Uncomm&ent"
      End
      Begin VB.Menu mnuDash10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "&Find or Replace.."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuDocumentation 
         Caption         =   "&Documentation..."
      End
      Begin VB.Menu mnuNeotext 
         Caption         =   "&Neotext.org..."
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
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

Private NoReSize As Boolean
Private hBarSizing As Boolean

Private WithEvents Timer1 As NTSchedule20.Timer
Attribute Timer1.VB_VarHelpID = -1

Private Sub CodeEdit1_Change()
    CodeEdit1.ErrorLine = 0
End Sub

Public Function CountWord(ByVal Text As String, ByVal Word As String) As Long
    Dim cnt As Long
    Dim pos As Long
    cnt = 0
    pos = InStr(Text, Word)
    Do Until pos = 0
        cnt = cnt + 1
        pos = InStr(pos + 1, Text, Word)
    Loop
    CountWord = cnt
End Function

Private Sub CodeEdit1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        RefreshEditMenu
        Me.PopupMenu frmMain.mnuEdit
    End If
End Sub

Public Function RefreshEditMenu()
    With frmMain
        .mnuUndo.Enabled = CodeEdit1.CanUndo
        .mnuRedo.Enabled = CodeEdit1.CanRedo
            
        .mnuCut.Enabled = (CodeEdit1.SelLength > 0)
        .mnuCopy.Enabled = (CodeEdit1.SelLength > 0)
        
        .mnuPaste.Enabled = (Len(Clipboard.GetText(vbCFText)) > 0)
        
        .mnuDelete.Enabled = (CodeEdit1.SelLength > 0)
        .mnuSelectAll.Enabled = (Len(CodeEdit1.Text) > 0)
        
        .mnuComment.Enabled = (CodeEdit1.SelStart > 0)
        .mnuUncomment.Enabled = (CodeEdit1.SelStart > 0)
        
        .mnuFind.Enabled = CodeEdit1.Visible
    End With
End Function

Private Sub Form_Activate()
    mnuEdit.Visible = CodeEdit1.Visible
    SetDebugPreview
End Sub

Private Sub SetScriptSweets()

    Select Case CodeEdit1.Language
        Case "VBScript"
            CodeEdit1.ColorComment = CodeEdit1.ConvertColor("008000")
            CodeEdit1.ColorError = CodeEdit1.ConvertColor("FF0000")
            CodeEdit1.ColorOperator = CodeEdit1.ConvertColor("202020")
            CodeEdit1.ColorStatement = CodeEdit1.ConvertColor("800000")
            CodeEdit1.ColorText = CodeEdit1.ConvertColor("000000")
            CodeEdit1.ColorVariable = CodeEdit1.ConvertColor("000080")
            CodeEdit1.ColorValue = CodeEdit1.ConvertColor("808080")
        Case "JScript"
            CodeEdit1.ColorComment = CodeEdit1.ConvertColor("008000")
            CodeEdit1.ColorError = CodeEdit1.ConvertColor("FF0000")
            CodeEdit1.ColorOperator = CodeEdit1.ConvertColor("800000")
            CodeEdit1.ColorStatement = CodeEdit1.ConvertColor("000080")
            CodeEdit1.ColorText = CodeEdit1.ConvertColor("000000")
            CodeEdit1.ColorVariable = CodeEdit1.ConvertColor("000000")
            CodeEdit1.ColorValue = CodeEdit1.ConvertColor("808080")
    End Select
    
End Sub
Private Sub Form_Load()
    Set Timer1 = New NTSchedule20.Timer
    Timer1.Interval = 250
    Timer1.Enabled = True
        
    NoReSize = True
    
    CodeEdit1.Visible = dbSettings.GetSetting("ShowScript")
    txtDebug.Visible = dbSettings.GetSetting("ShowScript")
    mnuViewScript.Checked = dbSettings.GetSetting("ShowScript")
    mnuEdit.Visible = dbSettings.GetSetting("ShowScript")
    
    Toolbar1.Buttons(11).Visible = dbSettings.GetSetting("ShowScript")
    Toolbar1.Buttons(12).Visible = dbSettings.GetSetting("ShowScript")
    
    CodeEdit1.Language = dbSettings.GetSetting("Language")
    mnuJavaScript.Checked = (CodeEdit1.Language = "JScript")
    mnuVBScript.Checked = (CodeEdit1.Language = "VBScript")

    SetScriptSweets
    
    Me.Top = dbSettings.GetSetting("wTop")
    Me.Left = dbSettings.GetSetting("wLeft")
    Me.Height = dbSettings.GetSetting("wHeight")
    Me.Width = dbSettings.GetSetting("wWidth")
    Me.WindowState = dbSettings.GetSetting("wState")
    hSizer.Top = dbSettings.GetSetting("hSizer")
    lstObjects.ColumnHeaders(1).Width = dbSettings.GetSetting("wColumn1")
    lstObjects.ColumnHeaders(2).Width = dbSettings.GetSetting("wColumn2")
    lstObjects.ColumnHeaders(3).Width = dbSettings.GetSetting("wColumn3")
    lstObjects.ColumnHeaders(4).Width = dbSettings.GetSetting("wColumn4")
    lstObjects.ColumnHeaders(5).Width = dbSettings.GetSetting("wColumn5")
   
    NoReSize = False
    
    RefreshOperations

    CodeEdit1.Text = dbSettings.GetSetting(CodeEdit1.Language & "Text") & ""
    
    CodeEdit1.UndoDirty = False
    
    UpdateScriptCode
    
    ChangeServiceButton
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If NoReSize Or Me.WindowState = 1 Then Exit Sub
    
    Const Border = 60
        
    If CodeEdit1.Visible Then
        
        hSizer.Left = Border
        hSizer.Width = Me.ScaleWidth - (Border * 2)
        
        If hSizer.Top < lstObjects.Top Then
            If lstObjects.Visible <> False Then lstObjects.Visible = False
            hSizer.Top = lstObjects.Top
        Else
            If lstObjects.Visible <> True Then lstObjects.Visible = True
        End If
        
        If hSizer.Top + hSizer.Height > Me.ScaleHeight - Border Then
            If CodeEdit1.Visible <> False Then
                CodeEdit1.Visible = False
                txtDebug.Visible = False
            End If
            hSizer.Top = Me.ScaleHeight - hSizer.Height - Border
        Else
            If CodeEdit1.Visible <> True Then
                CodeEdit1.Visible = True
                txtDebug.Visible = True
            End If
        End If
        
    End If
    
    If Not hSizer.Visible = CodeEdit1.Visible Then hSizer.Visible = CodeEdit1.Visible
    
    lstObjects.Top = Toolbar1.Height + Border
    lstObjects.Left = Border
    lstObjects.Width = Me.ScaleWidth - (Border * 2)
    If CodeEdit1.Visible Then
        lstObjects.Height = hSizer.Top - lstObjects.Top - 20
    Else
        lstObjects.Height = Me.ScaleHeight - lstObjects.Top - Border
    End If
    
    If CodeEdit1.Visible Then
        CodeEdit1.Top = hSizer.Top + hSizer.Height + 20
        CodeEdit1.Left = Border
        CodeEdit1.Width = Me.ScaleWidth - (Border * 2)
        CodeEdit1.Height = Me.ScaleHeight - CodeEdit1.Top - txtDebug.Height - (Border * 3)
        
        txtDebug.Top = Me.ScaleHeight - txtDebug.Height - Border
        txtDebug.Left = Border
        txtDebug.Width = Me.ScaleWidth - (Border * 2)
        
    End If
    
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim tmp As Integer
    
    If CodeEdit1.UndoDirty Then
        tmp = MsgBox("Changes to the script have not been saved, do you want to save?", vbYesNoCancel, AppName)
        If tmp = vbYes Then
            UpdateScriptCode
        End If
        Cancel = (tmp = vbCancel)
    End If
    
    If Not Cancel Then
        ClearRemindMeScript
        
        CodeEdit1.Visible = False
        
        dbSettings.SetSetting "Language", CodeEdit1.Language
        dbSettings.SetSetting "wTop", Me.Top
        dbSettings.SetSetting "wLeft", Me.Left
        dbSettings.SetSetting "wHeight", Me.Height
        dbSettings.SetSetting "wWidth", Me.Width
        dbSettings.SetSetting "wState", IIf(Me.WindowState = 1, 0, Me.WindowState)
        dbSettings.SetSetting "hSizer", hSizer.Top
        dbSettings.SetSetting "wColumn1", lstObjects.ColumnHeaders(1).Width
        dbSettings.SetSetting "wColumn2", lstObjects.ColumnHeaders(2).Width
        dbSettings.SetSetting "wColumn3", lstObjects.ColumnHeaders(3).Width
        dbSettings.SetSetting "wColumn4", lstObjects.ColumnHeaders(4).Width
        dbSettings.SetSetting "wColumn5", lstObjects.ColumnHeaders(5).Width
        
        dbSettings.SetSetting CodeEdit1.Language & "Text", TrimStrip(TrimStrip(Replace(CodeEdit1.Text, "'", "''"), vbCrLf), Chr(13))
        
        Timer1.Enabled = False
        Set Timer1 = Nothing
        
#If VBIDE Then
        End
#Else
        KillApp RemindMeFileName
#End If
    End If
End Sub

Private Sub hSizer_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 And Not hBarSizing Then
        hBarSizing = True
        hSizer.BackColor = &H808080
    Else
        If Button = 1 And hBarSizing Then
            
            hSizer.Top = hSizer.Top + Y
            
            Form_Resize
            Me.Refresh
        Else
            Form_Resize
            hBarSizing = False
            hSizer.BackColor = &H8000000F
        End If
    End If
End Sub

Public Sub SetLanguage(ByVal Language As String)
    Language = LCase(Language)
    If Language = "jscript" Or Language = "javascript" Or Language = "js" Then
        mnuJavaScript_Click
    ElseIf Language = "vbscript" Or Language = "visualbasicscript" Or Language = "visualbscript" Or Language = "vbasicscript" Or Language = "vbs" Then
        mnuVBScript_Click
    End If
End Sub

Private Sub lstObjects_DblClick()
    If Not lstObjects.SelectedItem Is Nothing Then
        EditOperation
    End If
End Sub

Private Sub lstObjects_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Select Case Button
        Case 2
            RefreshOperationMenu True
            Me.PopupMenu mnuOperation
            mnuDash2.Visible = True
            mnuExit.Visible = True
    End Select
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuComment_Click()
    CodeEdit1.Comment True
    
End Sub

Private Sub mnuDocumentation_Click()
    If PathExists(AppPath & "Help\index.htm", True) Then
        OpenWebsite AppPath & "help\index.htm", False
    Else
        OpenWebsite WebSite & "/ipub/help/remindme"
    End If
End Sub

Private Sub mnuEdit_Click()
    RefreshEditMenu
End Sub

Private Sub mnuFind_Click()
    frmFind.Show
End Sub

Private Sub mnuRedo_Click()
    CodeEdit1.Redo
End Sub
Private Sub RefreshScriptMenu()
    mnuJavaScript.Checked = (CodeEdit1.Language = "JScript")
    mnuVBScript.Checked = (CodeEdit1.Language = "VBScript")
    mnuViewScript.Checked = CodeEdit1.Visible
    mnuUpdate.Enabled = CodeEdit1.Visible
    mnuJavaScript.Enabled = CodeEdit1.Visible
    mnuVBScript.Enabled = CodeEdit1.Visible
End Sub


Private Sub mnuRunSelected_Click()
    RunSelected
End Sub

Private Sub mnuScript_Click()
    RefreshScriptMenu
End Sub

Private Sub mnuUncomment_Click()
    CodeEdit1.Comment False
End Sub

Private Sub mnuViewScript_Click()
    CodeEdit1.Visible = (Not CodeEdit1.Visible)
    txtDebug.Visible = CodeEdit1.Visible
    mnuEdit.Visible = CodeEdit1.Visible
    
    Toolbar1.Buttons(11).Visible = CodeEdit1.Visible
    Toolbar1.Buttons(12).Visible = CodeEdit1.Visible
    
    dbSettings.SetSetting "ShowScript", CodeEdit1.Visible
    If CodeEdit1.Visible Then
        hSizer.Top = (Me.ScaleHeight - Toolbar1.Height) / 2
    End If
    
    Form_Resize
    mnuEdit.Visible = CodeEdit1.Visible
    RefreshScriptMenu
End Sub

Private Sub mnuJavaScript_Click()
    If CodeEdit1.UndoDirty Then
        dbSettings.SetSetting CodeEdit1.Language & "Text", TrimStrip(TrimStrip(Replace(CodeEdit1.Text, "'", "''"), vbCrLf), Chr(13))
        CodeEdit1.UndoDirty = False
    End If
    
    CodeEdit1.Language = "JScript"
    dbSettings.SetSetting "Language", "JScript"
    CodeEdit1.Text = dbSettings.GetSetting(CodeEdit1.Language & "Text")
    
    If UpdateScriptCode Then
        SendMessage "updatescript"
    End If
    RefreshScriptMenu
    
    SetScriptSweets
End Sub

Private Sub mnuVBScript_Click()
    If CodeEdit1.UndoDirty Then
        dbSettings.SetSetting CodeEdit1.Language & "Text", TrimStrip(TrimStrip(Replace(CodeEdit1.Text, "'", "''"), vbCrLf), Chr(13))
        CodeEdit1.UndoDirty = False
    End If
    
    CodeEdit1.Language = "VBScript"
    dbSettings.SetSetting "Language", "VBScript"
    CodeEdit1.Text = dbSettings.GetSetting(CodeEdit1.Language & "Text")
    
    If UpdateScriptCode Then
        SendMessage "updatescript"
    End If
    RefreshScriptMenu
    
    SetScriptSweets
End Sub

Private Sub mnuUndo_Click()
    CodeEdit1.Undo
End Sub

Private Sub mnuDelete_Click()
    If CodeEdit1.SelLength = 0 Then
        CodeEdit1.SelLength = 1
    End If
    CodeEdit1.SelText = ""
End Sub
Private Sub mnuSelectAll_Click()
    CodeEdit1.SelStart = 0
    CodeEdit1.SelLength = Len(CodeEdit1.Text)
End Sub

Private Sub mnuCopy_Click()
    Clipboard.Clear
    Clipboard.SetText CodeEdit1.SelText, vbCFText
    Clipboard.SetText CodeEdit1.SelText, vbCFRTF
End Sub
Private Sub mnuCut_Click()
    Clipboard.Clear
    Clipboard.SetText CodeEdit1.SelText, vbCFText
    Clipboard.SetText CodeEdit1.SelText, vbCFRTF
    CodeEdit1.SelText = ""
End Sub
Private Sub mnuPaste_Click()
    CodeEdit1.SelText = Clipboard.GetText(vbCFText)
    CodeEdit1.SelStart = CodeEdit1.SelStart + Len(Clipboard.GetText(vbCFText))
End Sub

Private Sub mnuRemove_Click()
    DeleteOperation
End Sub

Private Sub mnuDisable_Click()
    ChangeEnabled
End Sub

Private Sub mnuConfig_Click()
    EditOperation
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuNeotext_Click()
    RunFile AppPath & "Neotext.org.url"

End Sub

Private Sub mnuWizard_Click()
    NewOperation
End Sub

Private Sub mnuOperation_Click()
    RefreshOperationMenu False
End Sub
Public Function RefreshOperationMenu(Optional ByVal PopMnu As Boolean = False)
    Dim HasSelected As Boolean
    HasSelected = (Not lstObjects.SelectedItem Is Nothing)
    mnuConfig.Enabled = HasSelected

    mnuRemove.Enabled = HasSelected

    mnuDisable.Enabled = HasSelected

    If HasSelected Then
        Select Case lstObjects.SelectedItem.SubItems(1)
            Case "True"
                mnuDisable.Caption = "Disa&ble"

            Case "False"
                mnuDisable.Caption = "Ena&ble"

        End Select
    End If
    
    mnuRunSelected.Enabled = HasSelected

    On Error Resume Next
    mnuDash2.Visible = Not PopMnu
    mnuExit.Visible = Not PopMnu
    Err.Clear

End Function

Private Sub DisplayError(ByVal proc As String, ByVal numb As Long, ByVal line As Long, ByVal colu As Long, ByVal sour As String, ByVal desc As String)

    If numb > 0 Then
        
        MsgBox "An error occured while trying to run a scripted procedure." & vbCrLf & vbCrLf & _
                "Procedure: " & proc & vbCrLf & _
                "Number: " & numb & vbCrLf & _
                IIf(sour <> "", "Source: " & sour & vbCrLf, "") & _
                IIf(line > 0, "Line: " & line & vbCrLf, "") & _
                IIf(colu > 0, "Column: " & colu & vbCrLf, "") & _
                "Description: " & desc & vbCrLf, vbCritical, AppName
                
        If CodeEdit1.Visible And (line > 0) Then
            CodeEdit1.SelectRow line, colu
            CodeEdit1.ErrorLine = line
        End If

    End If
    
End Sub

Public Sub RunProcedure(ByVal ProcedureName As String)
    On Error GoTo catch
    
    ScriptControl1.ExecuteStatement ProcedureName

    Exit Sub
catch:
    If (ScriptControl1.Error.Number = 0) Then
        DisplayError ProcedureName, Err.Number, 0, 0, Err.Source, Err.Description
    Else
        DisplayError ProcedureName, ScriptControl1.Error.Number, ScriptControl1.Error.line, ScriptControl1.Error.Column, ScriptControl1.Error.Source, ScriptControl1.Error.Description
    End If
    Err.Clear
End Sub
Public Function UpdateScriptCode() As Boolean
    On Error GoTo catch
    
    ClearRemindMeScript
    ScriptControl1.Reset
    ScriptControl1.Language = CodeEdit1.Language
    ScriptControl1.AddCode CodeEdit1.Text
    
    Dim lineNum As Long
    lineNum = GetRemindMeScript(CodeEdit1.Language, CodeEdit1.Text)
    If Not lineNum = 0 Then
        Err.Raise 93, App.EXEName, "Invalid pattern string (RemindMe comment parse error)"
    End If

    UpdateScriptCode = True
    Exit Function
catch:
    DoTasks
    If (ScriptControl1.Error.Number = 0) Then
        DisplayError "N/A or Unknown", Err.Number, 0, 0, Err.Source, Err.Description
    Else
        DisplayError "N/A or Unknown", ScriptControl1.Error.Number, ScriptControl1.Error.line, ScriptControl1.Error.Column, ScriptControl1.Error.Source, ScriptControl1.Error.Description
    End If
    Err.Clear
    UpdateScriptCode = False
End Function

Private Sub mnuUpdate_Click()
    
    If UpdateScriptCode Then
    
        If CodeEdit1.UndoDirty Then
            dbSettings.SetSetting CodeEdit1.Language & "Text", TrimStrip(TrimStrip(Replace(CodeEdit1.Text, "'", "''"), vbCrLf), Chr(13))
            CodeEdit1.UndoDirty = False
        End If
        
        SendMessage "updatescript"
    
    End If
    
End Sub

Public Function ChangeServiceButton()
    If Not Toolbar1.Buttons("servicepause").Visible Then
        Dim isRunning As Boolean
        Debug.Print ProcessRunning(ServiceFileName)
        
        isRunning = (ProcessRunning(ServiceFileName) > 0)
        Toolbar1.Buttons("servicestop").Visible = isRunning
        Toolbar1.Buttons("servicestart").Visible = Not isRunning
    End If
End Function

Private Sub ProcessMessage(ByVal Msg As String)
    Dim inCmd As String
    Dim inParam As String
    
    inParam = Msg
    inCmd = RemoveNextArg(inParam, ":")
    
    Select Case LCase(inCmd)
        Case "error"
            Dim proc As String
            Dim numb As Long
            Dim line As Long
            Dim colu As Long
            Dim sour As String
            Dim desc As String
            
            proc = RemoveNextArg(inParam, ":proc" & vbCrLf)
            numb = CLng(RemoveNextArg(inParam, ":numb" & vbCrLf))
            line = CLng(RemoveNextArg(inParam, ":line" & vbCrLf))
            colu = CLng(RemoveNextArg(inParam, ":colu" & vbCrLf))
            sour = RemoveNextArg(inParam, ":sour" & vbCrLf)
            desc = RemoveNextArg(inParam, ":desc" & vbCrLf)
            
            DisplayError proc, numb, line, colu, sour, desc
        Case "status"
            Select Case inParam
                Case "paused"
                    Toolbar1.Buttons("servicepause").Visible = True
                    Toolbar1.Buttons("servicestop").Visible = False
                    Toolbar1.Buttons("servicestart").Visible = False
                Case "resumed"
                    Toolbar1.Buttons("servicepause").Visible = False
                    Toolbar1.Buttons("servicestop").Visible = True
                    Toolbar1.Buttons("servicestart").Visible = False
            End Select
    End Select
End Sub

Private Sub Timer1_OnTicking()
    ChangeServiceButton
    
    If CodeEdit1.UndoDirty And Not Me.Caption = AppName & "*" Then
        Me.Caption = AppName & "*"
    ElseIf Not CodeEdit1.UndoDirty And Not Me.Caption = AppName Then
        Me.Caption = AppName
    End If
    
    If Me.Visible Then
        If dbSettings.MessageWaiting(RemindMeFileName) Then
            Dim NextMsg As String
            Dim Msgs As Collection
            Set Msgs = dbSettings.MessageQueue(RemindMeFileName)
    
            Do Until Msgs.Count = 0
                NextMsg = CStr(Msgs(1))
    
                ProcessMessage NextMsg
    
                Msgs.Remove 1
            Loop
            Set Msgs = Nothing
        End If
    End If
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "servicestart"
            RunProcess AppPath & UtilityFileName, "/start", vbHide, True
        Case "servicestop"
            RunProcess AppPath & UtilityFileName, "/stop", vbHide, True
        Case "add"
            NewOperation
        Case "edit"
            EditOperation
        Case "delete"
            DeleteOperation
        Case "runselected"
            RunSelected
        Case "stop"
            StopOperation
        Case "enabled"
            ChangeEnabled
    End Select
    
End Sub
Private Sub RefreshOperations()
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim node
    Dim sch As String

    Dim oldSelIndex As Long
    If Not lstObjects.SelectedItem Is Nothing Then
        oldSelIndex = lstObjects.SelectedItem.Index
    End If

    lstObjects.ListItems.Clear

    DBConn.rsQuery rs, "SELECT * FROM Operations ORDER BY Name;"

    Do While Not rs.EOF
        Set node = lstObjects.ListItems.Add(, , rs("Name"), , "item")
        node.SubItems(1) = rs("Enabled")
        node.SubItems(2) = ScriptToDisplay(rs("ProcName"))

        Select Case rs("ScheduleType")
            Case 0
                node.SubItems(3) = "Manual Schedule"
                sch = "N/A"
            Case 1
                node.SubItems(3) = "Increment Schedule"
                Select Case rs("IncrementType")
                    Case 0
                        If rs("IncrementInterval") = 1 Then
                            sch = "Execute every minute starting at " & rs("ExecuteDate") & " " & rs("ExecuteTime")
                        Else
                            sch = "Execute every " & rs("IncrementInterval") & " minutes starting at " & rs("ExecuteDate") & " " & rs("ExecuteTime")
                        End If
                    Case 1
                        If rs("IncrementInterval") = 1 Then
                            sch = "Execute every hour starting at " & rs("ExecuteDate") & " " & rs("ExecuteTime")
                        Else
                            sch = "Execute every " & rs("IncrementInterval") & " hours starting at " & rs("ExecuteDate") & " " & rs("ExecuteTime")
                        End If
                    Case 2
                        If rs("IncrementInterval") = 1 Then
                            sch = "Execute every day starting at " & rs("ExecuteDate") & " " & rs("ExecuteTime")
                        Else
                            sch = "Execute every " & rs("IncrementInterval") & " days starting at " & rs("ExecuteDate") & " " & rs("ExecuteTime")
                        End If
                End Select
                node.SubItems(4) = sch
            Case 2
                node.SubItems(3) = "Set Schedule"
                node.SubItems(4) = "Execute at " & rs("ExecuteDate") & " " & rs("ExecuteTime")
        End Select

        node.Tag = rs("ID")

        rs.MoveNext
    Loop

    If oldSelIndex > lstObjects.ListItems.Count Then
        oldSelIndex = lstObjects.ListItems.Count
    End If
    If oldSelIndex > 0 Then
        lstObjects.ListItems(oldSelIndex).Selected = True
    End If

    If Not rs2.State = 0 Then rs2.Close
    Set rs2 = Nothing

    If Not rs.State = 0 Then rs.Close
    Set rs = Nothing
End Sub
Private Sub NewOperation()
    frmWizard.Caption = "New Operation"
    frmWizard.usrWizOperation1.Description = "New Operation"
    frmWizard.EditID = 0
    frmWizard.Show 1
    RefreshOperations
End Sub
Private Sub EditOperation()
    On Error GoTo catch
    
    If Not lstObjects.SelectedItem Is Nothing Then
        
        Dim rs As New ADODB.Recordset
        
        frmWizard.Caption = "Edit Operation"
        frmWizard.EditID = lstObjects.SelectedItem.Tag
        
        DBConn.rsQuery rs, "SELECT * FROM Operations WHERE ID=" & lstObjects.SelectedItem.Tag & ";"
        
        With frmWizard.usrWizOperation1
            .Description = rs("Name")
            .Procedure = rs("ProcName")
        End With
        With frmWizard.usrWizTimer1
            .Enabled = rs("Enabled")
            .ScheduleType = rs("ScheduleType")
            .ExecuteDate = rs("ExecuteDate")
            .ExecuteTime = rs("ExecuteTime")
            .IncrementType = rs("IncrementType")
            .IncrementInterval = rs("IncrementInterval")
        End With
        
        DBConn.rsQuery rs, "SELECT * FROM OperationParams WHERE ParentID=" & lstObjects.SelectedItem.Tag & " ORDER BY ParamNum;"
        
        Dim Enumerator As String
        Dim EnumValue As clsEnumValue
        With frmWizard.usrWizOperation1.Parameters
            If Not rs.EOF And Not rs.BOF Then
                rs.MoveFirst
                Do
                    Select Case rs("ParamType")
                        Case 1
                            Select Case CStr(rs("ParamValue"))
                                Case "True", "-1", "1", "Yes"
                                    .ListItems(CInt(rs("ParamNum"))).SubItems(1) = "Yes (Enabled)"
                                Case Else
                                    .ListItems(CInt(rs("ParamNum"))).SubItems(1) = "No (Disabled)"
                            End Select
                        Case 2, 3
                            .ListItems(CInt(rs("ParamNum"))).SubItems(1) = rs("ParamValue")
                        Case 4
                            If Not rs("ParamValue") = "" Then
                                Enumerator = Procedures(frmWizard.usrWizOperation1.Procedure).Parameters(DisplayToScript(.ListItems(CInt(rs("ParamNum"))).Text)).ParamType
                                For Each EnumValue In Enumerators(Enumerator).EnumValues
                                    If EnumValue.EnumValue = rs("ParamValue") Then
                                        .ListItems(CInt(rs("ParamNum"))).SubItems(1) = ScriptToDisplay(EnumValue.EnumName)
                                    End If
                                Next
                            End If
                    End Select
                    rs.MoveNext
                Loop Until rs.EOF Or rs.BOF
            End If
        End With
                
        If Not rs.State = 0 Then rs.Close
        Set rs = Nothing
    
        frmWizard.Show 1
    
        RefreshOperations
    Else
        MsgBox "Please select an operation to edit.", vbInformation, "Edit Operation"
    End If
    
    Exit Sub
catch:
    MsgBox "There was an error loading the operation.", vbExclamation, "Edit Operation"
    Err.Clear
End Sub

Private Function DeleteOperation()
    If Not lstObjects.SelectedItem Is Nothing Then
        Dim rs As New ADODB.Recordset
        
        DBConn.rsQuery rs, "DELETE * FROM OperationParams WHERE ParentID=" & lstObjects.SelectedItem.Tag & ";"
        
        DBConn.rsQuery rs, "DELETE * FROM Operations WHERE ID=" & lstObjects.SelectedItem.Tag & ";"
    
        If Not rs.State = 0 Then rs.Close
        Set rs = Nothing
        
        SendMessage "removeoperation:" & lstObjects.SelectedItem.Tag
        
        RefreshOperations
        
    Else
        MsgBox "Please select an operation to delete.", vbInformation, "Delete Operation"
    End If
End Function

Public Function StopOperation()
    
    SendMessage "stopoperations"

End Function

Public Function RunSelected()
    'mnuUpdate_Click
    
    If Not lstObjects.SelectedItem Is Nothing Then
        SendMessage "startoperation:" & lstObjects.SelectedItem.Tag
    Else
        MsgBox "Please select an operation to run.", vbInformation, "Run Operation"
    End If
End Function
Public Function ChangeEnabled()
    If Not lstObjects.SelectedItem Is Nothing Then
        Dim rs As New ADODB.Recordset
        
        If Not CBool(lstObjects.SelectedItem.SubItems(1)) Then
            DBConn.rsQuery rs, "UPDATE Operations SET Enabled=True WHERE ID = " & lstObjects.SelectedItem.Tag & ";"
        Else
            DBConn.rsQuery rs, "UPDATE Operations SET Enabled=False WHERE ID = " & lstObjects.SelectedItem.Tag & ";"
        End If
        
        If Not rs.State = 0 Then rs.Close
        Set rs = Nothing
        
        SendMessage "updateoperation:" & lstObjects.SelectedItem.Tag
        
        RefreshOperations
        
    Else
        MsgBox "Please select an operation to change its status.", vbInformation, "Enable/Disable Operation"
    End If
End Function
Public Function SendMessage(ByVal Message As String)
    If (ProcessRunning(ServiceFileName) > 0) Or Not IsCompiled Then
        Dim rs As New ADODB.Recordset
        DBConn.rsQuery rs, "INSERT INTO MessageQueue (MessageTo, MessageText) VALUES ('" & ServiceFileName & "','" & Message & "');"
        If Not rs.State = 0 Then rs.Close
        Set rs = Nothing
    End If
End Function

Private Sub txtDebug_Change()
    If txtDebug.Tag Then
        SetDebugPreview
    End If
End Sub

Private Sub txtDebug_Click()
    If txtDebug.ForeColor = &H80000011 Then
        txtDebug.SelStart = 0
        txtDebug.SelLength = 0
    End If
End Sub

Private Sub txtDebug_KeyPress(KeyAscii As Integer)

    SetDebugPreview
    
    If Chr(KeyAscii) = "!" Then
        txtDebug.Text = "! "
        KeyAscii = 0
        txtDebug.SelStart = 3
    ElseIf Chr(KeyAscii) = "?" Then
        txtDebug.Text = "? "
        KeyAscii = 0
        txtDebug.SelStart = 3
        
    ElseIf KeyAscii = 13 Then
    
        If UpdateScriptCode Then
            On Error Resume Next
            
            If Left(txtDebug.Text, 1) = "!" Then
                ScriptControl1.ExecuteStatement Mid(txtDebug.Text, 2)
               
            ElseIf Left(txtDebug.Text, 1) = "?" Then
                txtDebug.Text = ScriptControl1.Eval(Mid(txtDebug.Text, 2))
            Else
                MsgBox "Type ! and a line of code to execute a statement, or type ? and a line of code to evaluate an expression.", AppName
            End If
        
            If Not (ScriptControl1.Error.Number = 0) Then
                DisplayError txtDebug.Text, ScriptControl1.Error.Number, 0, ScriptControl1.Error.Column, ScriptControl1.Error.Source, ScriptControl1.Error.Description
            End If
            Err.Clear
        
        End If

    End If
    
End Sub

Public Sub SetDebugPreview()
    If txtDebug.Text = "" Then
        txtDebug.Tag = False
        txtDebug.ForeColor = &H80000011
        txtDebug.Text = "Type ! or ? and a statement to debug."
        txtDebug.Tag = True
    ElseIf txtDebug.Text = "Type ! or ? and a statement to debug." Then
        txtDebug.Tag = False
        txtDebug.ForeColor = &H80000008
        txtDebug.Text = ""
        txtDebug.Tag = True
    End If
End Sub
