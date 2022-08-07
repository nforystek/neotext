VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMainIDE 
   BackColor       =   &H8000000C&
   Caption         =   "Visual Max-FTP IDE"
   ClientHeight    =   9810
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   17430
   Icon            =   "frmMainIDE.frx":0000
   LinkTopic       =   "MDIForm1"
   Visible         =   0   'False
   Begin VB.PictureBox picToolBar 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   810
      Left            =   0
      ScaleHeight     =   810
      ScaleWidth      =   17430
      TabIndex        =   1
      Top             =   0
      Width           =   17424
      Begin MSComctlLib.ImageList ToolBarOver 
         Left            =   4965
         Top             =   180
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
               Picture         =   "frmMainIDE.frx":08CA
               Key             =   "schedule"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ToolBarOut 
         Left            =   5610
         Top             =   180
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
               Picture         =   "frmMainIDE.frx":0D1C
               Key             =   "schedule"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TabStrip WindowTabs 
         Height          =   405
         Left            =   3735
         TabIndex        =   2
         Top             =   345
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   714
         MultiRow        =   -1  'True
         Style           =   2
         TabFixedHeight  =   300
         Separators      =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar ScriptControls 
         Height          =   450
         Left            =   1305
         TabIndex        =   0
         Top             =   270
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   794
         ButtonWidth     =   820
         ButtonHeight    =   794
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         Tag             =   "shadow"
         X1              =   180
         X2              =   180
         Y1              =   90
         Y2              =   750
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         Index           =   3
         Tag             =   "highlight"
         X1              =   315
         X2              =   315
         Y1              =   90
         Y2              =   705
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000014&
         Index           =   3
         Tag             =   "highlight"
         X1              =   525
         X2              =   1665
         Y1              =   210
         Y2              =   210
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000010&
         Index           =   3
         Tag             =   "shadow"
         X1              =   480
         X2              =   1770
         Y1              =   75
         Y2              =   75
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000014&
         Index           =   3
         Tag             =   "highlight"
         X1              =   2085
         X2              =   2085
         Y1              =   120
         Y2              =   690
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         Index           =   3
         Tag             =   "shadow"
         X1              =   1875
         X2              =   1875
         Y1              =   165
         Y2              =   660
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000010&
         Index           =   3
         Tag             =   "shadow"
         X1              =   540
         X2              =   1695
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line8 
         BorderColor     =   &H80000014&
         Index           =   3
         Tag             =   "highlight"
         X1              =   600
         X2              =   1740
         Y1              =   735
         Y2              =   735
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5040
      Top             =   900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList CommonIcons 
      Left            =   5610
      Top             =   870
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainIDE.frx":116E
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainIDE.frx":1A48
            Key             =   "Module"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainIDE.frx":1DDD
            Key             =   "Object"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainIDE.frx":2194
            Key             =   "FolderClose"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainIDE.frx":252B
            Key             =   "JScript"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainIDE.frx":28DF
            Key             =   "VBScript"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainIDE.frx":2C93
            Key             =   "FolderOpen"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuProject 
      Caption         =   "&Project"
      Begin VB.Menu mnuNew 
         Caption         =   "&New Project"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Project.."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close Project"
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveProject 
         Caption         =   "&Save Project"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveProjectAs 
         Caption         =   "Save Project &As.."
      End
      Begin VB.Menu mnuDash232 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRun 
         Caption         =   "&Run Project"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuStop 
         Caption         =   "S&top Project"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuDebugMap 
         Caption         =   "Project &Map"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDash287 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecent 
         Caption         =   "(No Recent Files)"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuRecent 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecent 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecent 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDash32837 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewInstance 
         Caption         =   "New IDE &Instance"
      End
      Begin VB.Menu mnuDash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuItem 
      Caption         =   "&Item"
      Begin VB.Menu mnuNewItem 
         Caption         =   "&New Item.."
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuDash2873 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenItem 
         Caption         =   "&Open Item"
      End
      Begin VB.Menu mnuExamine 
         Caption         =   "&Examine Item"
      End
      Begin VB.Menu mnuDash4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteItem 
         Caption         =   "&Delete Item"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "&Rename Item"
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
      Begin VB.Menu mnuDash182 
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
      Begin VB.Menu mnuDash92 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "&Select All"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuDash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComment 
         Caption         =   "Co&mment"
      End
      Begin VB.Menu mnuUncomment 
         Caption         =   "Unc&omment"
      End
      Begin VB.Menu mnuDash873 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "&Find or Replace.."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuWinProject 
         Caption         =   "&Project Window"
      End
      Begin VB.Menu mnuWinDebug 
         Caption         =   "&Debug Window"
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuExperiments 
         Caption         =   "&Experiments"
         Begin VB.Menu mnuExperiment 
            Caption         =   "Experiment"
            Index           =   0
         End
      End
      Begin VB.Menu mnuExamples 
         Caption         =   "&JScript Examples"
         Begin VB.Menu mnuExample 
            Caption         =   "Example"
            Index           =   0
         End
      End
      Begin VB.Menu mnuExamples2 
         Caption         =   "&VBScript Examples"
         Begin VB.Menu mnuExample2 
            Caption         =   "Example"
            Index           =   0
         End
      End
      Begin VB.Menu mnuNoProjects 
         Caption         =   "(No Installed Projects)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDash45 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContents 
         Caption         =   "&Documentation..."
      End
      Begin VB.Menu mnuNeoWeb 
         Caption         =   "&Neotext.org..."
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMainIDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public Sub ProjectSet(ByVal pChanged As Boolean, Optional ByVal pFileName As String = "*")
    If Not (pFileName = "*") Then
        Project.FileName = pFileName
    End If
    Project.Changed = pChanged
    If Not pChanged Then
        Dim frm As Form
        For Each frm In Forms
             If TypeName(frm) = "frmScriptPage" Then
                frm.CodeEdit1.UndoDirty = False
             End If
         Next
    End If
    RefreshCaption
End Sub

Private Sub LoadGUI()

    Set ScriptControls.ImageList = ToolBarOut
    Set ScriptControls.DisabledImageList = ToolBarOut
    Set ScriptControls.HotImageList = ToolBarOver
    
    Dim btnX As Button
    
    Set btnX = ScriptControls.Buttons.Add(1, "script_newproject", , 0, "script_newprojectout")
    Set btnX = ScriptControls.Buttons.Add(2, "script_openproject", , 0, "script_openprojectout")
    Set btnX = ScriptControls.Buttons.Add(3, "script_saveproject", , 0, "script_saveprojectout")
    Set btnX = ScriptControls.Buttons.Add(4, , , 3)
    If GetCollectSkinValue("toolbarbutton_spacer") = "none" Then btnX.Visible = False

    Set btnX = ScriptControls.Buttons.Add(5, "script_addfile", , 0, "script_addfileout")
    Set btnX = ScriptControls.Buttons.Add(6, "script_removefile", , 0, "script_removefileout")

    Set btnX = ScriptControls.Buttons.Add(7, "script_undo", , 0, "script_undoout")
    Set btnX = ScriptControls.Buttons.Add(8, "script_redo", , 0, "script_redoout")
    Set btnX = ScriptControls.Buttons.Add(9, , , 3)
    If GetCollectSkinValue("toolbarbutton_spacer") = "none" Then btnX.Visible = False

    Set btnX = ScriptControls.Buttons.Add(10, "script_cut", , 0, "script_cutout")
    Set btnX = ScriptControls.Buttons.Add(11, "script_copy", , 0, "script_copyout")
    Set btnX = ScriptControls.Buttons.Add(12, "script_paste", , 0, "script_pasteout")
    Set btnX = ScriptControls.Buttons.Add(13, "script_find", , 0, "script_findout")
    Set btnX = ScriptControls.Buttons.Add(14, , , 3)
    If GetCollectSkinValue("toolbarbutton_spacer") = "none" Then btnX.Visible = False

    Set btnX = ScriptControls.Buttons.Add(15, "script_stop", , 0, "script_stopout")
    Set btnX = ScriptControls.Buttons.Add(16, "script_run", , 0, "script_runout")
    Set btnX = ScriptControls.Buttons.Add(17, , , 3)
    If GetCollectSkinValue("toolbarbutton_spacer") = "none" Then btnX.Visible = False
    
    WindowTabs.Separators = Not (GetCollectSkinValue("toolbarbutton_spacer") = "none")
    SetTooltip
    
    mnuDebugMap.Visible = IsDebugger
    
End Sub
Private Sub UnloadGUI()
    
    ScriptControls.Buttons.Clear
    Set ScriptControls.ImageList = Nothing
    Set ScriptControls.DisabledImageList = Nothing
    Set ScriptControls.HotImageList = Nothing

    Unload frmDebug
    Unload frmProjectExplorer
    
    Dim cnt As Long
    cnt = 0
    Do While cnt <= (Forms.Count - 1)
        If Not (TypeName(Forms(cnt)) = "frmMainIDE") Then
            Unload Forms(cnt)
        Else
            cnt = cnt + 1
        End If
    Loop
    
End Sub

Private Sub MDIForm_Load()

#If VBIDE Then
    mnuDebugMap.Visible = True
#End If

    SetIcecue Line1(3), "icecue_shadow"
    SetIcecue Line5(3), "icecue_shadow"
    SetIcecue Line4(3), "icecue_shadow"
    SetIcecue Line7(3), "icecue_shadow"
    
    SetIcecue Line2(3), "icecue_hilite"
    SetIcecue Line6(3), "icecue_hilite"
    SetIcecue Line3(3), "icecue_hilite"
    SetIcecue Line8(3), "icecue_hilite"
    
    LoadScriptGraphics
    
    LoadGUI
    
    RefreshHelpMenu
    
    If dbSettings.GetScriptingSetting("wState") = -1 Then
        dbSettings.SetScriptingSetting "wState", 0
        dbSettings.SetScriptingSetting "wLeft", ((Screen.Width / 2) - (Me.Width / 2))
        dbSettings.SetScriptingSetting "wTop", ((Screen.Height / 2) - (Me.Height / 2))
        dbSettings.SetScriptingSetting "wWidth", Me.Width
        dbSettings.SetScriptingSetting "wHeight", Me.Height
    End If
    
    Me.Move dbSettings.GetScriptingSetting("wLeft"), dbSettings.GetScriptingSetting("wTop"), dbSettings.GetScriptingSetting("wWidth"), dbSettings.GetScriptingSetting("wHeight")
    Me.WindowState = dbSettings.GetScriptingSetting("wState")

    EnableProjectWindow dbSettings.GetScriptingSetting("pVisible")
    EnableDebugWindow dbSettings.GetScriptingSetting("dVisible")

    ProjectSet False, ""
    
    CommonDialog1.InitDir = ProjectFolder
    
    RefreshWindowMenu
    
    RefreshRecent
    
End Sub

Public Sub RefreshRecent()

    mnuRecent(0).Enabled = Not (dbSettings.GetScriptingSetting("Recent0") & "" = "")
    mnuRecent(0).Caption = IIf(mnuRecent(0).Enabled, GetFileName(dbSettings.GetScriptingSetting("Recent0") & ""), "(No Recent Files)")
    mnuRecent(0).Tag = dbSettings.GetScriptingSetting("Recent0") & ""

    mnuRecent(1).Visible = Not (dbSettings.GetScriptingSetting("Recent1") & "" = "")
    mnuRecent(1).Caption = IIf(mnuRecent(1).Visible, GetFileName(dbSettings.GetScriptingSetting("Recent1") & ""), "")
    mnuRecent(1).Tag = dbSettings.GetScriptingSetting("Recent1") & ""
    
    mnuRecent(2).Visible = Not (dbSettings.GetScriptingSetting("Recent2") & "" = "")
    mnuRecent(2).Caption = IIf(mnuRecent(2).Visible, GetFileName(dbSettings.GetScriptingSetting("Recent2") & ""), "")
    mnuRecent(2).Tag = dbSettings.GetScriptingSetting("Recent2") & ""
    
    mnuRecent(3).Visible = Not (dbSettings.GetScriptingSetting("Recent3") & "" = "")
    mnuRecent(3).Caption = IIf(mnuRecent(3).Visible, GetFileName(dbSettings.GetScriptingSetting("Recent3") & ""), "")
    mnuRecent(3).Tag = dbSettings.GetScriptingSetting("Recent3") & ""

End Sub

Public Sub AddRecent(ByVal fname As String)

    If PathExists(fname, True) And Not LCase(GetFilePath(fname)) & "\" = LCase(AppPath & TemplatesFolder) Then
    
        Dim mnu
        For Each mnu In mnuRecent
            If LCase(Trim(mnu.Tag)) = LCase(Trim(fname)) Then
                Exit Sub
            End If
        Next
       
        dbSettings.SetScriptingSetting "Recent3", dbSettings.GetScriptingSetting("Recent2") & ""
        dbSettings.SetScriptingSetting "Recent2", dbSettings.GetScriptingSetting("Recent1") & ""
        dbSettings.SetScriptingSetting "Recent1", dbSettings.GetScriptingSetting("Recent0") & ""
        dbSettings.SetScriptingSetting "Recent0", fname
        
    End If
    
    RefreshRecent
End Sub

Public Sub ShowForm()
    Me.Show
    
    DoEvents
    If CBool(dbSettings.GetScriptingSetting("AboutWindow")) Then frmMainIDE.StartPage
    If CBool(dbSettings.GetScriptingSetting("ContentsWindow")) Then frmMainIDE.ContentsPage
                
    EnableProjectWindow dbSettings.GetScriptingSetting("pVisible")
    EnableDebugWindow dbSettings.GetScriptingSetting("dVisible")
    
    Me.SetFocus
End Sub
Public Sub RefreshCaption()
    If Project.Loaded Then
        Dim tmp As String
        tmp = " (" & GetFileName(Project.FileName) & ")" & IIf(Project.Changed, "*", "")
        Me.Caption = MaxIDEFormCaption & tmp
    Else
        Me.Caption = MaxIDEFormCaption
    End If
End Sub
Public Function RefreshHelpMenu()
    
    Dim mnu
    Dim cnt As Long
    Dim fso As New Scripting.FileSystemObject
    Dim f1 As Folder
    Dim f2 As File

    For Each mnu In mnuExample
        If mnu.Index <> 0 Then Unload mnu
    Next
    If PathExists(AppPath & ProjectFolder & "Examples\JScript") Then
        Set f1 = fso.GetFolder(AppPath & ProjectFolder & "Examples\JScript")
        cnt = 0
        For Each f2 In f1.Files
            If InStr(f2.name, MaxProjectExt) > 0 Then
                If Not cnt = 0 Then Load mnuExample(cnt)
                mnuExample(cnt).Visible = True
                mnuExample(cnt).Caption = Replace(f2.name, MaxProjectExt, "")
            
                cnt = cnt + 1
            End If
        Next
        mnuExamples.Visible = (cnt > 0)
    Else
        mnuExamples.Visible = False
    End If
    
    For Each mnu In mnuExample2
        If mnu.Index <> 0 Then Unload mnu
    Next
    If PathExists(AppPath & ProjectFolder & "Examples\VBScript") Then
        Set f1 = fso.GetFolder(AppPath & ProjectFolder & "Examples\VBScript")
        cnt = 0
        For Each f2 In f1.Files
            If InStr(f2.name, MaxProjectExt) > 0 Then
                If Not cnt = 0 Then Load mnuExample2(cnt)
                mnuExample2(cnt).Visible = True
                mnuExample2(cnt).Caption = Replace(f2.name, MaxProjectExt, "")
            
                cnt = cnt + 1
            End If
        Next
        mnuExamples2.Visible = (cnt > 0)
    Else
        mnuExamples2.Visible = False
    End If
    
    For Each mnu In mnuExperiment
        If mnu.Index <> 0 Then Unload mnu
    Next
    If PathExists(AppPath & ProjectFolder & "Experiment") Then
        Set f1 = fso.GetFolder(AppPath & ProjectFolder & "Experiment")
        cnt = 0
        For Each f2 In f1.Files
            If InStr(f2.name, MaxProjectExt) > 0 Then
                If Not cnt = 0 Then Load mnuExperiment(cnt)
                mnuExperiment(cnt).Visible = True
                mnuExperiment(cnt).Caption = Replace(f2.name, MaxProjectExt, "")
            
                cnt = cnt + 1
            End If
        Next
        mnuExperiments.Visible = (cnt > 0)
    Else
        mnuExperiments.Visible = False
    End If

    mnuNoProjects.Visible = Not (mnuExamples2.Visible Or mnuExamples.Visible Or mnuExperiments.Visible)
End Function

Public Function EnableProjectWindow(ByVal Enabled As Boolean)
    frmProjectExplorer.ctlDragger1.SetVisible Enabled
    frmProjectExplorer.Visible = Enabled And frmMainIDE.Visible
    mnuWinProject.Checked = Enabled
End Function

Public Function EnableDebugWindow(ByVal Enabled As Boolean)
    frmDebug.ctlDragger1.SetVisible Enabled
    frmDebug.Visible = Enabled And frmMainIDE.Visible
    mnuWinDebug.Checked = Enabled
End Function

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = Not CloseProject
    
    If Not Cancel And frmMainIDE.Visible Then
    
        dbSettings.SetScriptingSetting "AboutWindow", frmProjectExplorer.PageIsOpen("About Page")
        dbSettings.SetScriptingSetting "ContentsWindow", frmProjectExplorer.PageIsOpen("Help Contents")
        
    End If
    
End Sub

Private Sub MDIForm_Resize()
    On Error Resume Next
    Dim xPixel As Integer
    Dim yPixel As Integer
    xPixel = (Screen.TwipsPerPixelX)
    yPixel = (Screen.TwipsPerPixelY)
    
    Dim wLeft As Long
    Dim btn As Button
    For Each btn In ScriptControls.Buttons
        wLeft = wLeft + btn.Width
    Next
    
    With picToolBar
        .Align = 1
        .Top = 0
        .Height = ((GetSkinDimension("toolbarbutton_height") + 6) * Screen.TwipsPerPixelY) + (yPixel * 4)
    End With
    
    With ScriptControls
        .Left = (yPixel * 2)
        .Top = (yPixel * 2)
        .Width = wLeft
    End With
    
    wLeft = wLeft + (xPixel * 5)
    With WindowTabs
        .Height = (GetSkinDimension("toolbarbutton_height") * Screen.TwipsPerPixelY)
        .Left = wLeft
        .Top = (picToolBar.Height / 2) - (WindowTabs.Height / 2)
        .Width = picToolBar.ScaleWidth - (wLeft + (xPixel * 2))
    End With
    
    Err.Clear
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If frmMainIDE.Visible Then
    
        dbSettings.SetScriptingSetting "wState", Me.WindowState
        
        If Me.WindowState = 0 Then
            dbSettings.SetScriptingSetting "wTop", Me.Top
            dbSettings.SetScriptingSetting "wLeft", Me.Left
            dbSettings.SetScriptingSetting "wHeight", Me.Height
            dbSettings.SetScriptingSetting "wWidth", Me.Width
        
        End If
        
        dbSettings.SetScriptingSetting "dDocked", frmDebug.ctlDragger1.Docked And frmDebug.Visible
        dbSettings.SetScriptingSetting "dVisible", frmDebug.Visible
        
        dbSettings.SetScriptingSetting "dDockedHeight", frmDebug.ctlDragger1.DockedHeight
        dbSettings.SetScriptingSetting "dDockedWidth", frmDebug.ctlDragger1.DockedWidth
        dbSettings.SetScriptingSetting "dFloatTop", frmDebug.ctlDragger1.FloatingTop
        dbSettings.SetScriptingSetting "dFloatLeft", frmDebug.ctlDragger1.FloatingLeft
        dbSettings.SetScriptingSetting "dFloatHeight", frmDebug.ctlDragger1.FloatingHeight
        dbSettings.SetScriptingSetting "dFloatWidth", frmDebug.ctlDragger1.FloatingWidth
        
        dbSettings.SetScriptingSetting "pDocked", frmProjectExplorer.ctlDragger1.Docked And frmProjectExplorer.Visible
        dbSettings.SetScriptingSetting "pVisible", frmProjectExplorer.Visible
        
        dbSettings.SetScriptingSetting "pDockedHeight", frmProjectExplorer.ctlDragger1.DockedHeight
        dbSettings.SetScriptingSetting "pDockedWidth", frmProjectExplorer.ctlDragger1.DockedWidth
        dbSettings.SetScriptingSetting "pFloatTop", frmProjectExplorer.ctlDragger1.FloatingTop
        dbSettings.SetScriptingSetting "pFloatLeft", frmProjectExplorer.ctlDragger1.FloatingLeft
        dbSettings.SetScriptingSetting "pFloatHeight", frmProjectExplorer.ctlDragger1.FloatingHeight
        dbSettings.SetScriptingSetting "pFloatWidth", frmProjectExplorer.ctlDragger1.FloatingWidth
        
    End If
       
    Project.StopProject
    Project.ResetProject
    
    frmMainIDE.CloseProject

    Set MaxEvents = Nothing
    Set Project = Nothing
    
    UnloadGUI
  
    CleanUser True
    
    Cancel = False
    
    'If Not IsDebugger Or IsCompiled Then
        'TerminateProcess GetCurrentProcessId, 0
    'Else
    '    End
    'End If
           
End Sub

Private Function DoExamine()
    
    If TypeName(frmProjectExplorer.SelectedNode.Tag) = "clsItem" Then
                
        Dim nName As String
        Dim nPage As frmIEPage
        nName = "Examine:" & frmProjectExplorer.SelectedNode.Tag.ItemName
        Set nPage = frmProjectExplorer.GetInternetPage(nName)
        If nPage Is Nothing Then Set nPage = New frmIEPage
        If PathExists(AppPath & "Help\" & frmProjectExplorer.SelectedNode.Tag.ItemClass & ".htm") Then
            nPage.ShowForm nName, AppPath & "Help\" & frmProjectExplorer.SelectedNode.Tag.ItemClass & ".htm"
        Else
            nPage.ShowForm nName, NeoTextWebSite & "/ipub/help/max-ftp/" & frmProjectExplorer.SelectedNode.Tag.ItemClass & ".htm"
        End If
        Set nPage = Nothing
    End If
   
End Function

Public Sub StartPage()
    Dim nPage As frmIEPage
    Set nPage = frmProjectExplorer.GetInternetPage("About Page")
    If nPage Is Nothing Then Set nPage = New frmIEPage
    If PathExists(AppPath & HelpFolder & "ScriptingAbout.htm") Then
        nPage.ShowForm "About Page", AppPath & HelpFolder & "ScriptingAbout.htm"
    Else
        nPage.ShowForm "About Page", NeoTextWebSite & "/ipub/help/max-ftp/" & "ScriptingAbout.htm"
    End If
    Set nPage = Nothing
End Sub
Public Sub ContentsPage()
    Dim nPage As frmIEPage
    Set nPage = frmProjectExplorer.GetInternetPage("Help Contents")
    If nPage Is Nothing Then Set nPage = New frmIEPage
    If IsDocumentationInstalled Then
        nPage.ShowForm "Help Contents", AppPath & HelpFolder & "index.htm"
    Else
        nPage.ShowForm "Help Contents", NeoTextWebSite & "/ipub/help/max-ftp"
    End If
    Set nPage = Nothing
End Sub
Private Sub mnuContents_Click()
    ContentsPage
End Sub

Private Sub mnuAbout_Click()
    StartPage
End Sub

Private Sub mnuClose_Click()
    CloseProject
End Sub

Private Sub mnuComment_Click()
    Me.ActiveForm.CodeEdit1.Comment True
End Sub

Private Sub mnuDebugMap_Click()
    Dim pCompiler As New frmProjectCompiler

    Dim nName As String
    Dim nPage As frmScriptPage
    nName = GetFileTitle(Project.FileName)
    Set nPage = frmProjectExplorer.GetScriptPage(nName)
    If nPage Is Nothing Then
    
        Set nPage = New frmScriptPage
        
    End If
    
    Dim nItem As New clsItem
    nItem.ItemName = GetFileTitle(Project.FileName)
    nItem.ItemSource = pCompiler.MapProject(Project)

    If Not nItem.ItemSource = "" Then
        Set nPage.Item = nItem
        nPage.Locked = True
        
        If frmMainIDE.Visible Then
            nPage.ZOrder 0
        End If
    End If
    
    Unload pCompiler
End Sub

Private Sub mnuExamine_Click()
    DoExamine
End Sub

Private Sub mnuExample_Click(Index As Integer)
    If PathExists(AppPath & ProjectFolder & "Examples\JScript\" & mnuExample(Index).Caption & MaxProjectExt, True) Then
        OpenProject AppPath & ProjectFolder & "Examples\JScript\" & mnuExample(Index).Caption & MaxProjectExt
    End If
End Sub

Private Sub mnuExample2_Click(Index As Integer)
    If PathExists(AppPath & ProjectFolder & "Examples\VBScript\" & mnuExample2(Index).Caption & MaxProjectExt, True) Then
        OpenProject AppPath & ProjectFolder & "Examples\VBScript\" & mnuExample2(Index).Caption & MaxProjectExt
    End If
End Sub

Private Sub mnuExperiment_Click(Index As Integer)
    If PathExists(AppPath & ProjectFolder & "Experiment\" & mnuExperiment(Index).Caption & MaxProjectExt, True) Then
        OpenProject AppPath & ProjectFolder & "Experiment\" & mnuExperiment(Index).Caption & MaxProjectExt
    End If
End Sub

Private Sub mnuFind_Click()
    frmFind.Show
End Sub

Private Sub mnuItem_Click()
    RefreshScriptMenu
End Sub

Private Sub mnuNewInstance_Click()
    RunProcess AppPath & MaxIDEFileName, , 1, False
End Sub

Private Sub mnuOpenItem_Click()
    frmProjectExplorer.OpenItem
End Sub

Private Sub mnuRecent_Click(Index As Integer)
    If PathExists(mnuRecent(Index).Tag, True) Then
        OpenProject mnuRecent(Index).Tag
    End If
End Sub


Private Sub mnuRename_Click()
    DoRename
End Sub

Public Function DoRename()
    frmProjectExplorer.BeginRename False
End Function
Private Sub mnuProject_Click()
    RefreshScriptMenu
End Sub

Private Sub mnuStop_Click()
    Project.StopProject
End Sub

Private Sub mnuUncomment_Click()
    Me.ActiveForm.CodeEdit1.Comment False
End Sub

Private Sub mnuUndo_Click()
    Me.ActiveForm.CodeEdit1.Undo
End Sub
Private Sub mnuRedo_Click()
    Me.ActiveForm.CodeEdit1.Redo
End Sub

Private Sub mnuDelete_Click()
'    Me.ActiveForm.CodeEdit1.Clear
    If Me.ActiveForm.CodeEdit1.SelLength = 0 Then
        Me.ActiveForm.CodeEdit1.SelLength = 1
    End If
    Me.ActiveForm.CodeEdit1.SelText = ""
End Sub
Private Sub mnuSelectAll_Click()
    Me.ActiveForm.CodeEdit1.SelStart = 0
    Me.ActiveForm.CodeEdit1.SelLength = Len(Me.ActiveForm.CodeEdit1.Text)
End Sub

Private Sub mnuCopy_Click()
    Clipboard.Clear
    Clipboard.SetText Me.ActiveForm.CodeEdit1.SelText, vbCFText
End Sub
Private Sub mnuCut_Click()
    Clipboard.Clear
    Clipboard.SetText Me.ActiveForm.CodeEdit1.SelText, vbCFText
    Me.ActiveForm.CodeEdit1.SelText = ""
End Sub
Private Sub mnuPaste_Click()
    Me.ActiveForm.CodeEdit1.SelText = Clipboard.GetText(vbCFText)
    'Me.ActiveForm.CodeEdit1.SelStart = Me.ActiveForm.CodeEdit1.SelStart + Len(Clipboard.GetText(vbCFText))
End Sub

Private Sub mnuEdit_Click()
    RefreshEditMenu
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuNeoWeb_Click()
    RunFile AppPath & "Neotext.org.url"
End Sub

Private Sub mnuNew_Click()
    Load frmDialog
    frmDialog.InitDialog 0
    If frmDialog.IsOk Then
    
        Select Case frmDialog.ItemClass
            Case "Project"
                OpenProject AppPath & TemplatesFolder & "New" & frmDialog.ItemName & "Project" & MaxProjectExt, frmDialog.ItemName
                ProjectSet False, "New" & frmDialog.ItemName & "Project" & MaxProjectExt
            Case "Script"
                OpenProject AppPath & TemplatesFolder & "NewScriptTemplate" & MaxScriptExt
                ProjectSet False, "NewScriptTemplate" & MaxScriptExt
        End Select

    End If
    Unload frmDialog
End Sub

Private Sub mnuNewItem_Click()
    Load frmDialog
    frmDialog.InitDialog 1
    If frmDialog.IsOk Then
    
        Dim cItem As New clsItem
        cItem.LoadFromFile AppPath & TemplatesFolder & frmDialog.ItemClass & MaxScriptExt, Project.Language
        cItem.ItemPath = ""
        
        Dim cNode As MSComctlLib.Node
        Set cNode = frmProjectExplorer.AddItem(cItem)
        
        Project.Items.Add cItem, cItem.ItemName
        
        cNode.Parent.Expanded = True
        cNode.EnsureVisible
        cNode.Selected = True
        
        frmProjectExplorer.ShowScriptPage cItem.ItemName
        
        Set cNode = Nothing
        Set cItem = Nothing
        
        ProjectSet True
    
    End If
    Unload frmDialog
End Sub

Private Sub mnuOpen_Click()
    OpenProject
End Sub

Private Sub mnuDeleteItem_Click()
    If MsgBox("Are you sure you want to delete this item?", vbYesNo + vbQuestion, AppName) = vbYes Then
        With frmProjectExplorer

            Dim frm As Form
            For Each frm In Forms
                If TypeName(frm) = "frmScriptPage" Then

                    If frm.Item.ItemName = .SelectedNode.Tag.ItemName Then
                        Unload frm
                    End If

                End If
            Next
            
            Dim ItemName As String
            ItemName = .SelectedNode.Tag.ItemName

            .DeleteItem ItemName

            Project.Remove ItemName
            
            ProjectSet True
        
        End With
    End If
End Sub

Private Sub mnuRun_Click()
    
    Dim frm As Form
    For Each frm In Forms
        If TypeName(frm) = "frmScriptPage" Then
            If Not (frm.CodeEdit1.ErrorLine = 0) Then frm.CodeEdit1.ErrorLine = 0
'            If Not (frm.CodeEdit1.colorerror = frm.CodeEdit1.ConvertColor(GetCollectSkinValue("script_errorcolor"))) Then
'                frm.CodeEdit1.colorerror = frm.CodeEdit1.ConvertColor(GetCollectSkinValue("script_errorcolor"))
'            End If
        End If
    Next
    
    Dim Text As String
    Text = Project.RunProject
    
End Sub

Private Sub mnuSaveProject_Click()
    SaveProject False
End Sub

Private Sub mnuSaveProjectAs_Click()
    SaveProject True
End Sub

Public Function PromptSaveCancel() As Boolean
    If Project.Changed Then
        Select Case MsgBox("Do you want to save the current project?", vbYesNoCancel + vbQuestion, "Save Project?")
            Case vbYes
                PromptSaveCancel = Not (SaveProject(False))
            Case vbNo
                PromptSaveCancel = False
            Case vbCancel
                PromptSaveCancel = True
        End Select
    
    Else
        PromptSaveCancel = False
    End If
End Function

Public Function CloseProject() As Boolean
    If Not PromptSaveCancel Then
        
        Dim frm
        For Each frm In Forms
            If TypeName(frm) = "frmScriptPage" Then
                Unload frm
            End If
        Next
    
        Project.ResetProject

        frmProjectExplorer.RefreshProject
       
        ProjectSet False, ""
        
        CloseProject = True
    Else
        CloseProject = False
    End If
End Function

Public Function OpenProject(Optional ByVal pFileName As String = "", Optional ByVal pScript As String = "")
        
    If pFileName = "" Then
        On Error Resume Next
        
        CommonDialog1.CancelError = True
        CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNFileMustExist
        CommonDialog1.Filter = "Max Project Files (*" & MaxProjectExt & ", *" & MaxScriptExt & ")|*" & MaxProjectExt & ";*" & MaxScriptExt & "|All Files (*.*)|*.*"
        CommonDialog1.FilterIndex = 1
        
        CommonDialog1.ShowOpen
        pFileName = CommonDialog1.FileName
        
    End If
    
    If Err = 0 Then
            
         If CloseProject Then
            
            OpenProject = Project.LoadFromFile(pFileName, pScript)
            frmProjectExplorer.RefreshProject
            ProjectSet False
            AddRecent pFileName
            
        Else
            OpenProject = False
        End If
        
    Else
        Err.Clear
        OpenProject = False
    End If
    
End Function

Private Function SaveProject(ByVal SaveAs As Boolean) As Boolean
    On Error Resume Next
    
    If Not PathExists(Project.FileName, True) Or SaveAs Then
        CommonDialog1.CancelError = True
        CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
        If Project.IsTemplate Then
            CommonDialog1.Filter = "Max Script File (*" & MaxScriptExt & ")|*" & MaxScriptExt & "|All Files (*.*)|*.*"
        Else
            CommonDialog1.Filter = "Max Project File (*" & MaxProjectExt & ")|*" & MaxProjectExt & "|All Files (*.*)|*.*"
        End If
        CommonDialog1.FilterIndex = 1
        
        CommonDialog1.ShowSave
        
        If (Err = 0) Then Project.FileName = CommonDialog1.FileName
        
    End If
    
    If Err = 0 Then
        Screen.MousePointer = 11
        
        SaveProject = Project.SaveToFile(Project.FileName)
        
        If Err = 0 And SaveProject Then
            ProjectSet False
            
            DebugWinPrint "Save Complete " & Now
            frmProjectExplorer.RefreshName
            AddRecent Project.FileName
        
        Else
            DebugWinPrint "Save Failed: " & Err.Description
        End If
        
        Screen.MousePointer = 0
    Else
        Err.Clear
        SaveProject = False
    End If
End Function

Private Sub mnuWinDebug_Click()
    EnableDebugWindow Not mnuWinDebug.Checked
End Sub

Private Sub mnuWindow_Click(Index As Integer)
    ShowSelectedWindow Mid(mnuWindow(Index).Caption, 4)
End Sub

Private Sub mnuWinProject_Click()
    EnableProjectWindow Not mnuWinProject.Checked
End Sub

Private Sub picToolBar_Resize()
    On Error Resume Next
    
    Dim xPixel As Integer
    Dim yPixel As Integer
    xPixel = (Screen.TwipsPerPixelX)
    yPixel = (Screen.TwipsPerPixelY)
    
    Line1(3).x1 = 0
    Line1(3).X2 = 0
    Line1(3).y1 = yPixel
    Line1(3).Y2 = (picToolBar.Height - (2 * yPixel))
    
    Line2(3).x1 = xPixel
    Line2(3).X2 = xPixel
    Line2(3).y1 = yPixel
    Line2(3).Y2 = (picToolBar.Height - (2 * yPixel))
    
    Line3(3).x1 = picToolBar.Width - yPixel
    Line3(3).X2 = picToolBar.Width - yPixel
    Line3(3).y1 = 0
    Line3(3).Y2 = picToolBar.Height - yPixel
    
    Line4(3).x1 = picToolBar.Width - (xPixel * 2)
    Line4(3).X2 = picToolBar.Width - (xPixel * 2)
    Line4(3).y1 = yPixel
    Line4(3).Y2 = (picToolBar.Height - (2 * yPixel))
    
    Line5(3).x1 = 0
    Line5(3).X2 = picToolBar.Width
    Line5(3).y1 = 0
    Line5(3).Y2 = 0
    
    Line6(3).x1 = xPixel
    Line6(3).X2 = picToolBar.Width
    Line6(3).y1 = yPixel
    Line6(3).Y2 = yPixel
    
    Line7(3).x1 = xPixel
    Line7(3).X2 = picToolBar.Width - (xPixel * 2)
    Line7(3).y1 = picToolBar.Height - (yPixel * 2)
    Line7(3).Y2 = picToolBar.Height - (yPixel * 2)
    
    Line8(3).x1 = 0
    Line8(3).X2 = picToolBar.Width
    Line8(3).y1 = picToolBar.Height - yPixel
    Line8(3).Y2 = picToolBar.Height - yPixel
    
    Err.Clear

End Sub

Private Sub ScriptControls_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "script_newproject"
            mnuNew_Click
        Case "script_openproject"
            OpenProject
        Case "script_saveproject"
            SaveProject False
        Case "script_addfile"
            mnuNewItem_Click
        Case "script_removefile"
            mnuDeleteItem_Click
        Case "script_cut"
            mnuCut_Click
        Case "script_copy"
            mnuCopy_Click
        Case "script_paste"
            mnuPaste_Click
        Case "script_find"
            mnuFind_Click
        Case "script_stop"
            mnuStop_Click
        Case "script_run"
            mnuRun_Click
    End Select
End Sub

Private Sub ScriptControls_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RefreshToolbarMenu
End Sub

Private Sub WindowTabs_Click()
    If Not CancelTabClick Then
        If Not WindowTabs.SelectedItem Is Nothing Then
            ShowSelectedWindow WindowTabs.SelectedItem.Caption
        End If
    End If
End Sub

Public Sub SetTooltip()
    With Me
        If dbSettings.GetProfileSetting("ViewToolTips") Then
            .ScriptControls.Buttons(1).ToolTipText = "Create a new project."
            .ScriptControls.Buttons(2).ToolTipText = "Open an existing project."
            .ScriptControls.Buttons(3).ToolTipText = "Save the current project."
            
            .ScriptControls.Buttons(5).ToolTipText = "Adds a new file to the project."
            .ScriptControls.Buttons(6).ToolTipText = "Remove a file from the project."
            .ScriptControls.Buttons(7).ToolTipText = "Undo the last typed action."
            .ScriptControls.Buttons(8).ToolTipText = "Redo the last undo action."
            
            .ScriptControls.Buttons(10).ToolTipText = "Cut the selected code text."
            .ScriptControls.Buttons(11).ToolTipText = "Copy the selected code text."
            .ScriptControls.Buttons(12).ToolTipText = "Paste the clipboard contents."
            .ScriptControls.Buttons(13).ToolTipText = "Find or replace text in code."
            
            .ScriptControls.Buttons(15).ToolTipText = "Stop the project."
            .ScriptControls.Buttons(16).ToolTipText = "Run the project."
            
        Else
            .ScriptControls.Buttons(1).ToolTipText = ""
            .ScriptControls.Buttons(2).ToolTipText = ""
            .ScriptControls.Buttons(3).ToolTipText = ""
            
            .ScriptControls.Buttons(5).ToolTipText = ""
            .ScriptControls.Buttons(6).ToolTipText = ""
            .ScriptControls.Buttons(7).ToolTipText = ""
            .ScriptControls.Buttons(8).ToolTipText = ""
    
            .ScriptControls.Buttons(10).ToolTipText = ""
            .ScriptControls.Buttons(11).ToolTipText = ""
            .ScriptControls.Buttons(12).ToolTipText = ""
            .ScriptControls.Buttons(13).ToolTipText = ""
    
            .ScriptControls.Buttons(15).ToolTipText = ""
            .ScriptControls.Buttons(16).ToolTipText = ""
    
        End If
    End With
End Sub



