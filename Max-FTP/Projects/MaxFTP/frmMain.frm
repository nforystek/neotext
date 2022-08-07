VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C98B112F-745F-4542-B5B3-DDFADF1F6E2F}#1160.0#0"; "NTControls22.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Max-FTP Application Server"
   ClientHeight    =   2640
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3840
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DDEServer"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer timGlobal 
      Enabled         =   0   'False
      Interval        =   63
      Left            =   1344
      Top             =   1380
   End
   Begin VB.ListBox CopyItems 
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   180
      Width           =   1725
   End
   Begin MSComctlLib.ImageList imgOperations 
      Index           =   0
      Left            =   90
      Top             =   2010
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   "stopped"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C55
            Key             =   "paused"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0FEA
            Key             =   "running"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1374
            Key             =   "error"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgSchedule 
      Index           =   0
      Left            =   90
      Top             =   1380
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgDragDrop 
      Left            =   2505
      Top             =   1695
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1752
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A6C
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D86
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20A0
            Key             =   "abort"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgFiles 
      Left            =   1605
      Top             =   1725
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23BA
            Key             =   "folder"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgClient16 
      Index           =   0
      Left            =   1755
      Top             =   735
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgClient 
      Index           =   0
      Left            =   75
      Top             =   750
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgClient 
      Index           =   1
      Left            =   705
      Top             =   795
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgSchedule 
      Index           =   1
      Left            =   705
      Top             =   1395
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgClient16 
      Index           =   1
      Left            =   2535
      Top             =   735
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin NTControls22.BrowseButton BrowseButton1 
      Height          =   315
      Left            =   930
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      bBrowseTitle    =   "Browser for Project"
      bBrowseAction   =   0
      bFileFilter     =   "Max Project Files (*.mprj)|*.mprj|All Files (*.*)|*.*"
   End
   Begin VB.Label CmdLine 
      Height          =   465
      Left            =   165
      TabIndex        =   0
      Top             =   30
      Width           =   540
   End
   Begin VB.Menu mnuSysTray 
      Caption         =   "SysTray"
      Begin VB.Menu mnuShow 
         Caption         =   "Show..."
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuClientWin 
         Caption         =   "&New Client Window"
      End
      Begin VB.Menu mnuScheduleWin 
         Caption         =   "Schedule &Manager"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuActiveAppCache 
         Caption         =   "&Active App Cache"
      End
      Begin VB.Menu mnuDash23 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuScriptIDE 
         Caption         =   "&Scripting IDE"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRunScript 
         Caption         =   "&Run Script"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDash4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFavorites 
         Caption         =   "&Favorites"
         Begin VB.Menu mnuViewFav 
            Caption         =   "&Manage Favorites"
         End
         Begin VB.Menu mnuDash28764 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFavorite 
            Caption         =   "(No Favorites Found)"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuDash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetup 
         Caption         =   "&Setup"
         Begin VB.Menu mnuOptions 
            Caption         =   "Setup &Options"
         End
         Begin VB.Menu mnuFileAssociations 
            Caption         =   "&File Associations"
         End
         Begin VB.Menu mnuNetDrives 
            Caption         =   "Network &Drives"
         End
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
         Begin VB.Menu mnuContents 
            Caption         =   "&Documentation..."
         End
         Begin VB.Menu mnuWebSite 
            Caption         =   "&Neotext.org..."
         End
         Begin VB.Menu mnuAbout 
            Caption         =   "&About..."
         End
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit Max-FTP"
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

'Public WithEvents timGlobal As NTSchedule20.Timer

Public Sub RunScript()
    Dim fname As String
    fname = frmMain.BrowseButton1.Browse
    If PathExists(fname, True) Then
        MODPROCESS.RunProcess AppPath & MaxIDEFileName, "exec " & fname
    End If
End Sub
Private Sub CmdLine_Change()

    If CmdLine.Caption <> "" Then
        ExecuteFunction CmdLine.Caption
        CmdLine.Caption = ""
    End If

End Sub

Public Sub ResetToolTips(ByVal isON As Boolean)
    Dim xForm As Form
    For Each xForm In Forms
        Select Case TypeName(xForm)
            Case "frmFTPClientGUI"
                If Not IsHiddenForm(xForm) Then
                    xForm.SetTooltip
                End If
            Case "frmSchManager", "frmSchOperations"
                xForm.SetTooltip
        End Select
    Next
End Sub

Private Function RefreshFavorite2(ByRef cnt As Long, ByRef frm As Form, ByVal Directory As String) As Long

    Dim mnu
    Dim fso As New Scripting.FileSystemObject
    Dim f1 As Folder
    Dim f2 As File
    Dim f3 As Folder
    
    Set f1 = fso.GetFolder(Directory)
    For Each f2 In f1.Files
        If Right(LCase(Trim(f2.name)), Len(FTPSiteExt)) = FTPSiteExt Then
            If Not cnt = 0 Then Load frm.mnuFavorite(cnt)
            frm.mnuFavorite(cnt).Visible = True
            frm.mnuFavorite(cnt).Caption = Replace(f2.name, FTPSiteExt, "")
            frm.mnuFavorite(cnt).Tag = Directory & "\" & f2.name
            
            cnt = cnt + 1
        End If
    Next
    
    For Each f3 In f1.SubFolders
        RefreshFavorite2 cnt, frm, Directory & "\" & f3.name
    Next

End Function

Public Sub RefreshFavorite(ByRef frm As Form)

    Dim Directory As String
    Directory = GetMaxFavoritesDir(dbSettings.GetClientSetting("WinFavorites"))
    
    Dim mnu
    Dim cnt As Long
    Dim fso As New Scripting.FileSystemObject
    Dim f1 As Folder
    Dim f2 As File
    Dim f3 As Folder

    For Each mnu In frm.mnuFavorite
        If mnu.Index <> 0 Then Unload mnu
    Next
    
    frm.mnuFavorite(0).Caption = "(No Favorites Found)"
    
    If PathExists(Directory) Then
    
        Set f1 = fso.GetFolder(Directory)
        cnt = 0
        For Each f2 In f1.Files
            If Right(LCase(Trim(f2.name)), Len(FTPSiteExt)) = FTPSiteExt Then
                If Not cnt = 0 Then Load frm.mnuFavorite(cnt)
                frm.mnuFavorite(cnt).Visible = True
                frm.mnuFavorite(cnt).Caption = Replace(f2.name, FTPSiteExt, "")
                frm.mnuFavorite(cnt).Tag = Directory & "\" & f2.name
            
                cnt = cnt + 1
            End If
        Next
        
        For Each f3 In f1.SubFolders
            RefreshFavorite2 cnt, frm, Directory & "\" & f3.name
        Next
        
        frm.mnuFavorite(0).enabled = (cnt > 0)
    Else
        frm.mnuFavorite(0).enabled = False
        frm.mnuFavorite(0).Caption = "(No Favorites Found)"
    End If

End Sub

Public Sub RefreshFavorites()
    Dim frm As Form
    For Each frm In Forms
        If TypeName(frm) = "frmFTPClientGUI" Or TypeName(frm) = "frmMain" Then
            RefreshFavorite frm
        End If
    Next
End Sub

Public Sub RefreshShowAdvSettings()
    Dim frm As Form
    For Each frm In Forms
        Select Case TypeName(frm)
            Case "frmPassword"
                frm.sInfo.ShowAdvSettings = dbSettings.GetProfileSetting("ShowAdvSettings")
                If Not frm.sInfo.ShowAdvSettings Then
                    frm.sInfo.sPassive.Value = IIf((dbSettings.GetProfileSetting("ConnectionMode") = 0), 1, 0)
                    frm.sInfo.sSSL.Value = IIf((dbSettings.GetProfileSetting("SSL") = 0), 1, 0)
                End If
            Case "frmFavoriteSite", "frmSchOpProperties"
                frm.SiteInformation1.ShowAdvSettings = dbSettings.GetProfileSetting("ShowAdvSettings")
                If Not frm.SiteInformation1.ShowAdvSettings Then
                    frm.SiteInformation1.sPassive.Value = IIf((dbSettings.GetProfileSetting("ConnectionMode") = 0), 1, 0)
                    frm.SiteInformation1.sSSL.Value = IIf((dbSettings.GetProfileSetting("SSL") = 0), 1, 0)
                End If
                frm.SiteInformation2.ShowAdvSettings = dbSettings.GetProfileSetting("ShowAdvSettings")
                If Not frm.SiteInformation2.ShowAdvSettings Then
                    frm.SiteInformation2.sPassive.Value = IIf((dbSettings.GetProfileSetting("ConnectionMode") = 0), 1, 0)
                    frm.SiteInformation2.sSSL.Value = IIf((dbSettings.GetProfileSetting("SSL") = 0), 1, 0)
                End If
                
        End Select
    Next
End Sub

Public Sub ValidDataPortRange(ByRef txt)

    Dim NewVal As String
    
    Dim sHi As String
    Dim sLow As String
    
    NewVal = Replace(Replace(Replace(Replace(txt.Text, " ", ""), vbTab, ""), vbCr, ""), vbLf, "")
    
    sHi = NextArg(NewVal, "-")
    sLow = RemoveArg(NewVal, "-")
    If Not IsNumeric(sLow) Then
        sLow = sHi
        NewVal = sHi & "-" & sHi
    End If
    If Not IsNumeric(sHi) Then
        sHi = sLow
        NewVal = sLow & "-" & sLow
    End If
    
    If ((Not (InStr(NewVal, "-") > 0)) Or (Not IsNumeric(sHi)) Or (Not IsNumeric(sLow)) Or (InStr(sLow, ".") > 0) Or (InStr(sHi, ".") > 0) Or (InStr(sLow, "-") > 0) Or (InStr(sHi, "-") > 0)) Then
        MsgBox "Invalid data port range; (Must be one to two whole numbers seperated by a single dash greater then zero and less then 65535)", vbInformation, AppName
        txt.Text = dbSettings.GetProfileSetting("DefaultPortRange")
    Else
    
        If Not (((CLng(sHi) > 0) And (CLng(sLow) > 0)) And ((CLng(sLow) <= 65535) And (CLng(sHi) <= 65535))) Then
            MsgBox "Invalid data port range; (Must be one to two whole numbers seperated by a single dash greater then zero and less then 65535)", vbInformation, AppName
            txt.Text = dbSettings.GetProfileSetting("DefaultPortRange")
        Else
        
            If (CLng(sHi) < CLng(sLow)) Then
                Dim swp As String
                swp = sHi
                sHi = sLow
                sLow = swp
            End If
        
            txt.Text = sLow & "-" & sHi
        End If
    End If
    
End Sub


Public Sub Form_Load()
    
    Me.Caption = MaxMainFormCaption

    mnuScheduleWin.Visible = IsSchedulerInstalled
    mnuScriptIDE.Visible = IsScriptIDEInstalled
    mnuRunScript.Visible = IsScriptIDEInstalled
    mnuDash23.Visible = mnuScriptIDE.Visible

End Sub

Private Sub Form_Unload(Cancel As Integer)

    timGlobal.enabled = False
End Sub

Private Sub mnuActiveAppCache_Click()
    frmActiveCache.ShowForm
End Sub

Public Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseOverTray X

End Sub

Public Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Public Sub mnuClientWin_Click()
    ExecuteFunction "/client"
End Sub

Public Sub mnuContents_Click()
    GotoHelp
End Sub

Public Sub mnuExit_Click()
    ShutDownMaxFTP
End Sub

Public Sub mnuFavorite_Click(Index As Integer)
    If PathExists(mnuFavorite(Index).Tag) Then
    
        Dim ftpSite1 As New frmFavoriteSite
        ftpSite1.LoadSite mnuFavorite(Index).Tag
    
        Dim newClient3 As New frmFTPClientGUI
        newClient3.LoadClient
        newClient3.ShowClient
        newClient3.FTPOpenSite ftpSite1
        
        Unload ftpSite1
        
    End If
End Sub

Public Sub mnuFileAssociations_Click()
    frmFileAssoc.Show
End Sub


Public Sub mnuNetDrives_Click()
    frmNetDrives.Show
End Sub

Public Sub mnuOptions_Click()
    frmSetup.Show
End Sub


Public Sub mnuRunScript_Click()
    RunScript
End Sub

Public Sub mnuScheduleWin_Click()
    frmSchManager.Show
End Sub

Public Sub InitSysTrayMenu()
    Dim frm As Form, cnt As Integer
    Dim sMenu
    For Each sMenu In frmMain.mnuShow
        If sMenu.Index <> 0 Then
            Unload sMenu
        End If
    Next
    frmMain.mnuShow(0).Visible = False
    frmMain.mnuDash2.Visible = False
    
    cnt = 0
    For Each frm In Forms
        If IsHiddenForm(frm) Then
            If cnt > 0 Then
                Load frmMain.mnuShow(cnt)
            Else
                frmMain.mnuShow(0).Visible = True
                frmMain.mnuDash2.Visible = True
            End If
            frmMain.mnuShow(cnt).Caption = ShortName("Show " + frm.Caption)
            frmMain.mnuShow(cnt).Tag = frm.hwnd
            cnt = cnt + 1
        
        End If
    
    Next
    
    frmMain.RefreshFavorites
    
    modMenuUp.PopUp frmMain.hwnd, 0

End Sub

Function ShortName(ByVal LongName As String) As String

    Dim NewName As String
    If InStr(LongName, "-") > 0 Then
        NewName = Trim(Mid(LongName, InStr(LongName, "-")))
        LongName = Trim(Left(LongName, InStr(LongName, "-") - 1))
    Else
        NewName = ""
        End If
    If Len(LongName) > 25 Then
        LongName = Left(LongName, 25) + "... "
        End If
    If Trim(NewName) = "" Then
        NewName = LongName
    Else
        NewName = LongName + " " + NewName
        End If
    NewName = Replace(NewName, "&", "")
    ShortName = NewName

End Function

Public Sub mnuScriptIDE_Click()
    RunProcess AppPath & MaxIDEFileName, "", vbNormalFocus, False
End Sub

Public Sub mnuShow_Click(Index As Integer)
    
    Dim tForm As Form
    GetFormByHWND tForm, CLng(mnuShow(Index).Tag)
    
    tForm.Visible = True
    tForm.WindowState = 0
    If TypeName(tForm) = "frmFTPClientGUI" Then tForm.FormResize tForm
    
End Sub

Public Sub mnuViewFav_Click()
    frmFavorites.Show
End Sub

Public Sub mnuWebSite_Click()
    
    RunFile AppPath & "Neotext.org.url"

End Sub
Public Sub GlobalFunc()
    If MessageQueueLog(MaxFileName) > 0 Then
        ReadMessageQueue
    End If
    
    If dbSettings.GetProfileSetting("SystemTray") Then
        TrayIcon True
        PerformInTray
    ElseIf PerformOutTray Then
        TrayIcon False
    End If
End Sub
Private Sub ReadMessageQueue()
    Dim Messages As New Collection
    Dim Msg As Variant

    Set Messages = MessageQueueGet(MaxFileName)

    For Each Msg In Messages
        ExecuteFunction Msg
    Next

End Sub
Public Sub ProcessScheduleStatus(ByVal InParams As String)

    Dim frm As Form
    
    Select Case InParams
        Case "paused", "resumed"
            For Each frm In Forms
                If TypeName(frm) = "frmSchManager" Then
                    frm.SetServiceIcon (InParams = "paused")
                End If
            Next

        Case Else

            Dim nSID As Long
            nSID = CLng(RemoveNextArg(InParams, " "))
            For Each frm In Forms
                If TypeName(frm) = "frmSchOperations" Then
                    If frm.Sid = nSID Then
                        frm.UpdateStatus CLng(InParams)
                    End If
                End If
            Next
    
    End Select
    
End Sub

Private Sub timGlobal_Timer()
    GlobalFunc
End Sub
