VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5340
   ClientLeft      =   2415
   ClientTop       =   585
   ClientWidth     =   8835
   ClipControls    =   0   'False
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5340
   ScaleWidth      =   8835
   StartUpPosition =   2  'CenterScreen
   Tag             =   "prefrences"
   Visible         =   0   'False
   Begin VB.Frame Frame2 
      Height          =   4845
      Left            =   2775
      TabIndex        =   21
      Top             =   -45
      Width           =   6015
      Begin VB.PictureBox sProfileGeneral 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   4605
         Left            =   90
         ScaleHeight     =   4605
         ScaleWidth      =   5790
         TabIndex        =   23
         Top             =   150
         Visible         =   0   'False
         Width           =   5790
         Begin VB.CheckBox Check26 
            Caption         =   "Use Secure Sockets Layer as default option."
            Height          =   255
            Left            =   240
            TabIndex        =   87
            Top             =   1680
            Width           =   3615
         End
         Begin VB.CheckBox Check14 
            Caption         =   "Log schedule events to the LogFiles folder under Windows System  folder."
            Height          =   456
            Left            =   228
            TabIndex        =   84
            Top             =   2535
            Width           =   3036
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Left            =   495
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   1230
            Width           =   2145
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   2085
            Width           =   1950
         End
         Begin VB.TextBox Text2 
            Height          =   315
            Left            =   2715
            MaxLength       =   11
            TabIndex        =   32
            Text            =   "6000-9000"
            Top             =   1230
            Width           =   1230
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Show Advanced Connection Settings Per Login"
            Height          =   270
            Left            =   240
            TabIndex        =   31
            Top             =   75
            Width           =   3735
         End
         Begin VB.CheckBox Check10 
            Caption         =   "Show Mouseover Tool Tip Info"
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   4125
            Width           =   2820
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   2
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   450
            Width           =   1740
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Run Max-FTP in system tray mode"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   2
            Top             =   3135
            Width           =   2970
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Confirm aborting or closing when current connections are opened."
            Height          =   405
            Index           =   2
            Left            =   240
            TabIndex        =   3
            Top             =   3555
            Width           =   2760
         End
         Begin VB.Label Label7 
            Caption         =   "Data Connection Listening Adapter and Port Range:"
            Height          =   225
            Left            =   240
            TabIndex        =   51
            Top             =   885
            Width           =   3870
         End
         Begin VB.Label Label9 
            Caption         =   "Graphics Profile Folder:"
            Height          =   240
            Left            =   240
            TabIndex        =   43
            Top             =   2145
            Width           =   1755
         End
         Begin VB.Label Label1 
            Caption         =   "Default Connection Mode:"
            Height          =   270
            Index           =   2
            Left            =   225
            TabIndex        =   24
            Top             =   495
            Width           =   1890
         End
      End
      Begin VB.PictureBox sClientGeneral 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   4575
         Left            =   120
         ScaleHeight     =   4575
         ScaleWidth      =   5790
         TabIndex        =   70
         Top             =   180
         Visible         =   0   'False
         Width           =   5790
         Begin VB.CheckBox Check23 
            Caption         =   "Log client events to the LogFiles folder under the Windows System folder."
            Height          =   396
            Left            =   240
            TabIndex        =   83
            Top             =   2220
            Width           =   3240
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            Left            =   2610
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Top             =   180
            Width           =   1995
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   300
            Index           =   0
            Left            =   3492
            TabIndex        =   73
            Top             =   684
            Width           =   1080
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Index           =   0
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   1215
            Width           =   1860
         End
         Begin VB.ComboBox Combo1 
            Height          =   288
            Index           =   1
            Left            =   1365
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   1728
            Width           =   2115
         End
         Begin VB.Label Label1 
            Caption         =   "Default Internet Favorites Folder:"
            Height          =   300
            Index           =   0
            Left            =   225
            TabIndex        =   78
            Top             =   240
            Width           =   2370
         End
         Begin VB.Label Label2 
            Caption         =   "Max Size of FTP Log Data Display in bytes:"
            Height          =   288
            Index           =   0
            Left            =   228
            TabIndex        =   77
            Top             =   732
            Width           =   3336
         End
         Begin VB.Label Label3 
            Caption         =   "When files are Drag and Dropped:"
            Height          =   330
            Left            =   225
            TabIndex        =   76
            Top             =   1275
            Width           =   2580
         End
         Begin VB.Label Label2 
            Caption         =   "Overwrite files:"
            Height          =   252
            Index           =   1
            Left            =   252
            TabIndex        =   75
            Top             =   1788
            Width           =   1140
         End
      End
      Begin VB.PictureBox sPublic 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   4545
         Left            =   96
         ScaleHeight     =   4545
         ScaleWidth      =   5805
         TabIndex        =   27
         Top             =   180
         Visible         =   0   'False
         Width           =   5805
         Begin VB.CheckBox Check24 
            Caption         =   "Also use the Event Log when client or shedule actions are logged to the Windows\System32\Logfiles folder."
            Height          =   408
            Left            =   372
            TabIndex        =   85
            Top             =   3060
            Width           =   4248
         End
         Begin VB.CheckBox Check20 
            Caption         =   "No interface on scripts"
            Height          =   228
            Left            =   2520
            TabIndex        =   79
            Top             =   3684
            Width           =   1980
         End
         Begin VB.CheckBox Check19 
            Caption         =   "Stop transfering information when in stand by. (This option is not the same as service pause)"
            Height          =   465
            Left            =   375
            TabIndex        =   69
            Top             =   2472
            Width           =   3645
         End
         Begin VB.CheckBox Check17 
            Caption         =   "Allow anyone to stop or start the Schedule Service."
            Height          =   390
            Left            =   360
            TabIndex        =   50
            Top             =   2040
            Width           =   4005
         End
         Begin VB.CheckBox Check16 
            Caption         =   "Make all users of Max-FTP on this computer use a temporary profile.  (Saved data will not accumulate)."
            Height          =   555
            Left            =   345
            TabIndex        =   49
            Top             =   1488
            Width           =   3945
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Display all system files and folders when browsing."
            Height          =   225
            Left            =   345
            TabIndex        =   47
            Top             =   1188
            Width           =   4080
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Set Service to Manual"
            Height          =   372
            Index           =   1
            Left            =   2505
            TabIndex        =   41
            Top             =   4008
            Width           =   1965
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Set Service to Auto Start"
            Height          =   372
            Index           =   0
            Left            =   300
            TabIndex        =   40
            Top             =   4008
            Width           =   1965
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Keep session drives when Max-FTP isn't running."
            Height          =   285
            Left            =   345
            TabIndex        =   39
            Top             =   816
            Width           =   3975
         End
         Begin VB.CheckBox Check12 
            Caption         =   "Secure user infomation by domain, not computer, for users running Max-FTP form a shared folder."
            Height          =   420
            Left            =   345
            TabIndex        =   38
            Top             =   48
            Width           =   3840
         End
         Begin VB.Label Label10 
            Caption         =   "(Note: Aquiring domain may take a few moments)"
            Height          =   228
            Left            =   636
            TabIndex        =   45
            Top             =   516
            Width           =   3648
         End
         Begin VB.Label Label6 
            Caption         =   "Service Startup Options:"
            Height          =   240
            Left            =   360
            TabIndex        =   42
            Top             =   3696
            Width           =   1896
         End
      End
      Begin VB.PictureBox sClientLayout 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   4575
         Left            =   72
         ScaleHeight     =   4575
         ScaleWidth      =   5790
         TabIndex        =   22
         Top             =   108
         Visible         =   0   'False
         Width           =   5790
         Begin VB.CheckBox Check8 
            Caption         =   "Multi-Threaded transfer mode"
            Height          =   435
            Left            =   3048
            TabIndex        =   68
            Top             =   2136
            Width           =   1680
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Show Drive List"
            Height          =   285
            Index           =   1
            Left            =   684
            TabIndex        =   6
            Top             =   1728
            Width           =   1530
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Show Tool Bar"
            Height          =   315
            Index           =   1
            Left            =   684
            TabIndex        =   5
            Top             =   1308
            Width           =   1470
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Show Address Bar"
            Height          =   270
            Index           =   1
            Left            =   684
            TabIndex        =   7
            Top             =   2160
            Width           =   1785
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Show Double client window"
            Height          =   405
            Left            =   3048
            TabIndex        =   9
            Top             =   1656
            Width           =   1380
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Show FTP Log"
            Height          =   315
            Left            =   3048
            TabIndex        =   8
            Top             =   1260
            Width           =   1815
         End
         Begin VB.Label Label8 
            Caption         =   $"frmSetup.frx":08CA
            Height          =   810
            Left            =   240
            TabIndex        =   63
            Top             =   150
            Width           =   5310
         End
      End
      Begin VB.PictureBox sClientCache 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   4575
         Left            =   105
         ScaleHeight     =   4575
         ScaleWidth      =   5805
         TabIndex        =   33
         Top             =   165
         Visible         =   0   'False
         Width           =   5805
         Begin VB.CheckBox Check15 
            Caption         =   "Ask to remove Active App Cache when closing Max-FTP."
            Height          =   276
            Left            =   240
            TabIndex        =   48
            Top             =   1776
            Width           =   5076
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Keep open Active App Cache when new files are added."
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   37
            Top             =   1008
            Width           =   5088
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Remove files from Active App Cache when they are re-uploaded."
            Height          =   288
            Index           =   0
            Left            =   225
            TabIndex        =   36
            Top             =   612
            Width           =   5244
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Run files in their application when added to Active App Cache."
            Height          =   276
            Index           =   0
            Left            =   225
            TabIndex        =   35
            Top             =   255
            Width           =   5292
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Automaticlly upload files that are updated in Active App Cache."
            Height          =   276
            Index           =   2
            Left            =   240
            TabIndex        =   34
            Top             =   1392
            Width           =   5376
         End
      End
      Begin VB.PictureBox sAccessDenied 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   4575
         Left            =   90
         ScaleHeight     =   4575
         ScaleWidth      =   5805
         TabIndex        =   28
         Top             =   165
         Visible         =   0   'False
         Width           =   5805
         Begin VB.Label Label4 
            Caption         =   "This setting panel is only visible by Administrators of Max-FTP. (Contact your systems administrator whom installed Max-FTP)"
            Height          =   435
            Left            =   225
            TabIndex        =   29
            Top             =   180
            Width           =   4515
         End
      End
      Begin VB.PictureBox sProfileHistory 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   4560
         Left            =   120
         ScaleHeight     =   4560
         ScaleWidth      =   5805
         TabIndex        =   25
         Top             =   180
         Visible         =   0   'False
         Width           =   5805
         Begin VB.CheckBox Check11 
            Caption         =   "Lock History"
            Height          =   225
            Left            =   3390
            TabIndex        =   62
            Top             =   180
            Width           =   1245
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Clear Saved Sites"
            Height          =   405
            Left            =   180
            TabIndex        =   13
            Top             =   2490
            Width           =   1530
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   2400
            TabIndex        =   10
            Top             =   150
            Width           =   870
         End
         Begin VB.ListBox List1 
            Height          =   1230
            Left            =   180
            MultiSelect     =   2  'Extended
            TabIndex        =   11
            Top             =   585
            Width           =   4410
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Clear History"
            Height          =   405
            Index           =   3
            Left            =   2235
            TabIndex        =   12
            Top             =   2490
            Width           =   1230
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Remove"
            Height          =   405
            Index           =   1
            Left            =   3555
            TabIndex        =   14
            Top             =   2490
            Width           =   1050
         End
         Begin VB.Label Label1 
            Caption         =   "Max number of History Entries:"
            Height          =   270
            Index           =   3
            Left            =   165
            TabIndex        =   26
            Top             =   180
            Width           =   2205
         End
      End
      Begin VB.PictureBox sUsers 
         BorderStyle     =   0  'None
         Height          =   4455
         Left            =   150
         ScaleHeight     =   4455
         ScaleWidth      =   5730
         TabIndex        =   30
         Top             =   240
         Width           =   5730
      End
      Begin VB.PictureBox sProfileTransfer 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4575
         Left            =   105
         ScaleHeight     =   4575
         ScaleWidth      =   5805
         TabIndex        =   53
         Top             =   165
         Visible         =   0   'False
         Width           =   5805
         Begin VB.CheckBox Check25 
            Caption         =   "Auto adjust the transfer rates with these starting values:"
            Height          =   192
            Left            =   1380
            TabIndex        =   86
            Top             =   1848
            Width           =   4275
         End
         Begin VB.CheckBox Check22 
            Caption         =   "Client side allocation"
            Height          =   195
            Left            =   405
            TabIndex        =   81
            Top             =   585
            Visible         =   0   'False
            Width           =   1905
         End
         Begin VB.CheckBox Check21 
            Caption         =   "Server side allocation"
            Height          =   270
            Left            =   405
            TabIndex        =   80
            Top             =   285
            Width           =   1890
         End
         Begin VB.CheckBox Check18 
            Caption         =   $"frmSetup.frx":09C1
            Height          =   705
            Left            =   435
            TabIndex        =   66
            Top             =   1005
            Width           =   5220
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   4680
            TabIndex        =   64
            Text            =   "30"
            Top             =   165
            Width           =   1005
         End
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3495
            TabIndex        =   61
            Text            =   "16384"
            Top             =   3600
            Width           =   2175
         End
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3495
            TabIndex        =   60
            Text            =   "65536"
            Top             =   3195
            Width           =   2175
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3495
            TabIndex        =   59
            Text            =   "1048576"
            Top             =   2790
            Width           =   2175
         End
         Begin VB.Label Label5 
            Caption         =   "Reservation of file space:"
            Height          =   210
            Left            =   390
            TabIndex        =   82
            Top             =   75
            Width           =   1920
         End
         Begin VB.Label Label18 
            Caption         =   " (Read and write buffer sizes to hard drive and random access memory.  Adjusting these can improve or reduce performance)"
            Height          =   435
            Left            =   1260
            TabIndex        =   67
            Top             =   2205
            Width           =   4500
         End
         Begin VB.Label Label1 
            Caption         =   "Timeout in seconds:"
            Height          =   300
            Index           =   1
            Left            =   3075
            TabIndex        =   65
            Top             =   195
            Width           =   1530
         End
         Begin VB.Label Label17 
            Caption         =   "Note: Remote server to remote server transfers are not determined by this client and self rely on the server setup."
            Height          =   525
            Left            =   480
            TabIndex        =   58
            Top             =   4080
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.Label Label16 
            Caption         =   "Remote to local, or internet downloads:"
            Height          =   225
            Left            =   465
            TabIndex        =   57
            Top             =   3240
            Width           =   2970
         End
         Begin VB.Label Label15 
            Caption         =   "Local to remote, or internet uploads:"
            Height          =   225
            Left            =   450
            TabIndex        =   56
            Top             =   3645
            Width           =   2805
         End
         Begin VB.Label Label14 
            Caption         =   "Local to local, or local to MSN network:"
            Height          =   225
            Left            =   450
            TabIndex        =   55
            Top             =   2820
            Width           =   3015
         End
         Begin VB.Label Label13 
            Caption         =   "Transfer Rates:"
            Height          =   270
            Left            =   150
            TabIndex        =   54
            Top             =   1845
            Width           =   1155
         End
      End
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4740
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   8361
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Index           =   2
      Left            =   6285
      TabIndex        =   15
      Top             =   4890
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   375
      Index           =   0
      Left            =   7635
      TabIndex        =   20
      Top             =   4890
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Caption         =   "General prefrences"
      Height          =   660
      Index           =   0
      Left            =   1380
      TabIndex        =   18
      Top             =   5880
      Visible         =   0   'False
      Width           =   2685
      Begin VB.CheckBox InfoDialog 
         Caption         =   "Show Info Dialog at startup"
         Height          =   210
         Left            =   195
         TabIndex        =   19
         Top             =   285
         Width           =   2368
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Toolbars"
      Height          =   660
      Left            =   4620
      TabIndex        =   16
      Top             =   5955
      Visible         =   0   'False
      Width           =   2190
      Begin VB.CheckBox Check4 
         Caption         =   "Display FTP Toolbars"
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   17
         Top             =   255
         Width           =   1968
      End
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "(User profile set is to be removed)"
      Height          =   270
      Left            =   90
      TabIndex        =   46
      Top             =   4935
      Visible         =   0   'False
      Width           =   2535
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public MultiThread As Boolean


Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 2
            If SaveProperties Then Unload Me
        Case 1
            If MsgBox("Are you sure you want to delete the selected items?", vbQuestion + vbYesNo, AppName) = vbYes Then
                
                Dim dbConn1 As New clsDBConnection
                Dim rs1 As New ADODB.Recordset
                Dim cnt As Integer

                If List1.ListCount > 0 Then
                
                    cnt = 0
                    Do
                        If List1.Selected(cnt) Then
                            dbConn1.rsQuery rs1, "DELETE * FROM History WHERE ParentID=" & dbSettings.CurrentUserID & " AND URL='" & Replace(List1.List(cnt), "'", "''") & "';"
                            List1.RemoveItem cnt
                        Else
                            cnt = cnt + 1
                        End If
                    Loop Until cnt > List1.ListCount - 1
                End If
                
                If Not rs1.State = 0 Then rs1.Close
                Set rs1 = Nothing
                Set dbConn1 = Nothing

                List1.Tag = "UPDATE"
            End If
        Case 3
            If MsgBox("Are you sure you want to completely clear the history?", vbQuestion + vbYesNo, AppName) = vbYes Then
                
                Dim dbConn As New clsDBConnection
                Dim rs As New ADODB.Recordset

                dbConn.rsQuery rs, "DELETE * FROM History WHERE ParentID=" & dbSettings.CurrentUserID & ";"
                
                List1.Clear
                
                If Not rs.State = 0 Then rs.Close
                Set rs = Nothing
                Set dbConn = Nothing

                List1.Tag = "UPDATE"
            End If
    End Select
End Sub

Private Sub Command2_Click()
    If MsgBox("Are you sure you want to clear the saved sites info? (Connection" & vbCrLf & _
                "specific information such as username, password settings for sites)", vbQuestion + vbYesNo, AppName) = vbYes Then
        
        Dim dbConn As New clsDBConnection
        Dim rs As New ADODB.Recordset

        dbConn.rsQuery rs, "DELETE * FROM SiteCache WHERE ParentID=" & dbSettings.CurrentUserID & ";"
       
        If Not rs.State = 0 Then rs.Close
        Set rs = Nothing
        Set dbConn = Nothing

    End If
End Sub

Private Function ServiceChange(ByVal Inform As String, ByVal Params As String) As Boolean

    If Not ((ProcessRunning(ServiceFileName) = 0) And (Not IsDebugger(GetFileTitle(ServiceFileName)))) Then

        If Not (MsgBox("To change the service " & Inform & " option setup must stop and restart the schedule service," & vbCrLf & "Do you wish to continue and restart the service?", vbYesNo + vbQuestion, AppName) = vbNo) Then
            Me.MousePointer = 11
            NetStop MaxServiceName
            'RunProcess AppPath & ServiceFileName, "/uninstall", 0, True
            'RunProcess AppPath & ServiceFileName, "/install " & Trim(CStr(Index = 0)), 0, True
            RunProcess AppPath & ServiceFileName, Params, 0, True
            NetStart MaxServiceName
            Me.MousePointer = 0
            ServiceChange = True
        End If
    Else
        Me.MousePointer = 11
        'RunProcess AppPath & ServiceFileName, "/uninstall", 0, True
        'RunProcess AppPath & ServiceFileName, "/install " & Trim(CStr(Index = 0)), 0, True
        RunProcess AppPath & ServiceFileName, Params, 0, True
        Me.MousePointer = 0
        ServiceChange = True
    End If
End Function

Private Sub Check20_Click()
    If Check20.Tag = "ON" Then
    
        If Not ServiceChange("interface", "/interactive " & Trim(CStr(Abs(CInt(Check20.Value = 0))))) Then
            Check20.Value = BoolToCheck(Not dbSettings.GetPublicSetting("ServiceInterface"))
        Else
            dbSettings.SetPublicSetting "ServiceInterface", (Check20.Value = 0)
        End If
        Dim frm As Form
        For Each frm In Forms
            If TypeName(frm) = "frmSchOpProperties" Then
                frm.RefreshDisabilities
            End If
        Next
    End If
End Sub

Private Sub Command3_Click(Index As Integer)

    ServiceChange "startup", "/startup " & Trim(CStr(Index = 0))

End Sub

Private Sub Form_Load()

    Me.Caption = AppName + " Setup - [" & dbSettings.GetUserLoginName & "]"

    Dim treeNode As MSComctlLib.Node
    
    Set treeNode = TreeView1.Nodes.Add(, , "Public", "Public")
    
    Set treeNode = TreeView1.Nodes.Add(, , "Profile", "Profile")
    treeNode.Expanded = True
    Set treeNode = TreeView1.Nodes.Add("Profile", tvwChild, "ProfileGeneral", "General")
    Set treeNode = TreeView1.Nodes.Add("Profile", tvwChild, "ProfileHistory", "History")
    Set treeNode = TreeView1.Nodes.Add("Profile", tvwChild, "ProfileTransfer", "Transfer")
    Set treeNode = TreeView1.Nodes.Add(, , "Client", "Client")
    treeNode.Expanded = True
    Set treeNode = TreeView1.Nodes.Add("Client", tvwChild, "ClientGeneral", "General")
    Set treeNode = TreeView1.Nodes.Add("Client", tvwChild, "ClientLayout", "Layout")
    Set treeNode = TreeView1.Nodes.Add("Client", tvwChild, "ClientCache", "Cache")
    
    Set TreeView1.SelectedItem = TreeView1.Nodes("Public")
    TreeView1_NodeClick TreeView1.Nodes("Public")
    
    sAccessDenied.Move 150, 150, 5715, 4605
    sPublic.Move 150, 150, 5715, 4605
    sUsers.Move 150, 150, 5715, 4605
    sProfileGeneral.Move 150, 150, 5715, 4605
    sProfileHistory.Move 150, 150, 5715, 4605
    sProfileTransfer.Move 150, 150, 5715, 4605
    sClientLayout.Move 150, 150, 5715, 4605
    sClientGeneral.Move 150, 150, 5715, 4605
    sClientCache.Move 150, 150, 5715, 4605
    
    InitPropBoxes
    
    Dim col As Collection
    Dim ftp As New NTAdvFTP61.Client
    Set col = ftp.AllAdapters
    Set ftp = Nothing
    
    If col.count > 0 Then
        Dim cnt As Long
        For cnt = 1 To col.count
            Combo4.AddItem col.Item(cnt)
        Next
    Else
        Combo4.AddItem "(Unknown)"
    End If
    
    LoadProperties
    
    Label12.Visible = dbSettings.RemoveProfile

End Sub

Public Sub CheckNonemptyText(ByVal Display As String, ByVal DefaultValue As String, ByRef txtBox As TextBox)
    If (txtBox.Text = "") Then
        MsgBox Display & " must have a value.", vbInformation, AppName
        txtBox.Text = DefaultValue

    End If
End Sub

Public Sub CheckNumericalText(ByVal Display As String, ByVal DefaultValue As Long, ByRef txtBox As TextBox, Optional ByVal ZeroOK As Boolean = False)
    If Not IsAlphaNumeric(txtBox.Text) Then
        MsgBox Display & " must be a numerical value greater then zero.", vbInformation, AppName
        txtBox.Text = DefaultValue
    ElseIf (CSng(txtBox.Text) <= 0) Then
        MsgBox Display & " must be a numerical value greater then zero.", vbInformation, AppName
        txtBox.Text = DefaultValue
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    If Index = 0 Then
        CheckNumericalText "Max Size of Log Data in bytes", dbSettings.GetClientSetting("LogFileSize"), Text1(0)
    ElseIf Index = 1 Then
        CheckNumericalText "Internet Timeout in seconds", dbSettings.GetProfileSetting("TimeOut"), Text1(1)
    ElseIf Index = 3 Then
        CheckNumericalText "Max number of History Entries", dbSettings.GetProfileSetting("HistorySize"), Text1(3)
    End If
End Sub

Private Sub Text2_LostFocus()
    
    CheckNonemptyText "Default Port Range", dbSettings.GetProfileSetting("DefaultPortRange"), Text2
    
    frmMain.ValidDataPortRange Text2
    
End Sub

Private Sub Text3_LostFocus()
    CheckNumericalText "Transfer rates must be a numerical value greater then zero", dbSettings.GetProfileSetting("ftpLocalSize"), Text3
End Sub
Private Sub Text4_LostFocus()
    CheckNumericalText "Transfer rates must be a numerical value greater then zero", dbSettings.GetProfileSetting("ftpPacketSize"), Text4
End Sub
Private Sub Text5_LostFocus()
    CheckNumericalText "Transfer rates must be a numerical value greater then zero", dbSettings.GetProfileSetting("ftpBufferSize"), Text5
End Sub

Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)
    If Node.Key = "Profile" Or Node.Key = "Client" Then
        Node.Expanded = True
        If (Node.Key = "Profile") Then Set TreeView1.SelectedItem = TreeView1.Nodes("ProfileGeneral")
        If (Node.Key = "Client") Then Set TreeView1.SelectedItem = TreeView1.Nodes("ClientGeneral")
    End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

    Dim myKey As String
    myKey = Node.Key
    
    If dbSettings.CurrentUserAccessRights = ar_Administrator Then
        sPublic.Visible = (myKey = "Public")
        sUsers.Visible = (myKey = "Users")
        sAccessDenied.Visible = False
    Else
        sPublic.Visible = False
        sUsers.Visible = False
        sAccessDenied.Visible = (myKey = "Public") Or (myKey = "Users")
    End If
    
    sProfileGeneral.Visible = (myKey = "ProfileGeneral") Or (myKey = "Profile")
    sProfileHistory.Visible = (myKey = "ProfileHistory")
    sProfileTransfer.Visible = (myKey = "ProfileTransfer")

    sClientGeneral.Visible = (myKey = "ClientGeneral") Or (myKey = "Client")
    sClientLayout.Visible = (myKey = "ClientLayout")
    sClientCache.Visible = (myKey = "ClientCache")
    
    If (myKey = "Profile") Then Set TreeView1.SelectedItem = TreeView1.Nodes("ProfileGeneral")
    If (myKey = "Client") Then Set TreeView1.SelectedItem = TreeView1.Nodes("ClientGeneral")

End Sub

Private Sub InitPropBoxes()

    If Not IsSchedulerInstalled Then
        Label6.Visible = False
        Command3(0).Visible = False
        Command3(1).Visible = False
        Check17.Visible = False
    End If

    Combo1(2).AddItem "Passive"
    Combo1(2).AddItem "Active"
    
    Combo3.AddItem "(None)"
    
    Dim fso As New Scripting.FileSystemObject
    Dim fld As Folder
    For Each fld In fso.GetFolder(AppPath & GraphicsFolder).SubFolders
        If Not (LCase(fld.name) = "(none)") Then
            Combo3.AddItem fld.name
        End If
    Next
    
    SetAutoTypeList Me, List1

    Combo1(0).AddItem "Windows Favorites"
    Combo1(0).AddItem "MaxFTP Favorites"
    
    Combo2(0).AddItem "Ask"
    Combo2(0).AddItem "Copy"
    Combo2(0).AddItem "Move"

    Combo1(1).AddItem "Always Ask"
    Combo1(1).AddItem "Never Overwrite"
    Combo1(1).AddItem "Always Overwrite"
    Combo1(1).AddItem "Automatic Decide"
 End Sub
Private Function TextExists(ByRef combo As ComboBox, ByVal Text As String) As Boolean
    Dim cnt As Long
    If combo.ListCount > 0 Then
        For cnt = 0 To combo.ListCount
            If combo.List(cnt) = Text Then
                TextExists = True
            End If
        Next
    End If
End Function
Private Sub LoadProperties()
    Dim enc As New NTCipher10.ncode

    Check25.Value = BoolToCheck(dbSettings.GetProfileSetting("ftpAutoRate"))
    Check12.Value = BoolToCheck(dbSettings.GetPublicSetting("ServiceNetwork"))
    Check13.Value = BoolToCheck(dbSettings.GetPublicSetting("ServiceSession"))
    Check9.Value = BoolToCheck(dbSettings.GetPublicSetting("ServiceSystem"))
    Check16.Value = BoolToCheck(dbSettings.GetPublicSetting("ServiceReadOnly"))
    Check17.Value = BoolToCheck(dbSettings.GetPublicSetting("ServiceAllowAny"))
    Check20.Value = BoolToCheck(Not dbSettings.GetPublicSetting("ServiceInterface"))
    Check24.Value = BoolToCheck(dbSettings.GetPublicSetting("ServiceEventLog"))
    
    Check14.Value = BoolToCheck(dbSettings.GetProfileSetting("EventLog"))
    
    Check6.Value = BoolToCheck(dbSettings.GetProfileSetting("ShowAdvSettings"))
    Text2.Text = dbSettings.GetProfileSetting("DefaultPortRange")
    Combo1(2).ListIndex = dbSettings.GetProfileSetting("ConnectionMode")
    Check1(2).Value = BoolToCheck(dbSettings.GetProfileSetting("SystemTray"))
    Check2(2).Value = BoolToCheck(dbSettings.GetProfileSetting("PromptAbortClose"))
    Check8.Value = BoolToCheck(dbSettings.GetClientSetting("MultiThread"))
    MultiThread = dbSettings.GetClientSetting("MultiThread")
    Check10.Value = BoolToCheck(dbSettings.GetProfileSetting("ViewToolTips"))
    
    If TextExists(Combo3, dbSettings.GetProfileSetting("GraphicsFolder")) Then
        Combo3.Text = dbSettings.GetProfileSetting("GraphicsFolder")
    Else
        Combo3.ListIndex = 0
    End If
    
    Text1(3).Text = dbSettings.GetProfileSetting("HistorySize")
    Check11.Value = BoolToCheck(dbSettings.GetProfileSetting("HistoryLock"))
    
    Check18.Value = BoolToCheck(dbSettings.GetProfileSetting("LargeFileMode"))

    If Combo4.ListCount >= dbSettings.GetProfileSetting("AdapterIndex") Then
        Combo4.ListIndex = (dbSettings.GetProfileSetting("AdapterIndex") - 1)
    ElseIf Combo4.ListCount >= 1 Then
        Combo4.ListIndex = 0
    End If
    
    Check26.Value = dbSettings.GetProfileSetting("SSL")
    Text3.Text = dbSettings.GetProfileSetting("ftpLocalSize")
    Text4.Text = dbSettings.GetProfileSetting("ftpBufferSize")
    Text5.Text = dbSettings.GetProfileSetting("ftpPacketSize")

    Check21.Value = BoolToCheck(dbSettings.GetProfileSetting("ServerAlloc"))
    Check22.Value = BoolToCheck(dbSettings.GetProfileSetting("ClientAlloc"))
    
    Check1(1).Value = BoolToCheck(dbSettings.GetClientSetting("ViewToolBar"))
    Check2(1).Value = BoolToCheck(dbSettings.GetClientSetting("ViewDriveList"))
    Check3(1).Value = BoolToCheck(dbSettings.GetClientSetting("ViewAddressBar"))
    Check7.Value = BoolToCheck(dbSettings.GetClientSetting("ViewLog"))
    Check5.Value = BoolToCheck(dbSettings.GetClientSetting("ViewDoubleWindow"))
    Check23.Value = BoolToCheck(dbSettings.GetClientSetting("EventLog"))
    
    Combo1(0).ListIndex = IIf(dbSettings.GetClientSetting("WinFavorites"), 0, 1)
    Text1(0).Text = dbSettings.GetClientSetting("LogFileSize")
    Combo2(0).ListIndex = dbSettings.GetClientSetting("DragOption")

    Text1(1).Text = dbSettings.GetProfileSetting("TimeOut")
    Combo1(1).ListIndex = dbSettings.GetClientSetting("Overwrite")

    Check1(0).Value = BoolToCheck(dbSettings.GetClientSetting("ActiveAppRun"))
    Check2(0).Value = BoolToCheck(dbSettings.GetClientSetting("ActiveAppRemove"))
    Check3(0).Value = BoolToCheck(dbSettings.GetClientSetting("ActiveAppOpen"))
    Check3(2).Value = BoolToCheck(dbSettings.GetClientSetting("ActiveAppUpload"))
    Check15.Value = BoolToCheck(dbSettings.GetClientSetting("ActiveAppAsk"))
    
    Check19.Value = BoolToCheck(dbSettings.GetPublicSetting("ServiceStandBy"))
    Set enc = Nothing

End Sub

Private Function SaveProperties() As Boolean
    Dim enc As New NTCipher10.ncode
    
    Dim xForm
    Dim Cancel As Boolean
    Cancel = False
    If IsNumeric(Text3.Text) And IsNumeric(Text4.Text) And IsNumeric(Text5.Text) Then
    
        If (Not (CCur(Text3.Text) <= modBitValue.LongBound And CCur(Text3.Text) > 0)) Or _
            (Not (CCur(Text4.Text) <= modBitValue.LongBound And CCur(Text4.Text) > 0)) Or _
            (Not (CCur(Text5.Text) <= modBitValue.LongBound And CCur(Text5.Text) > 0)) Then
        
            MsgBox "Transfer buffer sizes must be a above zero, and not exceeding " & modBitValue.LongBound & ".", vbInformation, vbOKOnly
            SaveProperties = False
            Exit Function
            
        End If
    Else
            MsgBox "Transfer buffer sizes must be a number a above zero, and not exceeding " & modBitValue.LongBound & ".", vbInformation, vbOKOnly
            SaveProperties = False
            Exit Function
    End If
    
    If Not (dbSettings.GetPublicSetting("ServiceNetwork") = (Check12.Value = 1)) Then
      
        Dim tmp As String
        If Trim(UCase(dbSettings.GetMachineName)) = "MSHOME" Or Trim(UCase(dbSettings.GetMachineName)) = "WORKGROUP" Then
            tmp = "This option is not recomended on the " & Trim(UCase(dbSettings.GetMachineName)) & " workgroup." & vbCrLf & vbCrLf
        End If
                
        If MsgBox("Warning: Changing the network roaming status of users data " & vbCrLf & _
                    "will corrupt any current data in the database.  You may wish " & vbCrLf & _
                    "to use the " & MaxUtilityFileName & " program and backup all your current" & vbCrLf & _
                    "database before continuing, such data may not be imported" & vbCrLf & _
                    "unless you reverse any changes you've made to this option." & vbCrLf & vbCrLf & tmp & "Do you still wish to continue?", vbYesNo + vbCritical, AppName) = vbYes Then
        
                    If Not (ProcessRunning(ServiceFileName) = 0) Then
                        MessageQueueAdd ServiceFileName, "unloadschedules"
                        MessageQueueAdd ServiceFileName, "loadschedules"
                    End If
                    
            dbSettings.SetPublicSetting "ServiceNetwork", (Check12.Value = 1)
        Else
            Cancel = True
        End If
        
    End If
    
    dbSettings.SetProfileSetting "ftpAutoRate", (Check25.Value = 1)
    
    dbSettings.SetPublicSetting "ServiceSession", (Check13.Value = 1)
    dbSettings.SetPublicSetting "ServiceSystem", (Check9.Value = 1)
    dbSettings.SetPublicSetting "ServiceReadOnly", (Check16.Value = 1)
    dbSettings.SetPublicSetting "ServiceAllowAny", (Check17.Value = 1)
    dbSettings.SetPublicSetting "ServiceInterface", Not (Check20.Value = 1)
    dbSettings.SetPublicSetting "ServiceEventLog", (Check24.Value = 1)
    
    dbSettings.SetProfileSetting "EventLog", (Check14.Value = 1)
    dbSettings.SetProfileSetting "SSL", Check26.Value
    dbSettings.SetProfileSetting "ShowAdvSettings", (Check6.Value = 1)
    frmMain.RefreshShowAdvSettings
    dbSettings.SetProfileSetting "DefaultPortRange", Text2.Text
    dbSettings.SetProfileSetting "ConnectionMode", Combo1(2).ListIndex
    dbSettings.SetProfileSetting "SystemTray", (Check1(2).Value = 1)

    dbSettings.SetProfileSetting "PromptAbortClose", (Check2(2).Value = 1)
    dbSettings.SetClientSetting "MultiThread", (Check8.Value = 1)
    dbSettings.SetProfileSetting "ViewToolTips", (Check10.Value = 1)
    
    frmMain.ResetToolTips (Check10.Value = 1)
    If Not (dbSettings.GetProfileSetting("GraphicsFolder") = Combo3.Text) Then
        MsgBox "Changing the graphics profile folder will not take effect until you re-run Max-FTP.", vbOKOnly + vbInformation, AppName
    End If
    dbSettings.SetProfileSetting "GraphicsFolder", Combo3.Text
    dbSettings.SetProfileSetting "HistorySize", Text1(3).Text
    dbSettings.SetProfileSetting "HistoryLock", (Check11.Value = 1)
    If List1.Tag = "UPDATE" Then UpdateAutoTypeLists
    
    If Combo4.Visible Then
        dbSettings.SetProfileSetting "AdapterIndex", (Combo4.ListIndex + 1)
    End If
 
    dbSettings.SetProfileSetting "ServerAlloc", (Check21.Value = 1)
    dbSettings.SetProfileSetting "ClientAlloc", (Check22.Value = 1)
    
    dbSettings.SetProfileSetting "LargeFileMode", (Check18.Value = 1)
    
    dbSettings.SetProfileSetting "ftpLocalSize", Text3.Text
    dbSettings.SetProfileSetting "ftpBufferSize", Text4.Text
    dbSettings.SetProfileSetting "ftpPacketSize", Text5.Text
    For Each xForm In Forms
        If TypeName(xForm) = "frmFTPClientGUI" Then
            xForm.myClient0.AutoRate = dbSettings.GetProfileSetting("ftpAutoRate")
            xForm.myClient1.AutoRate = dbSettings.GetProfileSetting("ftpAutoRate")
            xForm.myClient1.TransferRates(NTAdvFTP61.RateTypes.HardDrive) = CLng(Text3.Text)
            xForm.myClient1.TransferRates(NTAdvFTP61.RateTypes.Download) = CLng(Text4.Text)
            xForm.myClient1.TransferRates(NTAdvFTP61.RateTypes.Upload) = CLng(Text5.Text)
        End If
    Next
    
    dbSettings.SetClientSetting "ViewToolBar", (Check1(1).Value = 1)
    dbSettings.SetClientSetting "ViewDriveList", (Check2(1).Value = 1)
    dbSettings.SetClientSetting "ViewAddressBar", (Check3(1).Value = 1)
        
    dbSettings.SetClientSetting "ViewLog", (Check7.Value = 1)
    dbSettings.SetClientSetting "ViewDoubleWindow", (Check5.Value = 1)
    dbSettings.SetClientSetting "EventLog", (Check23.Value = 1)
    
    If (MultiThread And Not dbSettings.GetClientSetting("MultiThread")) Or (Not MultiThread And dbSettings.GetClientSetting("MultiThread")) Then
        For Each xForm In Forms
            If TypeName(xForm) = "frmFTPClientGUI" Then
                xForm.EnableTransferInfo dbSettings.GetClientSetting("MultiThread")
            End If
        Next
    End If

    dbSettings.SetClientSetting "WinFavorites", (Combo1(0).ListIndex = 0)
    frmMain.RefreshFavorites
    
    dbSettings.SetClientSetting "LogFileSize", Text1(0).Text
    dbSettings.SetClientSetting "DragOption", Combo2(0).ListIndex

    dbSettings.SetProfileSetting "TimeOut", Text1(1).Text
    dbSettings.SetClientSetting "Overwrite", Combo1(1).ListIndex

    dbSettings.SetClientSetting "ActiveAppRun", (Check1(0).Value = 1)
    dbSettings.SetClientSetting "ActiveAppRemove", (Check2(0).Value = 1)
    dbSettings.SetClientSetting "ActiveAppOpen", (Check3(0).Value = 1)
    dbSettings.SetClientSetting "ActiveAppUpload", (Check3(2).Value = 1)
    dbSettings.SetClientSetting "ActiveAppAsk", (Check15.Value = 1)
    
    dbSettings.SetPublicSetting "ServiceStandBy", (Check19.Value = 1)
    For Each xForm In Forms
        If TypeName(xForm) = "frmFTPClientGUI" Then
            xForm.myClient0.PauseOnStandBy = (Check19.Value = 1)
            xForm.myClient1.PauseOnStandBy = (Check19.Value = 1)
        End If
    Next
        
    Set enc = Nothing
    
    Check20.Tag = "ON"
    
    InitSystemTray
    
    SaveProperties = (Not Cancel)
End Function
