VERSION 5.00
Begin VB.Form frmUtility 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Max-FTP Database Utility"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   Icon            =   "frmUtility.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5565
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   2100
      Top             =   450
   End
   Begin VB.Frame Frame4 
      Height          =   450
      Left            =   75
      TabIndex        =   11
      Top             =   -30
      Width           =   5415
      Begin VB.OptionButton Option3 
         Caption         =   "Reset Database"
         Height          =   195
         Left            =   3840
         TabIndex        =   2
         Top             =   165
         Width           =   1485
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Restore Database"
         Height          =   195
         Left            =   1965
         TabIndex        =   1
         Top             =   165
         Width           =   1620
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Backup Database"
         Height          =   195
         Left            =   135
         TabIndex        =   0
         Top             =   165
         Value           =   -1  'True
         Width           =   1635
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   75
      TabIndex        =   7
      Top             =   510
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   150
         TabIndex        =   13
         Top             =   2130
         Width           =   4695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Restore"
         Enabled         =   0   'False
         Height          =   360
         Left            =   4185
         TabIndex        =   15
         Top             =   2505
         Width           =   1065
      End
      Begin MaxUtility.ctlOptions ctlDBOptions2 
         Height          =   1800
         Left            =   450
         TabIndex        =   4
         Top             =   285
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   3175
      End
      Begin VB.Label Label1 
         Caption         =   "Enter the file to restore data from in the above text box."
         Height          =   255
         Left            =   165
         TabIndex        =   9
         Top             =   2580
         Width           =   3930
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3015
      Left            =   75
      TabIndex        =   8
      Top             =   510
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CommandButton Command3 
         Caption         =   "Reset"
         Enabled         =   0   'False
         Height          =   360
         Left            =   4185
         TabIndex        =   17
         Top             =   2505
         Width           =   1065
      End
      Begin MaxUtility.ctlOptions ctlDBOptions3 
         Height          =   1800
         Left            =   450
         TabIndex        =   5
         Top             =   285
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   3175
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   75
      TabIndex        =   6
      Top             =   510
      Width           =   5415
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   150
         TabIndex        =   14
         Top             =   2130
         Width           =   4695
      End
      Begin MaxUtility.ctlOptions ctlDBOptions1 
         Height          =   1800
         Left            =   450
         TabIndex        =   3
         Top             =   285
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   3175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Backup"
         Enabled         =   0   'False
         Height          =   360
         Left            =   4185
         TabIndex        =   16
         Top             =   2505
         Width           =   1065
      End
      Begin VB.Label Label3 
         Caption         =   "Enter the file to backup data to in the above text box."
         Height          =   255
         Left            =   165
         TabIndex        =   10
         Top             =   2580
         Width           =   4005
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   240
      Left            =   90
      TabIndex        =   12
      Top             =   3585
      Width           =   5355
   End
End
Attribute VB_Name = "frmUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private DatabaseOpen As Boolean
Dim BrowseButton1LB As Control
Attribute BrowseButton1LB.VB_VarHelpID = -1

Private Sub ShowTask()
    Frame1.Visible = Option1.Value
    Frame2.Visible = Option2.Value
    Frame3.Visible = Option3.Value
    If Frame2.Visible Then Text1.Tag = ""
    If Frame1.Visible Then Text2.Tag = ""
End Sub
Private Sub EnableForm(ByVal IsEnabled As Boolean)
    If IsEnabled Then
        Me.MousePointer = 0
    Else
        Me.MousePointer = 11
    End If
    Frame1.enabled = IsEnabled
    Frame2.enabled = IsEnabled
    Frame3.enabled = IsEnabled
    Frame4.enabled = IsEnabled
End Sub
Public Function GetRestoreOptions()
    If PathExists(Text1.Text, True) Then
        Dim archive As New clsArchive
        ctlDBOptions2.Options = archive.GetOptionsFromBackup(Text1.Text)
        Set archive = Nothing
    Else
        ctlDBOptions2.Options = 0
    End If
End Function

Private Sub SetProps(ByRef OBJ As Object)
    OBJ.BorderStyle = 2
    OBJ.BrowseAction = 0
    OBJ.BrowseTitle = "Backup to File"
    OBJ.FileFilter = "Max-FTP Database Backups|*.madb"
    OBJ.FileFilterIndex = 1
End Sub
Private Sub Form_Load()
    
    Set BrowseButton1LB = Controls.Add("NTControls22.BrowseButton", "BrowseButton1LB", Me)
    
    BrowseButton1LB.Left = 4995
    BrowseButton1LB.Top = 2640
    BrowseButton1LB.Visible = True
    BrowseButton1LB.ZOrder 0
    
    Dim tmp As NTControls22.BrowseButton
    
    Set tmp = BrowseButton1LB
    
    tmp.BorderStyle = 2
    tmp.BrowseAction = 0
    tmp.BrowseTitle = "Backup to File"
    tmp.FileFilter = "Max-FTP Database Backups|*.madb"
    tmp.FileFilterIndex = 1
    
    Set tmp = Nothing
        
    Label2.Tag = Timer
    
    Timer1.Interval = 1
    Timer1.enabled = True

End Sub

Private Sub Text1_Change()
    GetRestoreOptions
End Sub

Private Sub Command1_Click()
    EnableForm False
    
    UtilityRestore ctlDBOptions2.Options, Text1.Text, False, True
    
    EnableForm True
End Sub

Private Sub Command2_Click()
    EnableForm False
    
    If InStr(LCase(Text2.Text), "installer" & MaxDBBackupExt) > 0 Then
        MsgBox "Invalid file name!", vbOKOnly, AppName
    Else
        UtilityBackup ctlDBOptions1.Options, Text2.Text, False
    End If
    
    EnableForm True
End Sub

Private Sub Command3_Click()
    EnableForm False
    
    UtilityReset ctlDBOptions3.Options, False, True
    
    EnableForm True
End Sub

Private Sub ctlDBOptions1_OptionChanged()
    Command2.enabled = (Not ctlDBOptions1.Options = bo_None) And (Not DatabaseOpen)
End Sub

Private Sub ctlDBOptions2_OptionChanged()
    Command1.enabled = (Not ctlDBOptions2.Options = bo_None) And (Not DatabaseOpen)
End Sub

Private Sub ctlDBOptions3_OptionChanged()
    Command3.enabled = (Not ctlDBOptions3.Options = bo_None) And (Not DatabaseOpen)
End Sub

Public Sub ShowForm()
    
    DatabaseOpen = True
    If (Text1.Text = "") Then
        Text1.Text = BrowseButton1LB.GetFolderByAction(5) & "\maxbackup" & MaxDBBackupExt
        Text2.Text = BrowseButton1LB.GetFolderByAction(5) & "\maxbackup" & MaxDBBackupExt
    Else
        Option2.Value = True
    End If

    GetRestoreOptions
    Me.Show
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Timer1.enabled = False
    
    Set BrowseButton1LB = Nothing

    End
End Sub

Private Sub Option1_Click()
    ShowTask
End Sub

Private Sub Option2_Click()
    ShowTask
End Sub

Private Sub Option3_Click()
    ShowTask
End Sub

Private Sub Timer1_Timer()
    If Label2.Caption <> "" Then
        If Timer - Label2.Tag > 0.6 Then
            Label2.Tag = Timer
            Label2.Font.Underline = Not Label2.Font.Underline
        End If
    End If
    
    DatabaseOpen = (ProcessRunning(MaxFileName) > 0) Or (ProcessRunning(ServiceFileName) > 0) Or (ProcessRunning(MaxIDEFileName) > 0)
    
    If (ProcessRunning(ServiceFileName) > 0) Then
        Label2.Caption = "Please close the Max-FTP Scheduole Service to continue."
    ElseIf DatabaseOpen Then
        Label2.Caption = "Please close all Max-FTP applications to continue."
    Else
        Label2.Caption = ""
    End If
    
    Command1.enabled = Not DatabaseOpen
    Command2.enabled = Not DatabaseOpen
    Command3.enabled = Not DatabaseOpen
    
    If Frame1.Visible Then
        BrowseButton1LB.FilterPath = Text2.Text
    ElseIf Frame2.Visible Then
        BrowseButton1LB.FilterPath = Text1.Text
    End If

    BrowseButton1LB.Visible = (Not Frame3.Visible)
    
    If Not BrowseButton1LB.CurrentAction = -2 Then
        If Not BrowseButton1LB.BrowseReturn = "" Then
            If Frame2.Visible Then
                If Not Text1.Tag = BrowseButton1LB.BrowseReturn Then
                    Text1.Tag = BrowseButton1LB.BrowseReturn
                    Text1.Text = BrowseButton1LB.BrowseReturn
                    GetRestoreOptions
                End If
            ElseIf Frame1.Visible Then
                If Not Text2.Tag = BrowseButton1LB.BrowseReturn Then
                    Text2.Tag = BrowseButton1LB.BrowseReturn
                    Text2.Text = BrowseButton1LB.BrowseReturn
                End If
            End If
        End If
    End If

End Sub
