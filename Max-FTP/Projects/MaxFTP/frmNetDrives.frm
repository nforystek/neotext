VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNetDrives 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Network Drive Connections"
   ClientHeight    =   4620
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5820
   HelpContextID   =   14
   Icon            =   "frmNetDrives.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   5820
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1800
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4170
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   345
      HelpContextID   =   14
      Index           =   0
      Left            =   4545
      TabIndex        =   4
      Top             =   4170
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Height          =   4050
      HelpContextID   =   14
      Left            =   105
      TabIndex        =   5
      Top             =   15
      Width           =   5610
      Begin VB.CommandButton Command1 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   360
         HelpContextID   =   14
         Index           =   1
         Left            =   4575
         TabIndex        =   3
         Top             =   2835
         Width           =   840
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Add"
         Height          =   360
         HelpContextID   =   14
         Index           =   0
         Left            =   4575
         TabIndex        =   2
         Top             =   2400
         Width           =   840
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1485
         HelpContextID   =   14
         Index           =   0
         Left            =   150
         TabIndex        =   0
         Top             =   375
         Width           =   5295
         _ExtentX        =   9335
         _ExtentY        =   2625
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Drive Letter"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Share Name"
            Object.Width           =   6174
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1515
         HelpContextID   =   14
         Index           =   1
         Left            =   150
         TabIndex        =   1
         Top             =   2385
         Width           =   4260
         _ExtentX        =   7514
         _ExtentY        =   2667
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Drive Letter"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Share Name"
            Object.Width           =   4305
         EndProperty
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   240
         X2              =   5310
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   240
         X2              =   5310
         Y1              =   2025
         Y2              =   2025
      End
      Begin VB.Label Label1 
         Caption         =   "Max-FTP Session Connections:"
         Height          =   225
         Index           =   1
         Left            =   225
         TabIndex        =   7
         Top             =   2160
         Width           =   2490
      End
      Begin VB.Label Label1 
         Caption         =   "Current Network Connections:"
         Height          =   180
         Index           =   0
         Left            =   225
         TabIndex        =   6
         Top             =   165
         Width           =   2490
      End
   End
   Begin MSComctlLib.ImageList mIcons 
      Left            =   0
      Top             =   3915
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNetDrives.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNetDrives.frx":151E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNetDrives.frx":2172
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNetDrives.frx":2DC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNetDrives.frx":3A1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNetDrives.frx":466E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNetDrives.frx":52C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNetDrives.frx":5F16
            Key             =   "MaxUpLevel"
            Object.Tag             =   "MaxUpLevel"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNetDrives.frx":64B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNetDrives.frx":6A4E
            Key             =   "MaxClosedFolder"
            Object.Tag             =   "MaxClosedFolder"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmNetDrives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Function DriveOnList(driveList As Control, tDrive As String) As Integer
On Error GoTo catch

    Dim cnt As Integer
    Dim found As Boolean
    cnt = 0
    found = False
    Do Until found Or cnt >= driveList.ListCount
        If UCase(Left(driveList.List(cnt), 1)) = tDrive Then
            found = True
        Else
            cnt = cnt + 1
            End If
        Loop
    If found Then
        DriveOnList = cnt
    Else
        DriveOnList = -1
    End If

Exit Function
catch:
    Err.Raise Err.Number, App.EXEName, Err.Description
End Function
Private Sub SetupDriveConnectionList(toList As Control, dList As Control)
On Error GoTo catch

    Dim lDrive As Integer
    Dim foundDrive As Integer
    foundDrive = False
    
    Const tFirstDrive = 68
    Const tLastDrive = 90
    
    lDrive = tFirstDrive
    Do Until lDrive > tLastDrive
        foundDrive = DriveOnList(dList, Chr(lDrive))
        If foundDrive > -1 Then
            If InStr(dList.List(foundDrive), "\\") > 0 Then
                toList.AddItem Chr(lDrive) + Mid(dList.List(foundDrive), 2)
            End If
        Else
            toList.AddItem Chr(lDrive) + ":"
        End If
        lDrive = lDrive + 1
    Loop


Exit Sub
catch:
    Err.Raise Err.Number, App.EXEName, Err.Description
End Sub

Private Sub AddDriveToDatabase(ByVal DriveLetter As String, ByVal ShareName As String)

    Dim dbConn As New clsDBConnection
    
    dbConn.dbQuery "INSERT INTO SessionDrives (ParentID, DriveLetter, ShareName) VALUES (" & dbSettings.CurrentUserID & ",'" & Replace(DriveLetter, "'", "''") & "','" & Replace(ShareName, "'", "''") & "');"

    Set dbConn = Nothing

End Sub

Private Sub RemoveDriveFromDatabase(ByVal DriveLetter As String)

    Dim dbConn As New clsDBConnection

    dbConn.dbQuery "DELETE FROM SessionDrives WHERE DriveLetter='" & Replace(DriveLetter, "'", "''") & "' AND ParentID=" & dbSettings.CurrentUserID & ";"

    Set dbConn = Nothing

End Sub

Sub SetWindowsNetList()

    Drive1.Refresh
    ListView1(0).ListItems.Clear
    Set ListView1(0).SmallIcons = mIcons
    Dim cnt As Integer
    Dim nodX As ListItem
    For cnt = 0 To Drive1.ListCount
        If InStr(Drive1.List(cnt), "\\") > 0 Then
            Set nodX = ListView1(0).ListItems.Add(, , Left(UCase(Drive1.List(cnt)), 2), , 9)
            nodX.SubItems(1) = Trim(Mid(Drive1.List(cnt), 3))
        End If
    Next

End Sub

Private Sub Command1_Click(Index As Integer)

    Select Case Index
        Case 0
        
            SetupDriveConnectionList frmNetConnection.Combo1, Drive1
            frmNetConnection.Show
            Do
                modCommon.DoTasks
            Loop Until frmNetConnection.Visible = False
            
            If frmNetConnection.IsOk Then
                Dim nodX As ListItem
                Set nodX = ListView1(1).ListItems.Add(, , Left(frmNetConnection.Combo1.Text, 2), , 9)
                nodX.SubItems(1) = Trim(frmNetConnection.Text1.Text)
            
                AddDriveToDatabase frmNetConnection.DriveLetter, frmNetConnection.ShareName
                
                Dim dbNet As New clsNetwork
                dbNet.ConnectNetworkDrive frmNetConnection.DriveLetter, frmNetConnection.ShareName, , , True
                Set dbNet = Nothing
    
            End If
            
            SetWindowsNetList
            RefreshDatabaseDriveList
            
            Unload frmNetConnection
            
        Case 1
            If Not ListView1(1).SelectedItem Is Nothing Then
                If MsgBox("Are you sure you want to revome drive " + ListView1(1).SelectedItem.Text, vbQuestion + vbYesNo, AppName) = vbYes Then
                    
                    RemoveDriveFromDatabase ListView1(1).SelectedItem.Text
                
                    Dim dbNet2 As New clsNetwork
                    dbNet2.CancelConnection ListView1(1).SelectedItem.Text
                    Set dbNet2 = Nothing
                    SetWindowsNetList
                    
                    RefreshDatabaseDriveList
    
                End If
            Else
                MsgBox "You must select a drive from the Max Session List to remove it.", vbInformation, AppName
            End If
    End Select
    Unload frmNetConnection

End Sub

Private Sub Command2_Click(Index As Integer)

    Unload Me
    
End Sub

Private Sub Form_Load()
    SetIcecue Line1(0), "icecue_shadow"
    SetIcecue Line1(1), "icecue_hilite"
    
    Dim dbNet As New clsNetwork
    dbNet.OpenSessionDrives
    Set dbNet = Nothing
                
    SetWindowsNetList
    
    RefreshDatabaseDriveList

End Sub
Private Sub RefreshDatabaseDriveList()
    
    ListView1(1).ListItems.Clear
    Set ListView1(1).SmallIcons = mIcons
    
    Dim nodX As ListItem
    
    Dim dbConn As New clsDBConnection
    Dim rs As New ADODB.Recordset
    
    dbConn.rsQuery rs, "SELECT * FROM SessionDrives WHERE ParentID=" & dbSettings.CurrentUserID & ";"
    Do Until rsEnd(rs)
        Set nodX = ListView1(1).ListItems.Add(, , rs("DriveLetter"), , 9)
        nodX.SubItems(1) = rs("ShareName")
        rs.MoveNext
    Loop
    
    rsClose rs
    Set dbConn = Nothing
    
    If ListView1(1).ListItems.Count > 0 Then ListView1(1).ListItems.Item(1).Selected = True
    
    Command1(1).enabled = Not (ListView1(1).SelectedItem Is Nothing)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim sForm As Form
    For Each sForm In Forms
        If sForm.Tag = "ftp" Then
            sForm.pViewDrives(0).Refresh
            sForm.pViewDrives(1).Refresh
        End If
    Next

End Sub

Private Sub ListView1_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)

    If Index = 1 Then Command1(1).enabled = True

End Sub
