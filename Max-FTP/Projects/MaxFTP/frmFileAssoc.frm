VERSION 5.00
Object = "{C98B112F-745F-4542-B5B3-DDFADF1F6E2F}#1054.0#0"; "NTControls22.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFileAssoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Associations"
   ClientHeight    =   3840
   ClientLeft      =   3660
   ClientTop       =   615
   ClientWidth     =   8595
   HelpContextID   =   13
   Icon            =   "frmFileAssoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3840
   ScaleWidth      =   8595
   StartUpPosition =   2  'CenterScreen
   Tag             =   "fileassoc"
   Visible         =   0   'False
   Begin VB.PictureBox pFileIcon 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   1155
      ScaleHeight     =   240
      ScaleMode       =   0  'User
      ScaleWidth      =   240
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Close"
      Height          =   345
      HelpContextID   =   13
      Index           =   0
      Left            =   7530
      TabIndex        =   10
      Top             =   3405
      Width           =   945
   End
   Begin VB.Frame Frame1 
      Height          =   3300
      HelpContextID   =   13
      Index           =   0
      Left            =   60
      TabIndex        =   11
      Top             =   0
      Width           =   8430
      Begin VB.CheckBox Check2 
         Caption         =   "Assume Line Feed (non Windows Systems)"
         Height          =   195
         Left            =   4365
         TabIndex        =   6
         Top             =   975
         Width           =   3480
      End
      Begin NTControls22.BrowseButton BrowseButton1 
         Height          =   315
         Left            =   7965
         TabIndex        =   8
         Top             =   1260
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         bBrowseTitle    =   "Browse for Application"
         bBrowseAction   =   0
         bFileFilter     =   "Applications|*.exe;*.bat;*.com|All Files|*.*"
         bFileFilterIndex=   0
         bEnabled        =   0   'False
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2535
         Left            =   120
         TabIndex        =   0
         Top             =   210
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   4471
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File Types"
            Object.Width           =   6174
         EndProperty
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Windows Default"
         Height          =   210
         HelpContextID   =   13
         Left            =   4230
         TabIndex        =   9
         Top             =   1605
         Width           =   1545
      End
      Begin VB.OptionButton Option4 
         Caption         =   "ASCII Text"
         Height          =   270
         HelpContextID   =   13
         Index           =   1
         Left            =   6645
         TabIndex        =   5
         Top             =   675
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Binary Data"
         Height          =   270
         HelpContextID   =   13
         Index           =   0
         Left            =   5250
         TabIndex        =   4
         Top             =   675
         Width           =   1140
      End
      Begin VB.TextBox Text2 
         Height          =   315
         HelpContextID   =   13
         Left            =   4185
         TabIndex        =   7
         Top             =   1245
         Width           =   3690
      End
      Begin VB.TextBox Text1 
         Height          =   300
         HelpContextID   =   13
         Left            =   4500
         TabIndex        =   3
         Top             =   255
         Width           =   3780
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Delete"
         Height          =   315
         HelpContextID   =   13
         Index           =   1
         Left            =   855
         TabIndex        =   2
         Top             =   2835
         Width           =   750
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&New"
         Height          =   315
         HelpContextID   =   13
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   2835
         Width           =   615
      End
      Begin VB.Label Label3 
         Height          =   210
         Left            =   3345
         TabIndex        =   16
         Top             =   2910
         Width           =   2850
      End
      Begin VB.Label Label5 
         Caption         =   "Transfer this file type as:"
         Height          =   210
         Left            =   3330
         TabIndex        =   14
         Top             =   690
         Width           =   1830
      End
      Begin VB.Label Label2 
         Caption         =   "Application:"
         Height          =   255
         Left            =   3330
         TabIndex        =   13
         Top             =   1260
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "File Extentions:"
         Height          =   240
         Left            =   3315
         TabIndex        =   12
         Top             =   285
         Width           =   1140
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   3240
         X2              =   3240
         Y1              =   210
         Y2              =   3170
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   3225
         X2              =   3225
         Y1              =   210
         Y2              =   3170
      End
   End
   Begin MSComctlLib.ImageList imgFiles 
      Left            =   360
      Top             =   3300
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
            Picture         =   "frmFileAssoc.frx":08CA
            Key             =   "folder"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFileAssoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private CancelFileUpdate As Boolean

Private dbConn As clsDBConnection

Private Sub BrowseButton1_ButtonClick(ByVal BrowseReturn As String)
If BrowseReturn <> "" Then Text2.Text = BrowseReturn
End Sub

Private Sub Check1_Click()
    If Not CancelFileUpdate Then
        dbConn.dbQuery "UPDATE FileAssociations SET WindowsApp=" & Check1.Value & " WHERE ID=" & Mid(ListView1.SelectedItem.Key, 5) & ";"
            
        SetWinApp
    End If
End Sub

Private Sub Check2_Click()
    If Not CancelFileUpdate Then
        dbConn.dbQuery "UPDATE FileAssociations SET AssumeLineFeed=" & Check2.Value & " WHERE ID=" & Mid(ListView1.SelectedItem.Key, 5) & ";"

    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    
    Select Case Index
        Case 0
            
            dbConn.rsQuery rs, "INSERT INTO FileAssociations (Locked, DisplayName, Extentions, TransferType, WindowsApp, ApplicationExe, AssumeLineFeed) " & _
                    "VALUES (No,'New Type','*.*',0,Yes,'',Yes);"
            
            dbConn.rsQuery rs, "SELECT * FROM FileAssociations WHERE Locked=No AND DisplayName='New Type' AND Extentions='*.*' AND TransferType=0 AND WindowsApp=Yes AND ApplicationExe='' AND AssumeLineFeed=Yes;"
            
            Dim dbID As Long
            dbID = rs("ID")
            
            RefreshFileAssocList
            
            ListView1.ListItems("FILE" & dbID).EnsureVisible
            ListView1.ListItems("FILE" & dbID).Selected = True
            ListView1.SetFocus
            ListView1.StartLabelEdit
        
        Case 1
            
            dbConn.rsQuery rs, "DELETE FROM FileAssociations WHERE ID=" & Mid(ListView1.SelectedItem.Key, 5) & ";"
        
            RefreshFileAssocList
        
    End Select

    rsClose rs
    
    SetAssociation

End Sub

Private Sub Command4_Click(Index As Integer)
    Unload Me
End Sub
Private Sub SetWinApp()
        
        If Check1.Value = 0 Then
            Text2.enabled = True
            Text2.BackColor = &H80000005
            BrowseButton1.enabled = True
        Else
            Text2.BackColor = &H8000000F
            Text2.enabled = False
            BrowseButton1.enabled = False
        End If

End Sub
Private Sub SetAssociation()
    Dim rs As New ADODB.Recordset
    
    If Not ListView1.SelectedItem Is Nothing Then
    
        CancelFileUpdate = True
            
        dbConn.rsQuery rs, "SELECT * FROM FileAssociations WHERE ID=" & Mid(ListView1.SelectedItem.Key, 5) & ";"
    
        Text1.Text = rs("Extentions") & ""
        Option4(rs("TransferType")).Value = True
            
        Text2.Text = rs("ApplicationExe") & ""
        If rs("WindowsApp") Then
            Check1.Value = 1
        Else
            Check1.Value = 0
        End If
        If rs("AssumeLineFeed") Then
            Check2.Value = 1
        Else
            Check2.Value = 0
        End If
    
        If rs("Locked") Then
            Label3.Caption = "This file association is locked."
            Text1.BackColor = &H8000000F
            Text2.BackColor = &H8000000F
            Label1.enabled = False
            Label2.enabled = False
            Text1.enabled = False
            Text2.enabled = False
            Label5.enabled = False
            Option4(0).enabled = False
            Option4(1).enabled = False
            Check1.enabled = False
            BrowseButton1.enabled = False
            Command1(1).enabled = False
            Check2.enabled = False

        Else
            Label3.Caption = ""
            Label1.enabled = True
            Label2.enabled = True
            Text1.BackColor = &H80000005
            Text1.enabled = True
            Label5.enabled = True
            Option4(0).enabled = True
            Option4(1).enabled = True
            Check1.enabled = True
            Command1(1).enabled = True
            Check2.enabled = Option4(1).Value
            SetWinApp
        End If
    
        rsClose rs
    
        CancelFileUpdate = False
        
    Else
        Label3.Caption = ""
        Label1.enabled = False
        Label2.enabled = False
        Text1.BackColor = &H8000000F
        Text2.BackColor = &H8000000F
        Text1.enabled = False
        Text2.enabled = False
        Label5.enabled = False
        Option4(0).enabled = False
        Option4(1).enabled = False
        Check1.enabled = False
        BrowseButton1.enabled = False
        Command1(1).enabled = False
        Check2.enabled = False
    End If

End Sub

Private Sub Form_Load()

    SetIcecue Line1, "icecue_shadow"
    SetIcecue Line2, "icecue_hilite"
    
    Set dbConn = New clsDBConnection
    
    RefreshFileAssocList
    
    SetAssociation

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ClearFileAssocList
    
    Set dbConn = Nothing
End Sub

Private Sub ClearFileAssocList()

    Dim lstItem

    Set ListView1.SmallIcons = Nothing
    For Each lstItem In ListView1.ListItems
        RemoveAssociation lstItem.SmallIcon
    Next
    ListView1.ListItems.Clear

End Sub

Private Sub RefreshFileAssocList()

    ClearFileAssocList

    Dim rs As New ADODB.Recordset
    
    Dim fileType As String
    Dim lstItem
    ListView1.SmallIcons = frmMain.imgFiles
    dbConn.rsQuery rs, "SELECT * FROM FileAssociations;"
    Do Until rsEnd(rs)
                
        If InStr(rs("Extentions"), ",") > 0 Then
            fileType = Trim(Left(rs("Extentions"), InStr(rs("Extentions"), ",") - 1))
        Else
            fileType = rs("Extentions")
        End If
        fileType = Replace(fileType, "*", "")
        
        GetAssociation fileType, fileType, pFileIcon(1)
        
        LoadAssociation pFileIcon(1)
        Set lstItem = ListView1.ListItems.Add(, "FILE" & rs("ID"), rs("DisplayName"), , pFileIcon(1).Tag)
                
        rs.MoveNext
    Loop

End Sub

Private Sub ListView1_AfterLabelEdit(Cancel As Integer, NewString As String)
    If Not CancelFileUpdate Then
        dbConn.dbQuery "UPDATE FileAssociations SET DisplayName='" & Replace(NewString, "'", "''") & "' WHERE ID=" & Mid(ListView1.SelectedItem.Key, 5) & ";"
    End If
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
    Cancel = Not (Label3.Caption = "")
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    SetAssociation
End Sub

Private Sub Option4_Click(Index As Integer)
    If Not CancelFileUpdate Then
        dbConn.dbQuery "UPDATE FileAssociations SET TransferType=" & Index & " WHERE ID=" & Mid(ListView1.SelectedItem.Key, 5) & ";"
    End If
    Check2.enabled = Option4(1).Value
End Sub

Private Sub Text1_Change()
    If Not CancelFileUpdate Then
        dbConn.dbQuery "UPDATE FileAssociations SET Extentions='" & Replace(Text1.Text, "'", "''") & "' WHERE ID=" & Mid(ListView1.SelectedItem.Key, 5) & ";"
    End If
End Sub

Private Sub Text2_Change()
    If Not CancelFileUpdate Then
        dbConn.dbQuery "UPDATE FileAssociations SET ApplicationExe='" & Replace(Text2.Text, "'", "''") & "' WHERE ID=" & Mid(ListView1.SelectedItem.Key, 5) & ";"
    End If
    BrowseButton1.FilterPath = MapFolderVariables(Text2.Text)
End Sub
