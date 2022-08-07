
VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "To-Doster Change Control Software"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6840
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Closed"
      Height          =   225
      Left            =   240
      TabIndex        =   4
      Top             =   2235
      Width           =   825
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
      Height          =   330
      Index           =   3
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1830
      Width           =   900
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Add"
      Height          =   330
      Index           =   2
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1110
      Width           =   900
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Edit"
      Height          =   330
      Index           =   1
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1470
      Width           =   900
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5820
      Left            =   1275
      TabIndex        =   0
      Top             =   15
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   10266
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Product"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Comments"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   135
      TabIndex        =   5
      Top             =   6000
      Width           =   945
   End
   Begin VB.Image Image1 
      Height          =   1305
      Left            =   -30
      Picture         =   "frmMain.frx":2CFA
      Top             =   -165
      Width           =   1275
   End
   Begin VB.Menu mnuAction 
      Caption         =   "Action"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserRefreshList()
    If Check1.Value Then
        RefreshList ListView1, "None"
    Else
        RefreshList ListView1, EncryptString("Closed")
    End If

End Sub

Private Sub Check1_Click()
    UserRefreshList
End Sub

Private Sub Command1_Click(Index As Integer)
    
    Select Case Index
        Case 2 'add
            
            Load frmEntry
            
            frmEntry.Caption = "Add Entry"
            
            frmEntry.txtDateTime.Text = Now
            frmEntry.cmbProduct.ListIndex = 0
            frmEntry.cmbType.ListIndex = 0
            frmEntry.cmbStatus.ListIndex = 0
            frmEntry.txtComments.Tag = ""
            frmEntry.cmbProduct.Locked = False
    
            frmEntry.Show 1
            If frmEntry.IsOk Then
                
                Dim newnode
                
                Set newnode = ListView1.ListItems.Add(, , frmEntry.cDateTime)
                newnode.SubItems(1) = frmEntry.cProduct
                newnode.SubItems(2) = frmEntry.cType
                newnode.SubItems(3) = frmEntry.cComments
                newnode.SubItems(4) = frmEntry.cStatus
                
                newnode.Tag = InsertRecord(frmEntry.cDateTime, frmEntry.cProduct, frmEntry.cType, frmEntry.cComments, frmEntry.cStatus)
                
            End If
            Unload frmEntry
            
        Case 1 'edit
            
            If Not ListView1.SelectedItem Is Nothing Then
                
                Load frmEntry
                
                frmEntry.txtDateTime.Text = ListView1.SelectedItem.Text
                frmEntry.cmbProduct.Text = ListView1.SelectedItem.SubItems(1)
                frmEntry.cmbType.ListIndex = IsOnList(frmEntry.cmbType, ListView1.SelectedItem.SubItems(2))
                frmEntry.txtComments.Text = ListView1.SelectedItem.SubItems(3)
                frmEntry.txtComments.Tag = ListView1.SelectedItem.SubItems(3)
                frmEntry.cmbStatus.ListIndex = IsOnList(frmEntry.cmbStatus, ListView1.SelectedItem.SubItems(4))
                frmEntry.cmbProduct.Locked = True
                
   
                frmEntry.Caption = "Edit Entry"
                
                frmEntry.Show 1
                If frmEntry.IsOk Then
                    
                    UpdateRecord CLng(ListView1.SelectedItem.Tag), frmEntry.cDateTime, frmEntry.cProduct, frmEntry.cType, frmEntry.cComments, frmEntry.cStatus
                    
                    ListView1.SelectedItem.Text = frmEntry.cDateTime
                    ListView1.SelectedItem.SubItems(1) = frmEntry.cProduct
                    ListView1.SelectedItem.SubItems(2) = frmEntry.cType
                    ListView1.SelectedItem.SubItems(3) = frmEntry.cComments
                    ListView1.SelectedItem.SubItems(4) = frmEntry.cStatus
                
                End If
                Unload frmEntry
            
            End If
        
        Case 0 'remove
            
            If Not ListView1.SelectedItem Is Nothing Then
                
                If MsgBox("Are you sure you wanna delete the entry for " & ListView1.SelectedItem.Text & "?", vbQuestion + vbYesNo) = vbYes Then
                    
                    DeleteRecord CLng(ListView1.SelectedItem.Tag)
                     
                    ListView1.ListItems.Remove ListView1.SelectedItem.Index
               
                End If
            
            End If
        
        Case 3
            End
        
    End Select

End Sub




Private Sub Form_Load()
    LoadSettings
    
    lblVersion.Caption = "v " & App.Major & "." & App.Minor & "." & App.Revision
    
    If Settings.winTop > 0 Then Me.Top = Settings.winTop
    If Settings.winLeft > 0 Then Me.Left = Settings.winLeft
    If Settings.winHeight > 0 Then Me.Height = Settings.winHeight
    If Settings.winWidth > 0 Then Me.Width = Settings.winWidth
    
    If Settings.col1Width > 0 Then ListView1.ColumnHeaders.Item(1).Width = Settings.col1Width
    If Settings.col2Width > 0 Then ListView1.ColumnHeaders.Item(2).Width = Settings.col2Width
    If Settings.col3Width > 0 Then ListView1.ColumnHeaders.Item(3).Width = Settings.col3Width
    If Settings.col4Width > 0 Then ListView1.ColumnHeaders.Item(4).Width = Settings.col4Width
    If Settings.col5Width > 0 Then ListView1.ColumnHeaders.Item(5).Width = Settings.col5Width

    UserRefreshList

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Settings.winTop = Me.Top
    Settings.winLeft = Me.Left
    Settings.winHeight = Me.Height
    Settings.winWidth = Me.Width
    
    Settings.col1Width = ListView1.ColumnHeaders.Item(1).Width
    Settings.col2Width = ListView1.ColumnHeaders.Item(2).Width
    Settings.col3Width = ListView1.ColumnHeaders.Item(3).Width
    Settings.col4Width = ListView1.ColumnHeaders.Item(4).Width
    Settings.col5Width = ListView1.ColumnHeaders.Item(5).Width
    
    
    SaveSettings

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Width < 5535 Then Me.Width = 5535
    If Me.Height < 3285 Then Me.Height = 3285
    
    ListView1.Top = 0
    ListView1.Width = Me.ScaleWidth - ListView1.Left
    ListView1.Height = Me.ScaleHeight
    
    lblVersion.Left = 0
    lblVersion.Top = Me.ScaleHeight - lblVersion.Height
    lblVersion.Width = ListView1.Left
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    End
End Sub


Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim sKey As Integer
    sKey = ColumnHeader.Index - 1
    
    If sKey = ListView1.SortKey Then
        If ListView1.SortOrder = 0 Then
            ListView1.SortOrder = 1
        Else
            ListView1.SortOrder = 0
        End If
    Else
        ListView1.SortKey = sKey
    End If

End Sub

Private Sub ListView1_DblClick()
    Command1_Click 1
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnuAction
    End If
End Sub

Private Sub mnuAdd_Click()
    Command1_Click 2
End Sub

Private Sub mnuEdit_Click()
    Command1_Click 1
End Sub

Private Sub mnuRefresh_Click()
    UserRefreshList
End Sub

Private Sub mnuRemove_Click()
    Command1_Click 0
End Sub