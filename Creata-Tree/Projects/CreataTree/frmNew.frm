VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmNew 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New"
   ClientHeight    =   3705
   ClientLeft      =   6795
   ClientTop       =   2550
   ClientWidth     =   3945
   Icon            =   "frmNew.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   315
      Index           =   1
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3300
      Width           =   840
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Index           =   0
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3300
      Width           =   840
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3270
      Left            =   90
      TabIndex        =   4
      Top             =   -30
      Width           =   3780
      Begin VB.FileListBox File1 
         Height          =   285
         Left            =   2415
         TabIndex        =   5
         Top             =   1470
         Visible         =   0   'False
         Width           =   795
      End
      Begin CreataTree.ctlNav ctlNav1 
         Height          =   1275
         Left            =   150
         TabIndex        =   1
         Top             =   1830
         Width           =   3480
         _ExtentX        =   6112
         _ExtentY        =   2487
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1425
         Left            =   150
         TabIndex        =   0
         Top             =   225
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   2514
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   5997
         EndProperty
      End
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public IsOk As Boolean
Private nTree As clsItem

Public Property Get Template() As String
    Select Case Me.Caption
        Case "New Tree"
            Template = ListView1.SelectedItem.Text
        Case "New Item"
            Template = ListView1.SelectedItem.Tag
    End Select
End Property

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            IsOk = True
            Me.Hide
        Case 1
            IsOk = False
            Me.Hide
    End Select
End Sub
Public Function SetupTree() As Long
    Me.Caption = "New Tree"
    LoadFiles TreeFileExt
    If ListView1.ListItems.Count > 0 Then ListView1.ListItems(1).Selected = True
    SetupTree = ListView1.ListItems.Count
End Function
Public Function SetupItem(ByRef nBase As clsItem) As Long
    Me.Caption = "New Item"
    
    Set nTree = nBase
    
    Dim nNode As ListItem
    Dim nItem As clsItem
    For Each nItem In nBase.SubItems
        If nItem.IsTemplate Then
            Set nNode = ListView1.ListItems.Add(, , nItem.Value("Label"))
            nNode.Tag = nItem.Key
            Set nNode = Nothing
        End If
    Next
    If ListView1.ListItems.Count > 0 Then ListView1.ListItems(1).Selected = True
    SetupItem = ListView1.ListItems.Count
End Function
Private Function LoadFiles(ByVal Ext As String)
    File1.Pattern = "*" & Ext
    File1.Path = TemplateFolder
    If File1.ListCount > 0 Then
        Dim cnt As Long
        For cnt = 0 To File1.ListCount - 1
            ListView1.ListItems.Add , , Replace(File1.List(cnt), Ext, vbNullString, , , vbTextCompare)
        Next
    End If
End Function

Private Sub Form_Activate()
    ListView1_Click
End Sub

Private Sub Form_Load()

        ctlNav1.Visible = True
        Frame1.Height = 3270
        Command1(0).Top = 3300
        Command1(1).Top = 3300
        Me.Height = 4080

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        IsOk = False
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set nTree = Nothing
End Sub

Private Sub ListView1_Click()
    If Not ListView1.SelectedItem Is Nothing Then
        Dim nBase As New clsItem
        Select Case Me.Caption
            Case "New Tree"
                OpenXMLFile nBase, TemplateFolder & Template & TreeFileExt
                ctlNav1.Refresh nBase, True, True
            Case "New Item"
                nBase.XMLText = nTree.SubItem(Template).XMLText
                ctlNav1.Refresh nBase, True, True
        End Select
        Set nBase = Nothing
    End If
    
End Sub

Private Sub ListView1_DblClick()
    IsOk = True
    Me.Hide
End Sub
