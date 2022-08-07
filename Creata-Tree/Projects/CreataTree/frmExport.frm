VERSION 5.00
Object = "{C98B112F-745F-4542-B5B3-DDFADF1F6E2F}#1036.0#0"; "NTControls22.ocx"
Begin VB.Form frmExport 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export Tree"
   ClientHeight    =   1968
   ClientLeft      =   6756
   ClientTop       =   3528
   ClientWidth     =   5940
   Icon            =   "frmExport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1968
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      Height          =   345
      Index           =   1
      Left            =   4650
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1485
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Export"
      Height          =   345
      Index           =   0
      Left            =   4650
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1095
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Options"
      Height          =   930
      Left            =   75
      TabIndex        =   1
      Top             =   975
      Width           =   4500
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Include tree code in html file."
         Height          =   240
         Left            =   135
         TabIndex        =   3
         Top             =   540
         Width           =   2715
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Include custimization in html file."
         Height          =   240
         Left            =   135
         TabIndex        =   2
         Top             =   255
         Width           =   2715
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Folder"
      Height          =   930
      Left            =   75
      TabIndex        =   6
      Top             =   30
      Width           =   5805
      Begin NTControls22.BrowseButton BrowseButton1 
         Height          =   252
         Left            =   5436
         TabIndex        =   8
         Top             =   264
         Width           =   252
         _ExtentX        =   445
         _ExtentY        =   445
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   255
         Width           =   5250
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   5235
      End
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private dirty As Boolean
Private nBase As clsItem

Public Function ShowForm(ByRef tBase As clsItem)
    dirty = False
    Set nBase = tBase
    If Right(ExportFolder, 1) = "\" Then
        Text1.Text = Left(ExportFolder, Len(ExportFolder) - 1)
    ElseIf ExportFolder = "" Then
        Text1.Text = AppPath & "\My Trees\Export\"
    Else
        Text1.Text = ExportFolder
    End If
    
    Label1.Caption = "Folder '\" & nBase.Value("Label") & "' will be created under the above folder."
    
    Check1.Value = -CInt(Ini.Setting("IncludeCustom"))
    Check2.Value = -CInt(Ini.Setting("IncludeCode"))
    
End Function

Private Sub BrowseButton1_ButtonClick(ByVal BrowseReturn As String)
    If Not BrowseReturn = vbNullString Then
        Text1.Text = BrowseReturn
        If Not Right(ExportFolder, 1) = "\" Then ExportFolder = BrowseReturn & "\"
        dirty = True
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0

            If Not (Ini.Setting("IncludeCustom") = CBool(Check1.Value)) Then
                Ini.Setting("IncludeCustom") = CBool(Check1.Value)
                dirty = True
            End If
            If Not (Ini.Setting("IncludeCode") = CBool(Check2.Value)) Then
                Ini.Setting("IncludeCode") = CBool(Check2.Value)
                dirty = True
            End If
            
            Me.Hide
            ExportFolder = Text1.Text
            
            If Not Right(ExportFolder, 1) = "\" Then ExportFolder = ExportFolder & "\"
        
            If ExportHTMLTree(nBase, ExportFolder & nBase.Value("Label") & "\", Check1.Value, Check2.Value) Then
                If dirty Then
                    
                    frmMain.nForm.Changed = True
                    frmMain.RefreshForm
                End If
                Unload Me
            Else
                Me.Show
            End If
            
        Case 1
            Unload Me
    End Select
End Sub

