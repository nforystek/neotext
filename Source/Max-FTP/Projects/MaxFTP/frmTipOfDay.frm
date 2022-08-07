VERSION 5.00
Begin VB.Form frmTipOfDay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tip of the Day"
   ClientHeight    =   3492
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5172
   Icon            =   "frmTipOfDay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3492
   ScaleWidth      =   5172
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CheckBox Check1 
      Caption         =   "Show Tips at Startup"
      Height          =   315
      Left            =   195
      TabIndex        =   0
      Top             =   3000
      Width           =   2130
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Next Tip"
      Height          =   360
      Index           =   1
      Left            =   2805
      TabIndex        =   1
      Top             =   2985
      Width           =   1020
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   360
      Index           =   0
      Left            =   3960
      TabIndex        =   2
      Top             =   2985
      Width           =   1020
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2670
      Left            =   120
      ScaleHeight     =   2628
      ScaleWidth      =   4872
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   135
      Width           =   4920
      Begin VB.Label TipText 
         BackColor       =   &H00FFFFFF&
         Height          =   1185
         Left            =   135
         TabIndex        =   4
         Top             =   1305
         Width           =   4560
      End
      Begin VB.Image Image1 
         Height          =   576
         Left            =   108
         Picture         =   "frmTipOfDay.frx":08CA
         Top             =   120
         Width           =   1728
      End
   End
End
Attribute VB_Name = "frmTipOfDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
            TipText.Caption = GetNextTip
    End Select
End Sub

Function GetNextTip() As String
    Dim db As New clsDBConnection
    Dim rs As New ADODB.Recordset
    
    db.rsQuery rs, "SELECT * FROM TipOfDay WHERE Viewed=0;"
    
    If rs.EOF Or rs.BOF Then
        db.rsQuery rs, "UPDATE TipOfDay SET Viewed=0;"
        db.rsQuery rs, "SELECT * FROM TipOfDay WHERE Viewed=0;"
    End If
    
    If Not rs.EOF And Not rs.BOF Then
        GetNextTip = rs("TipText")
        
        db.rsQuery rs, "UPDATE TipOfDay SET Viewed=-1 WHERE TipText = '" & Replace(rs("TipText"), "'", "''") & "';"
        
    Else
        GetNextTip = "No Tips in database!!!!!  FIX IT!"
    End If
    
    If rs.State <> 0 Then rs.Close
    Set rs = Nothing
End Function

Private Sub Form_Load()
    Check1.Value = BoolToCheck(dbSettings.GetProfileSetting("TipOfDay"))
    
    TipText.Caption = GetNextTip()
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    dbSettings.SetProfileSetting "TipOfDay", Check1.Value
End Sub
