


VERSION 5.00
Begin VB.Form frmEquipment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Password"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   Icon            =   "frmEquipment.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   5580
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1065
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1440
      Width           =   2595
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4110
      TabIndex        =   4
      Top             =   1050
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   4110
      TabIndex        =   3
      Top             =   525
      Width           =   1185
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1065
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1095
      Width           =   2595
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1065
      TabIndex        =   0
      Top             =   555
      Width           =   2595
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Confirm:"
      Height          =   180
      Index           =   1
      Left            =   165
      TabIndex        =   8
      Top             =   1485
      Width           =   840
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Password:"
      Height          =   180
      Index           =   0
      Left            =   165
      TabIndex        =   7
      Top             =   1140
      Width           =   840
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Username:"
      Height          =   210
      Left            =   165
      TabIndex        =   6
      Top             =   585
      Width           =   840
   End
   Begin VB.Label Label1 
      Height          =   270
      Left            =   225
      TabIndex        =   5
      Top             =   135
      Width           =   4665
   End
End
Attribute VB_Name = "frmEquipment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public PassType As Integer
Public IsOk As Boolean
Public Username As String
Public Password As String

Public Sub ShowShareBox(ByVal pCaption As String, ByVal pUsername As String, ByVal pPassword As String)
    PassType = 1
    Label1.Caption = pCaption
    If pUsername = "" Then pUsername = ".\" & dbSettings.GetUserLoginName
    Username = pUsername
    Password = pPassword
    
    Text2(1).Visible = False
    Label3(1).Visible = False
    
    Me.Height = 1980
    
    Text1.Text = Username
    Text2(0).Text = Password
    
    Me.Show
    
    If Not (Text1.Text = "") Then Text2(0).SetFocus
End Sub

Public Sub ShowConfirmBox(ByVal pCaption As String, ByVal pUsername As String, ByVal pPassword As String)
    PassType = 2
    Label1.Caption = pCaption
    Username = pUsername
    Password = pPassword
    
    Text2(1).Visible = True
    Label3(1).Visible = True
    
    Me.Height = 2235

    Text1.Text = Username
    Text2(0).Text = Password

    Me.Show

End Sub

Private Sub Command1_Click(Index As Integer)
    
    If Index = 0 Then
        
        If PassType = 2 Then
            If Trim(Text2(0).Text) <> Trim(Text2(1).Text) Then
                MsgBox "You must enter the password twice.  Try again, or click cancel.", vbInformation, AppName
                Exit Sub
            End If
        
        End If
    
    End If
    
    Username = Text1.Text
    Password = Text2(0).Text
    Select Case Index
        Case 0
            IsOk = True
        Case 1
            IsOk = False
    End Select
    Me.Hide
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = True
        IsOk = False
        Me.Hide
    End If
End Sub