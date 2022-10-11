VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRestrict 
   Caption         =   "Restrictions"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9600
   Icon            =   "frmRestrict.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5880
      Top             =   1680
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Close"
      Height          =   330
      Left            =   3075
      TabIndex        =   3
      Top             =   855
      Width           =   1080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Remove"
      Height          =   330
      Left            =   1905
      TabIndex        =   2
      Top             =   840
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add Project"
      Height          =   330
      Left            =   720
      TabIndex        =   1
      Top             =   870
      Width           =   1080
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4170
   End
End
Attribute VB_Name = "frmRestrict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Private Sub Command1_Click()
    On Error GoTo cancelit
    
    CommonDialog1.ShowOpen
    
    If PathExists(CommonDialog1.FileName, True) Then
        List1.AddItem CommonDialog1.FileName
        Dim txt As String
        Dim cnt As Long
        For cnt = 0 To List1.ListCount - 1
            txt = txt & List1.List(cnt) & vbCrLf
        Next
        SaveSetting "BasicNeotext", "Options", "RestrictList", txt
    Else
        MsgBox "File not found", "Restrictions", vbOKOnly
    
    End If
    
    Exit Sub
cancelit:
    Err.Clear
End Sub

Private Sub Command2_Click()
    Dim cnt As Long
    If List1.ListCount > 0 Then
        Dim txt As String
        Do While cnt <= List1.ListCount - 1
            If List1.Selected(cnt) Then
                List1.RemoveItem cnt
            Else
                txt = txt & List1.List(cnt) & vbCrLf
                cnt = cnt + 1
            End If
        Loop
        SaveSetting "BasicNeotext", "Options", "RestrictList", txt
        
    End If
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim txt As String
    txt = GetSetting("BasicNeotext", "Options", "RestrictList", "")
    Do Until txt = ""
        List1.AddItem RemoveNextArg(txt, vbCrLf)
    Loop
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    List1.Top = (Screen.TwipsPerPixelY * 4)
    List1.Left = (Screen.TwipsPerPixelX * 4)
    List1.Width = Me.ScaleWidth - (Screen.TwipsPerPixelX * 8)
    List1.Height = Me.ScaleHeight - (Screen.TwipsPerPixelX * 12) - Command3.Height
    Command3.Left = Me.ScaleWidth - (Screen.TwipsPerPixelX * 4) - Command3.Width
    Command2.Left = Command3.Left - (Screen.TwipsPerPixelX * 4) - Command2.Width
    Command1.Left = Command2.Left - (Screen.TwipsPerPixelX * 4) - Command1.Width
    Command1.Top = List1.Top + List1.Height + (Screen.TwipsPerPixelY * 4)
    Command2.Top = List1.Top + List1.Height + (Screen.TwipsPerPixelY * 4)
    Command3.Top = List1.Top + List1.Height + (Screen.TwipsPerPixelY * 4)
    If Me.Height < ((Screen.TwipsPerPixelY * 4) * 4) + List1.Height + Command1.Height Then
        Me.Height = ((Screen.TwipsPerPixelY * 4) * 4) + List1.Height + Command1.Height
    End If
    If Me.Width < (((Screen.TwipsPerPixelX * 4) * 5) + Command1.Width + Command2.Width + Command3.Width) Then
        Me.Width = (((Screen.TwipsPerPixelX * 4) * 5) + Command1.Width + Command2.Width + Command3.Width)
    End If
    If Err Then Err.Clear
    On Error GoTo 0
End Sub

