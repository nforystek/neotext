VERSION 5.00
Begin VB.Form frmMIDIProp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MIDI Properties"
   ClientHeight    =   1356
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3228
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1356
   ScaleWidth      =   3228
   StartUpPosition =   1  'CenterOwner
   Begin VB.VScrollBar VScroll1 
      Height          =   315
      Left            =   1320
      TabIndex        =   7
      Top             =   540
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Preview"
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   315
      Index           =   1
      Left            =   2040
      TabIndex        =   5
      Top             =   540
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Index           =   0
      Left            =   2040
      TabIndex        =   4
      Top             =   60
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   780
      TabIndex        =   3
      Text            =   "1"
      Top             =   540
      Width           =   555
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   780
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   810
   End
   Begin VB.Label Label2 
      Caption         =   "Note"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Channel"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "frmMIDIProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public IsOK As Boolean

Private Const MaxNote = 100

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0 'ok
            If IsNumeric(Text1.Text) Then
                If CInt(Text1.Text) > 1 And CInt(Text1.Text) < MaxNote Then
                    IsOK = True
                    Me.Hide
                Else
                    MsgBox "Must enter a number from 1 to " & MaxNote, vbInformation, "MIDI Properties"
                End If
            Else
                MsgBox "Must enter a number from 1 to " & MaxNote, vbInformation, "MIDI Properties"
            End If
        Case 1 'cancel
            IsOK = False
            Me.Hide
    End Select
End Sub

Private Sub Command2_Click()
    If IsNumeric(Text1.Text) Then
        If CInt(Text1.Text) >= 1 And CInt(Text1.Text) <= MaxNote Then
            StartNote Combo1.ListIndex, Text1.Text, MaxNote
        Else
            MsgBox "Must enter a number from 1 to " & MaxNote, vbInformation, "MIDI Properties"
        End If
    Else
        MsgBox "Must enter a number from 1 to " & MaxNote, vbInformation, "MIDI Properties"
    End If
    
End Sub

Private Sub Form_Load()
    Combo1.AddItem "1"
    Combo1.AddItem "2"
    Combo1.AddItem "3"
    Combo1.AddItem "4"
    Combo1.AddItem "5"
    Combo1.AddItem "6"
    Combo1.AddItem "7"
    Combo1.AddItem "8"
    Combo1.AddItem "9"
    Combo1.AddItem "10"
    Combo1.AddItem "11"
    Combo1.AddItem "12"
    Combo1.AddItem "13"
    Combo1.AddItem "14"
    Combo1.AddItem "15"
    Combo1.AddItem "16"
    VScroll1.Max = 1
    VScroll1.Min = MaxNote
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 1 Then
        IsOK = False
        Cancel = False
        Me.Hide
    End If
End Sub

Private Sub Text1_Change()
    If IsNumeric(Text1.Text) Then
        If CInt(Text1.Text) > -1 And CInt(Text1.Text) < -MaxNote Then
            VScroll1.Value = Text1.Text
        End If
    End If
End Sub

Private Sub VScroll1_Change()
    Text1.Text = VScroll1.Value
End Sub

