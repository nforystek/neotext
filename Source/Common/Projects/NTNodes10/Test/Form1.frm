VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Line handler test"
   ClientHeight    =   5730
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   13755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   382
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   917
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   5565
      Left            =   6525
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   40
      Top             =   90
      Width           =   7080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   252
      Index           =   19
      Left            =   2040
      TabIndex        =   37
      Top             =   1080
      Width           =   252
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      Height          =   252
      Index           =   18
      Left            =   2760
      TabIndex        =   36
      Top             =   1080
      Width           =   252
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Item(#)"
      Height          =   252
      Index           =   17
      Left            =   120
      TabIndex        =   34
      Top             =   1080
      Width           =   732
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Item(#)"
      Height          =   252
      Index           =   16
      Left            =   1080
      TabIndex        =   33
      Top             =   1080
      Width           =   852
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Item"
      Height          =   252
      Index           =   15
      Left            =   1080
      TabIndex        =   29
      Top             =   840
      Width           =   852
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Last"
      Height          =   252
      Index           =   13
      Left            =   1080
      TabIndex        =   27
      Top             =   600
      Width           =   852
   End
   Begin VB.CommandButton Command2 
      Caption         =   "First"
      Height          =   252
      Index           =   11
      Left            =   1080
      TabIndex        =   25
      Top             =   360
      Width           =   852
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Value"
      Height          =   252
      Index           =   9
      Left            =   1080
      TabIndex        =   22
      Top             =   120
      Width           =   852
   End
   Begin VB.OptionButton Option8 
      Caption         =   "Constant"
      Height          =   132
      Left            =   5400
      TabIndex        =   20
      Top             =   720
      Width           =   972
   End
   Begin VB.OptionButton Option7 
      Caption         =   "Manual"
      Height          =   132
      Left            =   5400
      TabIndex        =   19
      Top             =   480
      Value           =   -1  'True
      Width           =   852
   End
   Begin VB.Frame Frame1 
      Height          =   852
      Left            =   3120
      TabIndex        =   15
      Top             =   0
      Width           =   1092
      Begin VB.OptionButton Option3 
         Caption         =   "Global"
         Height          =   192
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   852
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Local"
         Height          =   192
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   852
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Normal"
         Height          =   192
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Value           =   -1  'True
         Width           =   852
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pre"
      Height          =   252
      Index           =   7
      Left            =   2520
      TabIndex        =   14
      Top             =   840
      Width           =   492
   End
   Begin VB.CommandButton Command2 
      Caption         =   "App"
      Height          =   252
      Index           =   2
      Left            =   2520
      TabIndex        =   9
      Top             =   600
      Width           =   492
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ins"
      Height          =   252
      Index           =   1
      Left            =   2520
      TabIndex        =   8
      Top             =   360
      Width           =   492
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   252
      Index           =   0
      Left            =   2520
      TabIndex        =   7
      Top             =   120
      Width           =   492
   End
   Begin VB.TextBox Text1 
      Height          =   4212
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   1440
      Width           =   2892
   End
   Begin VB.Frame Frame2 
      Height          =   852
      Left            =   4320
      TabIndex        =   1
      Top             =   0
      Width           =   972
      Begin VB.OptionButton Option6 
         Caption         =   "None"
         Height          =   192
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Value           =   -1  'True
         Width           =   732
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Path"
         Height          =   192
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   732
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Temp"
         Height          =   192
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   732
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test"
      Height          =   252
      Left            =   5400
      TabIndex        =   0
      Top             =   120
      Width           =   972
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   4692
      Left            =   3120
      ScaleHeight     =   309
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   213
      TabIndex        =   6
      Top             =   960
      Width           =   3252
   End
   Begin VB.CommandButton Command2 
      Caption         =   "For"
      Height          =   252
      Index           =   6
      Left            =   2040
      TabIndex        =   13
      Top             =   840
      Width           =   492
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Rem"
      Height          =   252
      Index           =   5
      Left            =   2040
      TabIndex        =   12
      Top             =   600
      Width           =   492
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Del"
      Height          =   252
      Index           =   4
      Left            =   2040
      TabIndex        =   11
      Top             =   360
      Width           =   492
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pop"
      Height          =   252
      Index           =   3
      Left            =   2040
      TabIndex        =   10
      Top             =   120
      Width           =   492
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Item"
      Height          =   252
      Index           =   14
      Left            =   120
      TabIndex        =   28
      Top             =   840
      Width           =   732
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Last"
      Height          =   252
      Index           =   12
      Left            =   120
      TabIndex        =   26
      Top             =   600
      Width           =   732
   End
   Begin VB.CommandButton Command2 
      Caption         =   "First"
      Height          =   252
      Index           =   10
      Left            =   120
      TabIndex        =   24
      Top             =   360
      Width           =   732
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Value"
      Height          =   252
      Index           =   8
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   732
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "="
      Height          =   252
      Index           =   6
      Left            =   2520
      TabIndex        =   39
      Top             =   1080
      Width           =   252
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "="
      Height          =   252
      Index           =   5
      Left            =   2280
      TabIndex        =   38
      Top             =   1080
      Width           =   252
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "="
      Height          =   252
      Index           =   4
      Left            =   840
      TabIndex        =   35
      Top             =   1080
      Width           =   252
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "="
      Height          =   252
      Index           =   3
      Left            =   840
      TabIndex        =   32
      Top             =   840
      Width           =   252
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "="
      Height          =   252
      Index           =   2
      Left            =   840
      TabIndex        =   31
      Top             =   600
      Width           =   252
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "="
      Height          =   252
      Index           =   1
      Left            =   840
      TabIndex        =   30
      Top             =   360
      Width           =   252
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "="
      Height          =   252
      Index           =   0
      Left            =   840
      TabIndex        =   23
      Top             =   120
      Width           =   252
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Private Sub Enable(ByVal IsEnabled As Boolean)
    Command2(0).Enabled = IsEnabled
    Command2(1).Enabled = IsEnabled
    Command2(2).Enabled = IsEnabled
    Command2(3).Enabled = IsEnabled
    Command2(4).Enabled = IsEnabled
    Command2(5).Enabled = IsEnabled
    Command2(6).Enabled = IsEnabled
    Command2(7).Enabled = IsEnabled
    Command2(8).Enabled = IsEnabled
    Command2(9).Enabled = IsEnabled
    Command2(10).Enabled = IsEnabled
    Command2(11).Enabled = IsEnabled
    Command2(12).Enabled = IsEnabled
    Command2(13).Enabled = IsEnabled
    Command2(14).Enabled = IsEnabled
    Command2(15).Enabled = IsEnabled
    Command2(16).Enabled = IsEnabled
    Command2(17).Enabled = IsEnabled
    Command2(18).Enabled = IsEnabled
    Command2(19).Enabled = IsEnabled
    Frame1.Enabled = IsEnabled
    Frame2.Enabled = IsEnabled
    Option1.Enabled = IsEnabled
    Option2.Enabled = IsEnabled
    Option3.Enabled = IsEnabled
    Option4.Enabled = IsEnabled
    Option5.Enabled = IsEnabled
    Option6.Enabled = IsEnabled
    Option7.Enabled = IsEnabled
    Option8.Enabled = IsEnabled
End Sub
Private Sub Command1_Click()

    Dim Action As Long
    Dim Acted As Boolean

    If Command1.Caption = "Test" Then
        Command1.Caption = "Stop"
        
        Enable False
        If Option8.Value Then
            ResetObject
            Do While (Form1.Command1.Caption = "Stop") Or Acted
                Acted = RandomTest()
                If Acted Then
                    Acted = False
                    DebugTest
                    
                End If
            Loop
            DebugTest
        ElseIf Option7.Value Then
            Do While (Not Acted)
                Acted = RandomTest()
            Loop
            DebugTest
        End If
        
        Command1.Caption = "Test"
        Enable True
    Else
        Command1.Caption = "Test"
    End If
End Sub

Private Sub Command2_Click(Index As Integer)
    Enable False
    Form1.Label1(5).Tag = False
    Do Until PreformAction(CLng(Index)) Or Form1.Label1(5).Tag = True
        DoEvents
    Loop
    Enable True
    DebugTest
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Command1.Caption = "Test"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set test = Nothing
    End
End Sub

Private Sub Option1_Click()
    If Option1.Enabled Then ResetObject
    DebugTest
End Sub

Private Sub Option2_Click()
    If Option2.Enabled Then ResetObject
    DebugTest
End Sub

Private Sub Option3_Click()
    If Option3.Enabled Then ResetObject
    DebugTest
End Sub

Private Sub Option4_Click()
    If Option4.Enabled Then ResetObject
    DebugTest
End Sub

Private Sub Option5_Click()
    If Option5.Enabled Then ResetObject
    DebugTest
End Sub

Private Sub Option6_Click()
    If Option6.Enabled Then ResetObject
    DebugTest
End Sub
