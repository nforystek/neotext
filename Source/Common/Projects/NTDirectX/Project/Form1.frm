VERSION 5.00
Object = "*\A..\..\NTDirectX.vbp"
Begin VB.Form Form1 
   Caption         =   "DirectX Test"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11130
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   11130
   StartUpPosition =   2  'CenterScreen
   Begin NTDirectX.Macroscopic Macroscopic1 
      Height          =   4665
      Left            =   2460
      TabIndex        =   0
      Top             =   1365
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   8229
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Form_Load()
    Macroscopic1.LoadFolder App.Path & "\ToggleBox"

End Sub

Private Sub Form_Resize()
    Macroscopic1.Left = 0
    Macroscopic1.Top = 0
    Macroscopic1.Width = Me.ScaleWidth '- (Me.Width - Me.ScaleWidth)
    Macroscopic1.Height = Me.ScaleHeight '- ScratchKad1.Top '- (Me.Height - Me.ScaleHeight)
    
End Sub
