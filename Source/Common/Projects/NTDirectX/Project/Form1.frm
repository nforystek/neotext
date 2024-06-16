VERSION 5.00
Object = "{6459C47B-7678-440A-8976-7FEB2C548409}#47.0#0"; "NTDirectX.ocx"
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
      Height          =   2550
      Left            =   4245
      TabIndex        =   0
      Top             =   1200
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   4498
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
