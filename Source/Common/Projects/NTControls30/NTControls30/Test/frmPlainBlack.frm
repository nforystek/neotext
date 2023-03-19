VERSION 5.00
Object = "*\A..\..\NTControls30.vbp"
Begin VB.Form frmPlainBlack 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin NTControls30.Neotext Neotext1 
      Height          =   2535
      Left            =   810
      TabIndex        =   0
      Top             =   615
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   4471
      LeftMargin      =   0   'False
      FontSize        =   8.25
   End
End
Attribute VB_Name = "frmPlainBlack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    Neotext1.Width = Me.ScaleWidth
    Neotext1.Height = Me.ScaleHeight
    Neotext1.Top = 0
    Neotext1.left = 0
    
    
    Neotext1.FileName = "C:\Development\Neotext\Common\Projects\NTControls30\Test\VBScript.txt"
End Sub
