VERSION 5.00
Object = "{5167AC20-C665-4BC1-B458-B10062ABCDC5}#217.0#0"; "NTControls30.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5505
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   5505
   Begin NTControls30.Neotext Neotext1 
      Height          =   1050
      Left            =   315
      TabIndex        =   0
      Top             =   675
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   1852
      FontSize        =   7.5
      ForeColor       =   -2147483643
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Resize()
    Neotext1.Top = 0
    Neotext1.Left = 0
    Neotext1.Width = Me.ScaleWidth
    Neotext1.Height = Me.ScaleHeight
End Sub
