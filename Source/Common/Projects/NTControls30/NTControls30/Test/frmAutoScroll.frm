VERSION 5.00
Object = "*\A..\..\NTControls30.vbp"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Rich Scroll Example"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8445
   ForeColor       =   &H80000011&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin NTControls30.Neotext Neotext1 
      Height          =   2595
      Left            =   2130
      TabIndex        =   0
      Top             =   1515
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   4577
      LeftMargin      =   0   'False
      FontSize        =   8.25
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

'    Dim cnt As Long
'    Dim str As String
'    Dim multiply As Long
'    For multiply = 0 To 9
'        For cnt = 0 To 9
'            str = str & String(100, CStr(cnt)) & vbCrLf
'        Next
'    Next


    Neotext1.Text = ReadFile("C:\Development\Kitchen\NTControls30\frmAutoScroll.log")
    Neotext1.leftmargin = True
    Neotext1.LineNumbers = True
    'Debug.Print Neotext1.textweight * Screen.TwipsPerPixelY
    
End Sub

Private Sub Form_Resize()
    Neotext1.Top = 0
    Neotext1.left = 0
    Neotext1.Width = Me.ScaleWidth
    Neotext1.Height = Me.ScaleHeight
    
End Sub
