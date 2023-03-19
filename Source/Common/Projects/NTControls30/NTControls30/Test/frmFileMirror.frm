VERSION 5.00
Object = "*\A..\..\NTControls30.vbp"
Begin VB.Form frmFileMirror 
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
      Height          =   2010
      Left            =   1245
      TabIndex        =   0
      Top             =   510
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   3545
      LeftMargin      =   0   'False
      FontSize        =   8.25
   End
End
Attribute VB_Name = "frmFileMirror"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'TOP DOWN
Option Compare Text

Private txt As String

Dim fnum As Integer

Private Sub Form_Load()
    On Error Resume Next

    
    If PathExists("C:\Development\Neotext\Common\Projects\NTControls30\Test\VBScript.tmp", True) Then Kill "C:\Development\Neotext\Common\Projects\NTControls30\Test\VBScript.tmp"
    txt = ReadFile("C:\Development\Neotext\Common\Projects\NTControls30\Test\VBScript.txt")
    WriteFile "C:\Development\Neotext\Common\Projects\NTControls30\Test\VBScript.tmp", txt
    


    Neotext1.FileName = "C:\Development\Neotext\Common\Projects\NTControls30\Test\VBScript.tmp"

    


End Sub

Private Sub Form_Resize()
    Neotext1.Width = Me.ScaleWidth
    Neotext1.Height = Me.ScaleHeight
    Neotext1.Top = 0
    Neotext1.left = 0
End Sub




