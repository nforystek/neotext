VERSION 5.00
Object = "*\A..\..\NTControls30.vbp"
Begin VB.Form frmColoring 
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
      Height          =   1965
      Left            =   885
      TabIndex        =   0
      Top             =   435
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   3466
      LeftMargin      =   0   'False
      FontSize        =   8.25
   End
End
Attribute VB_Name = "frmColoring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim txt As String


Private Sub Form_Load()
    On Error Resume Next

    
    If PathExists("C:\Development\Neotext\Common\Projects\NTControls30\Test\VBScript.tmp", True) Then Kill "C:\Development\Neotext\Common\Projects\NTControls30\Test\VBScript.tmp"
    txt = ReadFile("C:\Development\Neotext\Common\Projects\NTControls30\Test\VBScript.txt")
    WriteFile "C:\Development\Neotext\Common\Projects\NTControls30\Test\VBScript.tmp", txt
    
    
    
    Neotext1.SoSweet.Interpreter "C:\Development\Neotext\Common\Projects\NTControls30\Test\Sweets\Visual Basic 6.ssw"


    Neotext1.FileName = "C:\Development\Neotext\Common\Projects\NTControls30\Test\VBScript.tmp"

    Debug.Print Neotext1.TextRTF
    


End Sub

Private Sub Form_Resize()
    Neotext1.Width = Me.ScaleWidth
    Neotext1.Height = Me.ScaleHeight
    Neotext1.Top = 0
    Neotext1.left = 0
End Sub

