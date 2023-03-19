VERSION 5.00
Object = "{C98B112F-745F-4542-B5B3-DDFADF1F6E2F}#1356.0#0"; "NTControls22.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin NTControls22.SiteInformation SiteInformation1 
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   2990
   End
   Begin NTControls22.CodeEdit CodeEdit1 
      Height          =   3255
      Left            =   6360
      TabIndex        =   0
      Top             =   600
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5741
      FontSize        =   9
      ColorDream1     =   8388736
      ColorDream2     =   8388608
      ColorDream3     =   8421376
      ColorDream4     =   32768
      ColorDream5     =   32896
      ColorDream6     =   16512
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN
Option Compare Text


Private Sub CodeEdit1_Click()
    CodeEdit1.errorline = CodeEdit1.Linenumber + 1
        Debug.Print CodeEdit1.TextRTF
End Sub

Private Sub Form_Load()
    CodeEdit1.Language = "VBScript"
    'CodeEdit1.Text = ReadFile("C:\Development\Neotext\Common\Projects\NTControls30\Test\VBScript.txt")
    CodeEdit1.FileName = "\\desktop\c$\Development\Neotext\Common\Projects\NTControls30\Test\VBScript.txt"

    
End Sub
