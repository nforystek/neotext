VERSION 5.00
Object = "*\A..\..\NTControls30.vbp"
Begin VB.Form frmLogLoad 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin NTControls30.TextBox TextBox1 
      Height          =   2850
      Left            =   885
      TabIndex        =   0
      Top             =   975
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   5027
      MultipleLines   =   -1  'True
      Text            =   $"frmLogLoad.frx":0000
      ScrollBars      =   3
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7440
      Top             =   630
   End
End
Attribute VB_Name = "frmLogLoad"
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
    'On Error Resume Next
On Error GoTo 0

    'Neotext1.sosweet.Interpreter "C:\Development\Neotext\Common\Projects\NTControls30\Test\VBScript.ssw"
    
    If PathExists("C:\Development\Neotext\Common\Projects\NTControls30\Test\VBScript.tmp", True) Then Kill "C:\Development\Neotext\Common\Projects\NTControls30\Test\VBScript.tmp"
    txt = ReadFile("C:\Development\Neotext\Common\Projects\NTControls30\Test\VBScript.txt")

    WriteFile "C:\Development\Neotext\Common\Projects\NTControls30\Test\VBScript.tmp", txt
    
    fnum = FreeFile
    
    TextBox1.Text = txt ' "n3w90assdfgdfsgg8hj" & vbCrLf & _
"a3w9z89" & vbCrLf & _
"o34o9agv9sdfgsdfgha8v" & vbCrLf & _
"sdfgsdfgdssdgsdsdfg" & vbCrLf & _
"aflkhaoihn4ew2o" & vbCrLf & _
"s" & vbCrLf & _
"sasfklj as dhowe" & vbCrLf & _
"s" & vbCrLf
'TextBox1.Text = TextBox1.Text & TextBox1.Text

 '  Neotext1.FileName = "C:\Development\Neotext\Common\Projects\NTControls30\Test\VBScript.tmp"
'
' Neotext1.Locked = True
'
'
'    Neotext1.ScrollDown = True
'    Neotext1.SelStart = Len(Neotext1.Text)
    

    
  '  Timer1.Enabled = True
    


    If Err Then
        Debug.Print "ERROR: " & Err.Description
        Err.Clear

    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Timer1.Enabled = False
End Sub

Private Sub Form_Resize()
    TextBox1.Width = Me.ScaleWidth
    TextBox1.Height = Me.ScaleHeight
    TextBox1.Top = 0
    TextBox1.left = 0

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Close #fnum
End Sub



Private Sub Timer1_Timer()
    Open "C:\Development\Neotext\Common\Projects\NTControls30\Test\VBScript.tmp" For Binary Shared As #fnum
    Put #fnum, LOF(fnum), txt
    Close #fnum
End Sub
