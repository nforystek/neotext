VERSION 5.00
Object = "{BC0595AA-6535-436C-9B38-7F880B88E2B9}#14.0#0"; "NTControls30.ocx"
Begin VB.Form frmLogLoad 
   AutoRedraw      =   -1  'True
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   12180
   StartUpPosition =   3  'Windows Default
   Begin NTControls30.TextBox Textbox1 
      Height          =   2655
      Left            =   1260
      TabIndex        =   6
      Top             =   1080
      Width           =   5070
      _ExtentX        =   8943
      _ExtentY        =   4683
      Fontsize        =   8.25
      GreyNoTextMsg   =   "(Enter here)"
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Set Jumbo 2"
      Height          =   480
      Left            =   10800
      TabIndex        =   0
      Top             =   2475
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Set Jumbo 1"
      Height          =   510
      Left            =   10530
      TabIndex        =   5
      Top             =   1815
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Single"
      Height          =   540
      Left            =   10545
      TabIndex        =   4
      Top             =   3780
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Isolated"
      Height          =   480
      Left            =   10740
      TabIndex        =   3
      Top             =   3075
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CodePage +1"
      Height          =   435
      Left            =   10635
      TabIndex        =   2
      Top             =   1290
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "COdePage-1"
      Height          =   480
      Left            =   10845
      TabIndex        =   1
      Top             =   645
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3810
      Top             =   5655
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


Private Sub Command1_Click()


   ' Textbox1.codepage = Textbox1.codepage - 1

End Sub


Private Sub Command2_Click()


'Textbox1.codepage = Textbox1.codepage + 1

End Sub



Private Sub Command3_Click()


 '   Textbox1.seperator(1) = 5

    

End Sub



Private Sub Command4_Click()


'Textbox1.seperator(2) = 9



End Sub



Private Sub Command5_Click()


    txt = ReadFile("C:\Development\Neotext\Common\Projects\NTControls30\Test\VBScript.txt")



'

'Debug.Print Len(Chr(3) & "0,15" & Chr(3) & "1,14" & Chr(3) & "2,13" & Chr(3) & "3,12" & Chr(3) & "4,11" & Chr(3) & "5,10" & Chr(3) & "6,09" & Chr(3) & "7,08" & Chr(3) & "8,07" & Chr(3) & "9,06" & Chr(3) & "10,05" & Chr(3) & "11,04" & Chr(3) & "12,03" & Chr(3) & "13,02" & Chr(3) & "14,01" & Chr(3) & "15,00");



'Debug.Print Textbox1.CharacterCountOfEscapeColorFormatting(Chr(3) & "0,15zero" & Chr(3) & "1,14one" & Chr(3) & "2,13two" & Chr(3) & "3,12three" & Chr(3) & "4,11four" & Chr(3) & "5,10five" & Chr(3) & "6,09six" & Chr(3) & "7,08seven" & Chr(3) & "8,07eight" & Chr(3) & "9,06nine" & Chr(3) & "10,05ten" & Chr(3) & "11,04eleven" & Chr(3) & "12,03tweleve" & Chr(3) & "13,02thirteen" & Chr(3) & "14,01fourteen" & Chr(3) & "15,00fifteen")



'Debug.Print Textbox1.StripOfEscapeColorFormatting(Chr(3) & "0,15zero" & Chr(3) & "1,14one" & Chr(3) & "2,13two" & Chr(3) & "3,12three" & Chr(3) & "4,11four" & Chr(3) & "5,10five" & Chr(3) & "6,09six" & Chr(3) & "7,08seven" & Chr(3) & "8,07eight" & Chr(3) & "9,06nine" & Chr(3) & "10,05ten" & Chr(3) & "11,04eleven" & Chr(3) & "12,03tweleve" & Chr(3) & "13,02thirteen" & Chr(3) & "14,01fourteen" & Chr(3) & "15,00fifteen")

End Sub



Private Sub Command6_Click()


'Textbox1.Text = "a3w9z89" & vbCrLf & _
"o34o9agv9sdfgsdfgha8v" & vbCrLf & _
"sdfgsdfgdssdgsdsdfg" & vbCrLf & _
"aflkhaoihn4ew2o" & vbCrLf & _
"s" & vbCrLf & _
"sasfklj as dhowe" & vbCrLf & _
"s" & vbCrLf

End Sub



Private Sub Form_Load()


    'On Error Resume Next

On Error GoTo 0



    'Textbox1.sosweet.Interpreter "C:\Development\Neotext\Common\Projects\NTControls30\Test\VBScript.ssw"

    

'    If PathExists("C:\Development\Neotext\Common\Projects\NTControls30\Test\VBScript.tmp", True) Then Kill "C:\Development\Neotext\Common\Projects\NTControls30\Test\VBScript.tmp"

    txt = ReadFile("C:\Development\Neotext\Common\Projects\NTControls30\Test\VBScript.txt")

  '  Textbox1.Text = txt & txt & txt & txt & txt & txt & txt & txt & txt & txt & txt & txt & txt & _
                     txt & txt & txt & txt & txt & txt & txt & txt & txt & txt & txt & txt & txt ' "n3w90assdfgdfsgg8hj" & vbCrLf



'Textbox1.Text = "a3w9z89" & vbCrLf & _
"o34o9agv9sdfgsdfgha8v" & vbCrLf & _
"sdfgsdfgdssdgsdsdfg" & vbCrLf & _
"aflkhaoihn4ew2o" & vbCrLf & _
"s" & vbCrLf & _
"sasfklj as dhowe" & vbCrLf & _
"s" & vbCrLf

'

   ' txt = Chr(3) & "0,15zero" & Chr(3) & "1,14one" & Chr(3) & "2,13two" & Chr(3) & "3,12three" & Chr(3) & "4,11four" & Chr(3) & "5,10five" & Chr(3) & "6,09six" & Chr(3) & "7,08seven" & Chr(3) & "8,07eight" & Chr(3) & "9,06nine" & Chr(3) & "10,05ten" & Chr(3) & "11,04eleven" & Chr(3) & "12,03tweleve" & Chr(3) & "13,02thirteen" & Chr(3) & "14,01fourteen" & Chr(3) & "15,00fifteen"  '& Chr(2) & "bold" & Chr(29) & "italic" & Chr(31) & "underline" & Chr(30) & "strikethru"
   ' Textbox1.Text = txt & Chr(10) & Chr(10) & txt & Chr(10) & Chr(10) & txt & Chr(10) & Chr(10) & txt & Chr(10) & Chr(10) & txt & Chr(10) & Chr(10) & txt & Chr(10) & Chr(10) & txt & Chr(10) & Chr(10) & txt & Chr(10) & Chr(10) & txt & Chr(10) & Chr(10) & txt & Chr(10) & Chr(10) & txt & Chr(10) & Chr(10) & txt & Chr(10) & Chr(10) & txt & Chr(10) & Chr(10) & txt & Chr(10) & Chr(10) & txt & Chr(10) & Chr(10) & txt & Chr(10) & Chr(10) & txt & Chr(10) & Chr(10) & txt & Chr(10) & Chr(10) & txt & Chr(10) & Chr(10) & txt & Chr(10) & Chr(10) & txt & Chr(10) & Chr(10)
    

    Textbox1.Text = txt 'FullPaletteText ' Textbox1.ColorPalette
    
    
    
    Dim found As Long
    Do
    
        found = Textbox1.FindText("End", found)
        If found >= 0 Then
            
            Textbox1.SelStart = Textbox1.FindText("End", found)
            found = found + 1
        End If
        
    Loop While found >= 0
    


    'Textbox1.PasswordChar = "*"
    
    
    'Debug.Print Textbox1.FindText("Optional")
     


    If Err Then

        Debug.Print "ERROR: " & Err.Description

        Err.Clear



    End If

End Sub

Private Function FullPaletteText()

    Dim txt As String
    Dim tmp As Long
    Dim cnt As Long
   txt = txt & " "
    For cnt = 0 To 15
        tmp = CLng(InvertNum(cnt + 1, 16))
        txt = txt & Chr(3) & IIf(tmp <= 9, "0" & CStr(tmp), CStr(tmp)) & "," & IIf(cnt <= 9, "0" & CStr(cnt), CStr(cnt)) & IIf(cnt <= 9, CStr(cnt), CStr(cnt))
    Next
    txt = txt & vbLf
    For cnt = 16 To 87
        txt = txt & Chr(3) & "98," & CStr(cnt) & CStr(cnt)
        If (cnt - 15) Mod 12 = 0 Then
            txt = txt & vbLf
        End If
    Next

    For cnt = 88 To 98
        
        txt = txt & Chr(3) & CStr(CLng(InvertNum(cnt - 87, 18)) + 98 - 17) & "," & CStr(cnt + 1) & CStr(cnt)
        If (cnt - 15) Mod 12 = 0 Then
            txt = txt & vbLf
        End If
    Next
    txt = txt & Chr(3) & "16,9999"
    Debug.Print txt
    FullPaletteText = txt
End Function


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


    Timer1.Enabled = False

End Sub



Private Sub Form_Resize()


    Textbox1.Width = Me.ScaleWidth

    Textbox1.Height = Me.ScaleHeight

    Textbox1.Top = 0

    Textbox1.Left = 0



End Sub


Private Sub Form_Unload(Cancel As Integer)


    Close #fnum

End Sub




Public Function Strands(ByRef b() As Byte) As Strands

    Set Strands = New Strands
    Strands.Concat b
End Function


Private Sub Textbox1_ColorText(ByVal ViewOffset As Long, ByVal ViewWidth As Long)

'   Debug.Print "ColorView(" & ViewOffset & ", " & ViewWidth & ")"


    Textbox1.ColorText vbRed, , 5, 15

    Textbox1.ColorText vbBlue, , 8, 3
    
    
End Sub

Private Sub Textbox1_SelChange()

 '   Debug.Print "SelChange"
End Sub

Private Sub Timer1_Timer()


    Open "C:\Development\Neotext\Common\Projects\NTControls30\Test\VBScript.tmp" For Binary Shared As #fnum

    Put #fnum, LOF(fnum), txt

    Close #fnum

End Sub

