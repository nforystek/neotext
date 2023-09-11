VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "    "
   ClientHeight    =   7950
   ClientLeft      =   23625
   ClientTop       =   4170
   ClientWidth     =   13410
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7950
   ScaleWidth      =   13410
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command15 
      Caption         =   "Last"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2925
      TabIndex        =   24
      Top             =   900
      Width           =   975
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Check"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1020
      TabIndex        =   23
      Top             =   900
      Width           =   855
   End
   Begin VB.CommandButton Command12 
      Caption         =   "First"
      Height          =   375
      Left            =   1920
      TabIndex        =   22
      Top             =   900
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Point"
      Height          =   375
      Left            =   1920
      TabIndex        =   21
      Top             =   480
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Moving"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3990
      TabIndex        =   20
      Top             =   1110
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Add/Del"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3990
      TabIndex        =   19
      Top             =   885
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   900
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1020
      TabIndex        =   16
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Prev"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Rand 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Goto 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2940
      TabIndex        =   6
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Add *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2940
      TabIndex        =   5
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Del *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Del"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1020
      TabIndex        =   2
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Rand *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   60
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4245
      Top             =   810
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10080
      TabIndex        =   14
      Top             =   480
      Width           =   1515
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8400
      TabIndex        =   13
      Top             =   480
      Width           =   1635
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   12
      Top             =   480
      Width           =   1515
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   11
      Top             =   480
      Width           =   1635
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6840
      TabIndex        =   10
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5160
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   8
      Top             =   840
      Width           =   6495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10080
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Const Max = 3
Dim mode As Long

Private rows As Long
Private msg As String

Public Sub PrintPrograms()
 
    Me.Cls
    Dim str As String

    str = vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
    Dim H As Single
    Dim W As Single
    rows = 0
    
    H = 0
    Do Until str = ""

            
        W = 0
        Do While (W < Me.ScaleWidth) And (Not (str = ""))
            W = W + Me.TextWidth(Left(str, 1))
            If Left(str, 2) = vbCrLf Then
                str = Mid(str, 3)
                Exit Do
            Else
                Me.Print Left(str, 1);
                str = Mid(str, 2)
            End If
            
        Loop
        Me.Print
        H = H + Me.TextHeight("X")
    Loop
    If str <> "" Then
        Me.Print str
        H = H + Me.TextHeight("X")
        str = ""
    End If

    Do While H < Me.ScaleHeight
        H = H + Me.TextHeight("X")
        rows = rows + 1
    Loop

    Do While CountWord(msg, vbCrLf) >= rows
        RemoveNextArg msg, vbCrLf
    Loop
    
    Me.Print msg

End Sub

Public Sub PrintMessage(ByVal txt As String, Optional ByVal NewLine As Boolean = True)

    msg = msg & txt & IIf(NewLine, vbCrLf, "")
       
End Sub

Public Function CountWord(ByVal Text As String, ByVal Word As String, Optional ByVal Exact As Boolean = True) As Long
    Dim cnt As Long
    cnt = UBound(Split(Text, Word, , IIf(Exact, vbBinaryCompare, vbTextCompare)))
    If cnt > 0 Then CountWord = cnt
End Function

Public Function RemoveNextArg(ByRef TheParams As Variant, ByVal TheSeperator As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As String
    If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
        RemoveNextArg = Trim(Left(TheParams, InStr(1, TheParams, TheSeperator, Compare) - 1))
        TheParams = Trim(Mid(TheParams, InStr(1, TheParams, TheSeperator, Compare) + Len(TheSeperator)))
    Else
        RemoveNextArg = Trim(TheParams)
        TheParams = ""
    End If
End Function

Public Function NextArg(ByVal TheParams As Variant, ByVal TheSeperator As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As String
    If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
        NextArg = Trim(Left(TheParams, InStr(1, TheParams, TheSeperator, Compare) - 1))
    Else
        NextArg = Trim(TheParams)
    End If
End Function


Public Sub Status()

#If VBIDE Then
    Label1.ForeColor = IIf(Total(Addrs) = sec.count, vbBlack, vbRed)
    Label1.Caption = Total(Addrs) & " " & sec.count
#Else
    Label1.Caption = Total(Addrs)
#End If
    

 Label3.Caption = "Eol: " & -CInt(IsEol(Addrs)) & "   Inv: " & -CInt(IsInv(Addrs)) _
         & "   Mpar: " & -CInt(MPar(Addrs)) & "   Cpar: " & -CInt(CPar(Addrs)) & "   GPar: " & -CInt(GPar(Addrs))


'    Label7.Caption = Addrs.B
   ' Label8.Caption = Module.Focus(Addrs)

    Label6.Caption = "Point: " & Module.Point(Addrs)
    If Timer1.Enabled = False Then
        Command1.Caption = "Rand *"
        Command5.Caption = "Del *"
        Command6.Caption = "Add *"
        Command7.Caption = "Goto " & Max
        Command8.Caption = "Rand " & Max
    End If
    PrintPrograms
    
    
End Sub

Private Sub Command10_Click()
   Label4.Caption = "Next: " & NextNode(Addrs, (Check2.Value = 1))
    Status
End Sub

Private Sub Command11_Click()
#If VBIDE Then
    Dim Addr As Variant
    For Each Addr In sec
        sec.Free CLng(Addr)
    Next
#End If
    With Addrs
        .A = 0
        .B = 0
        .C = 0
        .D = 0
        .E = 0
        .F = 0
        .G = 0
        .H = 0
        .I = 0
        .J = 0
    End With
    Status
End Sub

Private Sub Command13_Click()
   ' Label7.Caption = "Check: " & Module.check(Addrs)
End Sub

Private Sub Command2_Click()
    Dim Addr As Long
    
    PushNode Addrs, Addr
    
    
    Label6.Caption = "Point: " & Addr
    Status
End Sub

Private Sub Command3_Click()
    If Text1.Text <> "" Then
            
        KillNode Addrs ',  CLng(Text1.Text)
    Else
        KillNode Addrs
    End If
    Status
End Sub

Private Sub Command1_Click()
    If Not Command1.Caption = "Stop" Then
        Command1.Caption = "Stop"
        Timer1.Enabled = True
        mode = 1
    Else
        Command1.Caption = "Rand *"
        Timer1.Enabled = False
        Status
    End If
End Sub

Private Sub Command4_Click()

    Label6.Caption = "Point: " & Module.Point(Addrs)
End Sub

Private Sub Command5_Click()
    If Not Command5.Caption = "Stop" Then
        Command5.Caption = "Stop"
        Timer1.Enabled = True
    Else
        Command5.Caption = "Del *"
        Timer1.Enabled = False
        Status
    End If
End Sub

Private Sub Command6_Click()
    If Not Command6.Caption = "Stop" Then
        Command6.Caption = "Stop"
        Timer1.Enabled = True
    Else
        Command6.Caption = "Add *"
        Timer1.Enabled = False
        Status
    End If
End Sub

Private Sub Command7_Click()
    If Not Command7.Caption = "Stop" Then
        Command7.Caption = "Stop"
        Timer1.Enabled = True
    Else
        Command7.Caption = "Goto " & Max
        Timer1.Enabled = False
        Status
    End If
End Sub

Private Sub Command8_Click()
    If Not Command8.Caption = "Stop" Then
        Command8.Caption = "Stop"
        Timer1.Enabled = True
    Else
        Command8.Caption = "Rand " & Max
        Timer1.Enabled = False
        Status
    End If
End Sub

Private Sub Command9_Click()
    Label5.Caption = "Prev: " & PrevNode(Addrs, (Check2.Value = 1))
    Status
End Sub

Private Sub Form_Load()
    Setup
    Command1.Caption = "Rand *"
    Command5.Caption = "Del *"
    Command6.Caption = "Add *"
    Command7.Caption = "Goto " & Max
    Command8.Caption = "Rand " & Max
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Do While Total(Addrs) > 0
'        KillNode Addrs
'    Loop
    
#If VBIDE = -1 Then
    CleanUp

#End If
End Sub

Private Sub Label4_Click()
    Text1.Text = NextArg(Label4.Caption, " ")
End Sub

Private Sub Label5_Click()
    Text1.Text = NextArg(Label5.Caption, " ")
End Sub

Private Sub Label6_Click()
    Text1.Text = NextArg(Label6.Caption, " ")
End Sub

Private Sub Label7_Click()
    Text1.Text = NextArg(Label7.Caption, " ")
End Sub

Private Sub Timer1_Timer()
    Static ref As Long
    If Command6.Caption = "Stop" Then
        PushNode Addrs
    ElseIf Command5.Caption = "Stop" Then
        KillNode Addrs
    ElseIf Command7.Caption = "Stop" Then
        Static toggle As Boolean
        If (Not toggle) And Total(Addrs) <= Abs(Max) Then
            If Total(Addrs) < Max Then
                PushNode Addrs
            Else
                toggle = Not toggle
            End If
        ElseIf Total(Addrs) >= 0 And toggle Then
            If Total(Addrs) > 0 Then
                KillNode Addrs
           Else
                toggle = Not toggle
            End If

        End If
        If Check2.Value = 1 Then
            If CBool((Rnd >= 0.5)) Then
                Label5.Caption = NextNode(Addrs)
            Else
                Label4.Caption = PrevNode(Addrs)
            End If
         End If
    ElseIf Command8.Caption = "Stop" Then
        If (Rnd >= 0.5) And Total(Addrs) < Max Then
            PushNode Addrs
        ElseIf Total(Addrs) > 0 Then
            KillNode Addrs
        End If
        If Check2.Value = 1 Then
            If CBool((Rnd >= 0.5)) Then
                Label5.Caption = NextNode(Addrs)
            Else
                Label4.Caption = PrevNode(Addrs)
            End If
         End If
         
    ElseIf Command1.Caption = "Stop" Then
        If (Rnd >= 0.5) Then
            PushNode Addrs
        ElseIf Total(Addrs) > 0 Then
            KillNode Addrs
        End If
        If Check2.Value = 1 Then
            If CBool((Rnd >= 0.5)) Then
                Label5.Caption = NextNode(Addrs)
            Else
                Label4.Caption = PrevNode(Addrs)
            End If
         End If
    End If

    Status
End Sub



