VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Manual Refresh"
      Height          =   1695
      Left            =   4455
      TabIndex        =   4
      Top             =   1935
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   195
      TabIndex        =   3
      Top             =   1470
      Width           =   4080
   End
   Begin VB.ListBox List2 
      Height          =   1230
      Left            =   180
      TabIndex        =   2
      Top             =   150
      Width           =   4110
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3930
      Top             =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Constant Random"
      Height          =   1620
      Left            =   4515
      TabIndex        =   1
      Top             =   135
      Width           =   990
   End
   Begin VB.ListBox List1 
      Height          =   5325
      Left            =   225
      TabIndex        =   0
      Top             =   1860
      Width           =   4035
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit



Private v As NTNodes10.FragMem



Public Function RandomPositive(ByVal Lowerbound As Integer, ByVal Upperbound As Integer) As Integer

    Randomize

    RandomPositive = Int((Upperbound - Lowerbound + 1) * Rnd + Lowerbound)

End Function



Private Sub Command1_Click()

    If Command1.Caption = "Constant Random" Then

        Timer1.Enabled = True

    

        Command1.Caption = "Stop"

    Else

        Timer1.Enabled = False

    

        Command1.Caption = "Constant Random"

    End If

End Sub



Private Sub Command2_Click()

    List1.Clear

    

    Dim cnt As Long

    For cnt = 1 To v.count

        List1.AddItem v.SizeOf(v.handle(cnt)) & ":" & v.handle(cnt) & ":" & v.Data(v.handle(cnt))

    Next

    

End Sub



Private Sub Form_Load()

    List2.AddItem "Some term"

    List2.AddItem "This is a sentence longer."

    List2.AddItem "This is an even longer sentence."

    List2.AddItem "Another term"

    List2.AddItem "word"

    

    Set v = New FragMem

    v.FilePath = AppPath & AppEXE(True, True) & ".idx"

    

    

    Dim cnt As Long

    For cnt = 1 To v.count

        List1.AddItem v.SizeOf(v.handle(cnt)) & ":" & v.handle(cnt) & ":" & v.Data(v.handle(cnt))

    Next



    

End Sub



Private Sub Form_Unload(Cancel As Integer)

    Set v = Nothing

End Sub



Private Sub List1_Click()

    If List1.ListIndex > -1 Then

        Text1.Text = v.Data(CLng(NextArg(RemoveArg(List1.List(List1.ListIndex), ":"), ":")))

    End If

End Sub



Private Sub List1_DblClick()

    If List1.ListIndex > -1 Then

        v.Dealloc CLng(NextArg(RemoveArg(List1.List(List1.ListIndex), ":"), ":"))

        List1.RemoveItem List1.ListIndex

    End If

End Sub



Private Sub List2_Click()

    If List2.ListIndex > -1 Then Text1.Text = List2.List(List2.ListIndex)

 

End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And Text1.Text <> "" Then



        Dim Size As Long

        Size = v.Allocate(Len(Text1.Text))

        v.Data(Size) = Text1.Text

        List1.AddItem Len(Text1.Text) & ":" & Size & ":" & v.Data(Size)

            

    End If

End Sub



Private Sub Timer1_Timer()

    Dim nextAction As Integer

    Dim lstIndex As Integer

    

    If List1.ListCount > 0 Then

        nextAction = RandomPositive(1, 3)

    Else

        nextAction = 1

    End If

    

    If nextAction > 1 Then

        If List1.ListCount - 1 > 0 Then

            lstIndex = RandomPositive(1, List1.ListCount) - 1

        End If

    End If

    

    Dim Size As Long

    

    Select Case nextAction

        Case 1 'add

            nextAction = RandomPositive(1, List2.ListCount) - 1

            Size = v.Allocate(Len(List2.List(nextAction)))

            v.Data(Size) = List2.List(nextAction)

            List1.AddItem Len(List2.List(nextAction)) & ":" & Size & ":" & v.Data(Size)

        Case 2 'realloc

            Size = CLng(NextArg(List1.List(lstIndex), ":"))

            Do

                nextAction = RandomPositive(1, List2.ListCount) - 1

            Loop Until Size <> Len(List2.List(nextAction))

            Size = CLng(NextArg(RemoveArg(List1.List(lstIndex), ":"), ":"))

            v.Realloc Size, Len(List2.List(nextAction))

            v.Data(Size) = List2.List(nextAction)

            List1.RemoveItem lstIndex

            List1.AddItem Len(List2.List(nextAction)) & ":" & Size & ":" & v.Data(Size)

        Case 3 'dealloc

            v.Dealloc CLng(NextArg(RemoveArg(List1.List(lstIndex), ":"), ":"))

            List1.RemoveItem lstIndex

    End Select

    

End Sub



Private Sub Timer2_Timer()

    'Picture1.Cls



    'Picture1.Print v.DebugPrint

    

End Sub

