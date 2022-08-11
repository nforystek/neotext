VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmSpellCheck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spell Checker"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   Icon            =   "frmSpellCheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin RichTextLib.RichTextBox CheckText 
      Height          =   1080
      Left            =   30
      TabIndex        =   7
      Top             =   255
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   1905
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmSpellCheck.frx":058A
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   345
      Index           =   3
      Left            =   4125
      TabIndex        =   6
      Top             =   2730
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change All"
      Enabled         =   0   'False
      Height          =   345
      Index           =   2
      Left            =   4125
      TabIndex        =   5
      Top             =   2070
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change"
      Enabled         =   0   'False
      Height          =   345
      Index           =   1
      Left            =   4125
      TabIndex        =   4
      Top             =   1665
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next Word"
      Enabled         =   0   'False
      Height          =   345
      Index           =   0
      Left            =   4125
      TabIndex        =   1
      Top             =   285
      Width           =   1215
   End
   Begin VB.ListBox Suggestion 
      Height          =   1425
      Left            =   30
      TabIndex        =   0
      Top             =   1650
      Width           =   4020
   End
   Begin VB.Label Label2 
      Caption         =   "Suggestions:"
      Height          =   300
      Left            =   105
      TabIndex        =   3
      Top             =   1410
      Width           =   1980
   End
   Begin VB.Label Label1 
      Caption         =   "Red text not in dictionary:"
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   45
      Width           =   2895
   End
End
Attribute VB_Name = "frmSpellCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private ControlDown As Boolean 'Tells if CTRL key is down, used to catch CTRL+Z (undo) and CTRL+V (paste)
Private ShiftDown As Boolean 'Tells if Shift key is down, used to catch Shift+Insert (paste)

Private Const NoSuggestions = "(No Suggestions)"

Private LastPos As Long
Private CurPos1 As Long
        
Private w1 As Word.Application
Private WordCount As Long

Public Finished As Boolean
Public Function IsSeperator(ByVal Character As String) As Boolean
    IsSeperator = (Character = " " Or Character = vbCrLf Or Character = vbCr Or Character = vbLf Or Character = "." Or Character = "," Or Character = ";" Or Character = ":" Or Character = """")
End Function


Private Sub InsertCharacters(ByVal SelStart As Long, ByVal SelLength As Long)
    If SelStart < CurPos1 Then
        CurPos1 = CurPos1 + SelLength
    End If
    
    If SelStart < LastPos Then
        LastPos = LastPos + SelLength
    End If
    
End Sub

Private Sub RemoveCharacters(ByVal SelStart As Long, ByVal SelLength As Long)
    If SelStart < CurPos1 Then
        CurPos1 = CurPos1 - SelLength
    End If

    If SelStart < LastPos Then
        LastPos = LastPos - SelLength
    End If

End Sub

Private Sub CheckText_KeyDown(KeyCode As Integer, Shift As Integer)
    If ControlDown Then
        'This is for when the user is using Undo CTRL+Z
        If KeyCode = 90 Then
            'KeyCode = 0
        ElseIf KeyCode = 88 Then
        ElseIf KeyCode = 86 Then 'Paste
            InsertCharacters CheckText.SelStart, Len(Clipboard.GetText)
        End If
    ElseIf ShiftDown And KeyCode = 45 Then
        InsertCharacters CheckText.SelStart, Len(Clipboard.GetText)
    Else
    
        If KeyCode = 17 And Shift = 2 Then
            ControlDown = True
            'Catch Control Key to watch for Undo being used
        ElseIf KeyCode = 16 And Shift = 1 Then
            ShiftDown = True
        Else
            
            On Error Resume Next
            If KeyCode = 8 Then 'backspace
                
                If Mid(CheckText.Text, CheckText.SelStart + 1, 1) = vbLf Then
                    If CheckText.SelLength = 0 Then
                        RemoveCharacters CheckText.SelStart, 2
                    Else
                        RemoveCharacters CheckText.SelStart, CheckText.SelLength
                    End If
                Else
                    If CheckText.SelLength = 0 Then
                        RemoveCharacters CheckText.SelStart, 1
                    Else
                        RemoveCharacters CheckText.SelStart, CheckText.SelLength
                    End If
                End If
            ElseIf KeyCode = 46 Then 'delete
                If Mid(CheckText.Text, CheckText.SelStart + 1, 1) = vbCr Then
                    If CheckText.SelLength = 0 Then
                        RemoveCharacters CheckText.SelStart, 2
                    Else
                        RemoveCharacters CheckText.SelStart, CheckText.SelLength
                    End If
                Else
                    If CheckText.SelLength = 0 Then
                        RemoveCharacters CheckText.SelStart + 1, 1
                    Else
                        RemoveCharacters CheckText.SelStart, CheckText.SelLength
                    End If
                End If
            ElseIf KeyCode = 13 Then
                If CheckText.SelLength > 0 Then
                    RemoveCharacters CheckText.SelStart, CheckText.SelLength
                End If
                InsertCharacters CheckText.SelStart, Len(vbCrLf)
            End If
        
        End If
    End If
    

End Sub

Private Sub CheckText_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 10 Then
        If CheckText.SelLength > 0 Then
            RemoveCharacters CheckText.SelStart, CheckText.SelLength
        End If
        InsertCharacters CheckText.SelStart, Len(Chr(KeyAscii))
    End If

End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0 'ignore
            WordCount = WordCount + 1
            
            Command1(0).Enabled = False
            EnableButtons False
            BeginCheck
            
        Case 1 'change
            If Not Suggestion.List(Suggestion.ListIndex) = NoSuggestions Then
            
                SetWord LastPos, Suggestion.List(Suggestion.ListIndex)
                
                Command1(0).Enabled = False
                EnableButtons False
                BeginCheck
            End If
        Case 2 'change all
            If Not Suggestion.List(Suggestion.ListIndex) = NoSuggestions Then
                  
                CheckText.Text = Replace(CheckText.Text, GetWord(LastPos), Suggestion.List(Suggestion.ListIndex))
                
                WordCount = WordCount + 1
                
                Command1(0).Enabled = False
                EnableButtons False
                BeginCheck
            End If
        Case 3 'close
            Finished = True
            Me.Hide
                        
    End Select
End Sub

Private Function SetWord(ByVal Pos1 As Long, ByVal NewWord As String) As String
    
    
    Me.MousePointer = 11
    Dim pos2 As Long
        
    pos2 = Pos1
    Do Until IsSeperator(Mid(CheckText.Text, pos2, 1)) Or pos2 > Len(CheckText.Text)
        DoEvents
        pos2 = pos2 + 1
    Loop
    Do While IsSeperator(Mid(CheckText.Text, pos2 + 1, 1))
        pos2 = pos2 + 1
    Loop
    
    CheckText.SelStart = Pos1 - 1
    CheckText.SelLength = (pos2 - Pos1)
    CheckText.SelText = NewWord     ' Replace the Misspelled word in the corresponding KLTextBox.
    
    If (pos2 - Pos1) > Len(NewWord) Then
        RemoveCharacters Pos1 - 1, (pos2 - Pos1) - Len(NewWord)
    ElseIf Len(NewWord) > (pos2 - Pos1) Then
        InsertCharacters Pos1 - 1, Len(NewWord) - (pos2 - Pos1)
    End If
    
    
    Me.MousePointer = 0
    
    
End Function
Private Function GetWord(ByVal Pos1 As Long) As String
    'Gets next word and DOESN'T change position like GetNextWord()
    
    Me.MousePointer = 11
    Dim pos2 As Long
        
    pos2 = Pos1
    Do Until IsSeperator(Mid(CheckText.Text, pos2, 1)) Or pos2 > Len(CheckText.Text)
        DoEvents
        pos2 = pos2 + 1
    Loop
    Do While IsSeperator(Mid(CheckText.Text, pos2 + 1, 1))
        pos2 = pos2 + 1
    Loop
    
    GetWord = Mid(CheckText.Text, Pos1, pos2 - Pos1)
    
    Me.MousePointer = 0
    
End Function
Private Function GetNextWord(ByRef Pos1 As Long) As String
    'Gets next word and CHANGES positions unlike GetWord()
    
    Me.MousePointer = 11
    Dim pos2 As Long
        
    pos2 = Pos1
    LastPos = Pos1
    Do Until IsSeperator(Mid(CheckText.Text, pos2, 1)) Or pos2 > Len(CheckText.Text)
        DoEvents
        pos2 = pos2 + 1
    Loop
    Do While IsSeperator(Mid(CheckText.Text, pos2 + 1, 1))
        pos2 = pos2 + 1
    Loop
    
    GetNextWord = Mid(CheckText.Text, Pos1, pos2 - Pos1)
    Pos1 = Pos1 + 1
    
    Me.MousePointer = 0
End Function
Private Function SelectWord(ByVal Pos1 As Long) As String


    Me.MousePointer = 11
    Dim pos2 As Long
        
    pos2 = Pos1
    LastPos = Pos1
    Do Until IsSeperator(Mid(CheckText.Text, pos2, 1)) Or pos2 > Len(CheckText.Text)
        DoEvents
        pos2 = pos2 + 1
    Loop
    Do While IsSeperator(Mid(CheckText.Text, pos2 + 1, 1))
        pos2 = pos2 + 1
    Loop
    
    CheckText.SelStart = 0
    CheckText.SelLength = Len(CheckText.Text)
    CheckText.SelColor = vbBlack
    CheckText.SelStart = Pos1 - 1
    CheckText.SelLength = (pos2 - Pos1)
    CheckText.SelColor = vbRed
    CheckText.SelLength = 0
    
    Me.MousePointer = 0

End Function

Public Sub Start(ByVal Text As String)
    
    WordCount = 1
    CurPos1 = 1
    CheckText.Text = Trim(Text)

    If Len(CheckText.Text) > 0 Then
        BeginCheck
    Else
        Finished = True
        Me.Visible = False
    End If
End Sub

Public Sub BeginCheck()
    Dim i, NextWord As String
    NextWord = GetNextWord(CurPos1)
    CurPos1 = CurPos1 + Len(NextWord)
    Do Until NextWord = "" And CurPos1 >= Len(CheckText.Text)
        
        If w1.CheckSpelling(NextWord) = False Then
            
            Me.Visible = True
            fMain.StayOnTop Me, True
            
            SelectWord LastPos
            
            For Each i In w1.GetSpellingSuggestions(NextWord)
                Suggestion.AddItem i
            Next
            
            If Suggestion.ListCount = 0 Then Suggestion.AddItem NoSuggestions
            
            Command1(0).Enabled = True
            
            Exit Sub
        End If
        
        NextWord = GetNextWord(CurPos1)
        CurPos1 = CurPos1 + Len(NextWord)
        WordCount = WordCount + 1
    Loop
    
    Finished = True
    Me.Visible = False
End Sub

Private Sub Form_Load()
    Finished = False
    Set w1 = New Word.Application
    w1.Application.Documents.Add
    ControlDown = False
    ShiftDown = False
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        MsgBox CurPos1
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = False
        Finished = True
        If Me.Visible Then
            fMain.NotStayOnTop Me, True
            Me.Visible = False
        End If
    End If
End Sub

Private Sub EnableButtons(ByVal IsEnabled As Boolean)
    If Not IsEnabled Then Suggestion.Clear
    Command1(1).Enabled = IsEnabled
    Command1(2).Enabled = IsEnabled
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    w1.Quit False
    Set w1 = Nothing
End Sub

Private Sub Suggestion_Click()
    If Suggestion.ListIndex > -1 Then
        EnableButtons True
    Else
        EnableButtons False
    End If
End Sub

Private Sub Suggestion_DblClick()
    If Suggestion.ListIndex > -1 Then
        Command1_Click 1
    End If
End Sub
