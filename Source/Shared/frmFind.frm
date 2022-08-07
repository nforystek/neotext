VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Replace &All"
      Height          =   330
      Index           =   3
      Left            =   4260
      TabIndex        =   10
      Top             =   1455
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1215
      TabIndex        =   1
      Top             =   450
      Width           =   2865
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2895
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1065
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Match Ca&se"
      Height          =   210
      Index           =   1
      Left            =   2085
      TabIndex        =   5
      Top             =   1545
      Width           =   1560
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Find Whole Word &Only"
      Height          =   210
      Index           =   0
      Left            =   2085
      TabIndex        =   6
      Top             =   1860
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search"
      Height          =   840
      Left            =   150
      TabIndex        =   12
      Top             =   990
      Width           =   1785
      Begin VB.OptionButton Option1 
         Caption         =   "Selected &Text"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   3
         Top             =   525
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&All Text"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   2
         Top             =   255
         Value           =   -1  'True
         Width           =   1545
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Replace..."
      Height          =   330
      Index           =   2
      Left            =   4260
      TabIndex        =   9
      Top             =   1050
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   330
      Index           =   1
      Left            =   4260
      TabIndex        =   8
      Top             =   450
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find &Next"
      Default         =   -1  'True
      Height          =   330
      Index           =   0
      Left            =   4260
      TabIndex        =   7
      Top             =   60
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1215
      TabIndex        =   0
      Top             =   75
      Width           =   2865
   End
   Begin VB.Label Label3 
      Caption         =   "Replace With:"
      Height          =   180
      Left            =   135
      TabIndex        =   14
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Direction:"
      Height          =   240
      Left            =   2100
      TabIndex        =   13
      Top             =   1110
      Width           =   750
   End
   Begin VB.Label Label1 
      Caption         =   "Find What:"
      Height          =   240
      Left            =   135
      TabIndex        =   11
      Top             =   120
      Width           =   810
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private CurrentPos As Long
Private AllDirect As Boolean
Private SelStart As Long
Private SelLength As Long
Private CurStart As Long
Private CurLength As Long

Private Function FindNext(ByVal Down As Boolean) As Boolean
    Dim searchText As String
    Dim findPos As Long
    Dim ret As Boolean
    
        If Down Then
            
            If Option1(2).Value Then
                If SelStart = 0 And SelLength = 0 Then
                    SelStart = GetTextObj.SelStart
                    SelLength = GetTextObj.SelLength
                    CurStart = GetTextObj.SelStart
                    CurLength = GetTextObj.SelLength
                End If
                searchText = Mid(GetTextObj.text, CurStart + 1, CurLength)
            Else
                SelStart = 0
                SelLength = 0
                CurStart = 0
                CurLength = 0
                CurrentPos = GetTextObj.SelStart + GetTextObj.SelLength + 1
                searchText = Mid(GetTextObj.text, CurrentPos)
            End If
            
            If Check1(1).Value Then
                findPos = InStr(searchText, Text1.text)
            Else
                findPos = InStr(LCase(searchText), LCase(Text1.text))
            End If
            
            If findPos > 0 Then
            
                If Option1(2).Value Then
                    GetTextObj.SelStart = CurStart - 1 + findPos
                    CurStart = CurStart + findPos
                    CurLength = CurLength - findPos
                Else
                    GetTextObj.SelStart = (CurrentPos - 2) + findPos
                End If
                GetTextObj.SelLength = Len(Text1.text)
                GetTextObj.SetFocus
                
                ret = True
            Else
                ret = False
            End If
        ElseIf Not Down Then
            
            If Option1(2).Value Then
                
                If SelStart = 0 And SelLength = 0 Then
                    SelStart = GetTextObj.SelStart
                    SelLength = GetTextObj.SelLength
                    CurStart = GetTextObj.SelStart
                    CurLength = GetTextObj.SelLength
                End If
                searchText = Mid(GetTextObj.text, CurStart + 1, CurLength)
            Else
                SelStart = 0
                SelLength = 0
                CurStart = 0
                CurLength = 0
                CurrentPos = GetTextObj.SelStart
                searchText = Left(GetTextObj.text, CurrentPos)
            End If
            
            If Check1(1).Value Then
                findPos = InStrRev(searchText, Text1.text)
            Else
                findPos = InStrRev(LCase(searchText), LCase(Text1.text))
            End If
            
            If findPos > 0 Then
            
                If Option1(2).Value Then
                    GetTextObj.SelStart = CurStart - 1 + findPos
                    CurLength = findPos - 1
                Else
                    GetTextObj.SelStart = findPos - 1
                End If
                GetTextObj.SelLength = Len(Text1.text)
                GetTextObj.SetFocus
                
                ret = True
            Else
                ret = False
            End If
        End If

    FindNext = ret
End Function
Private Function FindWrap(ByVal Button As Boolean) As Boolean
    
    If Combo1.ListIndex = 1 Then
        If (Not FindNext(True)) Then
            NoMatchFound False
        End If
    ElseIf Combo1.ListIndex = 2 Then
        If (Not FindNext(False)) Then
            NoMatchFound False
        End If
    ElseIf Combo1.ListIndex = 0 Then
        If (Not FindNext(AllDirect)) Then
            AllDirect = Not AllDirect
            CurStart = SelStart
            CurLength = SelLength
            If (Not FindNext(AllDirect)) Then
                NoMatchFound False
            End If
        End If
    End If

End Function
Private Sub Command1_Click(Index As Integer)
    
    If Not GetTextObj Is Nothing Then

        Select Case Index
            Case 0
            
                FindWrap True
                
            Case 1
                Unload Me
            Case 2
                If Command1(2).Caption = "&Replace..." Then
                    ReplaceDialog True
                Else
                    
                    If Check1(1).Value Then
                        If GetTextObj.SelText = Text1.text Then
                            GetTextObj.SelText = Text2.text
                        End If
                    Else
                        If LCase(GetTextObj.SelText) = LCase(Text1.text) Then
                            GetTextObj.SelText = Text2.text
                        End If
                    End If
                    
                    Command1_Click 0
                    
                End If
            Case 3
                If Option1(0).Value Then
                    If Check1(1).Value Then
                        GetTextObj.text = Replace(GetTextObj.text, Text1.text, Text2.text, , , vbBinaryCompare)
                    Else
                        GetTextObj.text = Replace(GetTextObj.text, Text1.text, Text2.text, , , vbTextCompare)
                    End If
                Else
                    If SelStart > 0 And SelLength > 0 Then
                        GetTextObj.SelStart = SelStart
                        GetTextObj.SelLength = SelLength
                    End If
                    If Check1(1).Value Then
                        GetTextObj.SelText = Replace(GetTextObj.SelText, Text1.text, Text2.text, , , vbBinaryCompare)
                    Else
                        GetTextObj.SelText = Replace(GetTextObj.SelText, Text1.text, Text2.text, , , vbTextCompare)
                    End If
                End If
            
        End Select
    End If
End Sub

Private Sub NoMatchFound(ByVal Button As Boolean)

    DoTasks
    
    If Button Then

        FindWrap False
    End If

    MsgBox "No match found.", vbInformation, "Notepad"
    
    TopMostForm Me, True
    
End Sub
Private Sub ReplaceDialog(ByVal Enabled As Boolean)
    If Enabled Then
        Me.Height = 2310
        Frame1.Top = 990
        Label2.Top = 1110
        Combo1.Top = 1065
        Check1(1).Top = 1545
        Check1(0).Top = 1860
        Label3.Visible = True
        Text2.Visible = True
        Command1(3).Visible = True
        Command1(2).Caption = "&Replace"
    Else
        Me.Height = 1830
        Frame1.Top = 480
        Label2.Top = 600
        Combo1.Top = 555
        Check1(1).Top = 1035
        Check1(0).Top = 1350
        Label3.Visible = False
        Text2.Visible = False
        Command1(3).Visible = False
        Command1(2).Caption = "&Replace..."
    End If
End Sub

Private Sub ValidateUI()
    
    Command1(0).Enabled = (Text1.text <> "")
    
    Option1(2).Enabled = (GetTextObj.SelLength > 0)

    Command1(2).Enabled = (Not GetTextObj.Locked)
    
    Command1(3).Enabled = (Not GetTextObj.Locked)
End Sub

Private Sub Form_Activate()
    ValidateUI
End Sub

Private Sub Form_GotFocus()
    ValidateUI
End Sub

Private Sub Form_Load()
    Combo1.AddItem "All"
    Combo1.AddItem "Down"
    Combo1.AddItem "Up"
    Combo1.ListIndex = 1
    Me.Visible = True
    DoTasks
    TopMostForm Me, True
    ReplaceDialog False
    AllDirect = True
    ValidateUI
    If Option1(2).Enabled Then Option1(2).Value = True
    SelStart = 0
    SelLength = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
    TopMostForm Me, False
End Sub

Private Sub Text1_Change()
    ValidateUI
End Sub




Attribute 