VERSION 5.00
Begin VB.UserControl ctlOptions 
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4068
   LockControls    =   -1  'True
   ScaleHeight     =   1800
   ScaleWidth      =   4068
   ToolboxBitmap   =   "ctlOptions.ctx":0000
   Begin VB.CheckBox Check1 
      Caption         =   "Graphics Folder"
      Height          =   210
      Index           =   9
      Left            =   300
      TabIndex        =   10
      Top             =   1500
      Width           =   1470
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Project Folder"
      Height          =   210
      Index           =   8
      Left            =   300
      TabIndex        =   9
      Top             =   1212
      Width           =   1470
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Schedule Info"
      Height          =   210
      Index           =   7
      Left            =   2595
      TabIndex        =   8
      Top             =   1500
      Width           =   1470
   End
   Begin VB.CheckBox Check1 
      Caption         =   "File Associations"
      Height          =   210
      Index           =   1
      Left            =   300
      TabIndex        =   2
      Top             =   645
      Width           =   1560
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Favorites Folder"
      Height          =   210
      Index           =   2
      Left            =   300
      TabIndex        =   3
      Top             =   930
      Width           =   1560
   End
   Begin VB.CheckBox Check1 
      Caption         =   "User Settings"
      Height          =   210
      Index           =   3
      Left            =   2370
      TabIndex        =   4
      Top             =   348
      Width           =   1305
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Active App"
      Height          =   195
      Index           =   4
      Left            =   2595
      TabIndex        =   5
      Top             =   636
      Width           =   1470
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Session Drives"
      Height          =   210
      Index           =   5
      Left            =   2595
      TabIndex        =   6
      Top             =   924
      Width           =   1470
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Visited History"
      Height          =   210
      Index           =   6
      Left            =   2595
      TabIndex        =   7
      Top             =   1212
      Width           =   1470
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Select All"
      Height          =   210
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   1560
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Public Settings"
      Height          =   210
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Top             =   360
      Width           =   1560
   End
End
Attribute VB_Name = "ctlOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private CheckboxCancel As Boolean
Public Event OptionChanged()

Public Property Get Options() As Integer
    Dim CheckboxOptions As Integer
    CheckboxOptions = 0
    BitWord(CheckboxOptions, bo_PublicSettings) = Check1(0).Value
    BitWord(CheckboxOptions, bo_FileAssociations) = Check1(1).Value
    BitWord(CheckboxOptions, bo_Favorites) = Check1(2).Value
    BitWord(CheckboxOptions, bo_Projects) = Check1(8).Value
    BitWord(CheckboxOptions, bo_UserSettings) = Check1(3).Value
    BitWord(CheckboxOptions, bo_ActiveApp) = Check1(4).Value
    BitWord(CheckboxOptions, bo_SessionDrives) = Check1(5).Value
    BitWord(CheckboxOptions, bo_VisitedHistory) = Check1(6).Value
    BitWord(CheckboxOptions, bo_Schedules) = Check1(7).Value
    BitWord(CheckboxOptions, bo_Graphics) = Check1(8).Value
    Options = CheckboxOptions
End Property

Public Property Let Options(ByVal newVal As Integer)
    CheckboxCancel = True
    
    Check1(0).Value = BoolToCheck(BitWord(newVal, bo_PublicSettings))
    Check1(1).Value = BoolToCheck(BitWord(newVal, bo_FileAssociations))
    Check1(2).Value = BoolToCheck(BitWord(newVal, bo_Favorites))
    Check1(8).Value = BoolToCheck(BitWord(newVal, bo_Projects))
    Check1(3).Value = BoolToCheck(BitWord(newVal, bo_UserSettings))
    Check1(4).Value = BoolToCheck(BitWord(newVal, bo_ActiveApp))
    Check1(5).Value = BoolToCheck(BitWord(newVal, bo_SessionDrives))
    Check1(6).Value = BoolToCheck(BitWord(newVal, bo_VisitedHistory))
    Check1(7).Value = BoolToCheck(BitWord(newVal, bo_Schedules))
    Check1(9).Value = BoolToCheck(BitWord(newVal, bo_Graphics))
    
    Check1(0).enabled = Check1(0).Value
    Check1(1).enabled = Check1(1).Value
    Check1(2).enabled = Check1(2).Value
    Check1(3).enabled = Check1(3).Value
    Check1(4).enabled = Check1(4).Value
    Check1(5).enabled = Check1(5).Value
    Check1(6).enabled = Check1(6).Value
    Check1(7).enabled = Check1(7).Value
    Check1(8).enabled = Check1(8).Value
    Check1(9).enabled = Check1(9).Value
    
    RefreshSelectAll
    
    CheckboxCancel = False
End Property

Private Function RefreshSelectAll()
    Dim newVal As Boolean
    Dim cnt As Integer

    newVal = True
    For cnt = 0 To Check1.count - 1
        If Check1(cnt).Value = 0 Then
            newVal = False
        End If
    Next

    If newVal Then
        For cnt = 0 To Check1.count - 1
            Check1(cnt).Value = 2
        Next
        Check2.Value = 1
    Else
        Check2.Value = 0
    End If
End Function
Private Function CheckboxClicked(ByVal Action As Integer, Optional ByVal Index As Integer)
    If Not CheckboxCancel Then
        CheckboxCancel = True
        
        Dim cnt As Integer
        Select Case Action
            Case 1
                If Check2.Value Then
                    For cnt = 0 To Check1.count - 1
                        If Check1(cnt).enabled Then
                            Check1(cnt).Value = 2
                        Else
                            Check1(cnt).Value = 0
                        End If
                    Next
                Else
                    For cnt = 0 To Check1.count - 1
                        Check1(cnt).Value = 0
                    Next
                End If
            Case 2

                If Check2.Value = 1 Then
                    For cnt = 0 To Check1.count - 1
                        If Not cnt = Index Then Check1(cnt).Value = 1
                    Next
                End If
                    
                    
                If Index = 3 Then
                    For cnt = 4 To 7
                        Check1(cnt).Value = Check1(3).Value
                    Next
                Else
                    If Check1(3).Value = 0 Then
                        For cnt = 4 To 7
                            If Check1(cnt).Value Then Check1(3).Value = Check1(cnt).Value
                        Next
                    End If
                End If
                
                RefreshSelectAll
        End Select
        If (Check1(3).Value = 1) Or (Check1(2).Value = 1) Then
            If Not (Check1(0).Value = 1) Then Check1(0).Value = 1
        End If
        
        CheckboxCancel = False
    End If
    
    RaiseEvent OptionChanged
End Function

Private Sub Check1_Click(Index As Integer)
    CheckboxClicked 2, Index
End Sub

Private Sub Check2_Click()
    CheckboxClicked 1
End Sub

Private Sub UserControl_Initialize()
    Check2.Value = 1
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 1800
    UserControl.Width = 4065
End Sub
