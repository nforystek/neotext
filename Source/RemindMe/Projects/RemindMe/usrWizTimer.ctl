VERSION 5.00
Begin VB.UserControl usrWizTimer 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   LockControls    =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "usrWizTimer.ctx":0000
   Begin VB.CheckBox Check1 
      Caption         =   "Enabled"
      Height          =   240
      Left            =   330
      TabIndex        =   0
      Top             =   90
      Value           =   1  'Checked
      Width           =   945
   End
   Begin VB.Frame Frame3 
      Height          =   435
      Left            =   300
      TabIndex        =   18
      Top             =   390
      Width           =   4170
      Begin VB.OptionButton Option5 
         Caption         =   "Day"
         Height          =   255
         Left            =   1545
         TabIndex        =   3
         Top             =   135
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Hour"
         Height          =   255
         Left            =   2235
         TabIndex        =   4
         Top             =   135
         Width           =   780
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Minute"
         Height          =   255
         Left            =   2985
         TabIndex        =   5
         Top             =   135
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Increment Type:"
         Height          =   225
         Left            =   135
         TabIndex        =   19
         Top             =   150
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Starting Date and Time"
      Height          =   1320
      Left            =   300
      TabIndex        =   14
      Top             =   1995
      Width           =   4185
      Begin VB.VScrollBar VScroll1 
         Height          =   300
         Left            =   3495
         TabIndex        =   12
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Date"
         Height          =   285
         Left            =   3525
         TabIndex        =   10
         Top             =   255
         Width           =   570
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1995
         TabIndex        =   11
         Text            =   "12:00 PM"
         Top             =   600
         Width           =   1470
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1995
         TabIndex        =   9
         Top             =   255
         Width           =   1470
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Height          =   270
         Left            =   105
         TabIndex        =   17
         Top             =   960
         Width           =   3945
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Perform operation at:"
         Height          =   210
         Left            =   195
         TabIndex        =   16
         Top             =   645
         Width           =   1725
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Starting Date:"
         Height          =   225
         Left            =   195
         TabIndex        =   15
         Top             =   300
         Width           =   1710
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4395
      Top             =   90
   End
   Begin VB.OptionButton Option6 
      Caption         =   "Set Schedule"
      Height          =   240
      Left            =   1365
      TabIndex        =   1
      Top             =   90
      Value           =   -1  'True
      Width           =   1290
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Increment Schedule"
      Height          =   240
      Left            =   2715
      TabIndex        =   2
      Top             =   90
      Width           =   1785
   End
   Begin VB.Frame Frame2 
      Caption         =   "Day Increment"
      Height          =   990
      Left            =   300
      TabIndex        =   13
      Top             =   915
      Width           =   4170
      Begin VB.TextBox txtDays 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2430
         TabIndex        =   8
         Top             =   570
         Width           =   465
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Perform operation every"
         Height          =   210
         Left            =   390
         TabIndex        =   7
         Top             =   600
         Width           =   2085
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Perform operation every day."
         Height          =   195
         Left            =   390
         TabIndex        =   6
         Top             =   300
         Value           =   -1  'True
         Width           =   2805
      End
      Begin VB.Label Label3 
         Caption         =   "days"
         Height          =   225
         Left            =   2985
         TabIndex        =   20
         Top             =   615
         Width           =   885
      End
   End
End
Attribute VB_Name = "usrWizTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public Property Get Enabled() As Boolean
    Enabled = (Check1.Value = 1)
End Property
Public Property Let Enabled(ByVal newVal As Boolean)
    Check1.Value = Abs(CInt(newVal))
End Property
Property Get ScheduleType() As Integer
    If Option6.Value Then
        ScheduleType = 2
    ElseIf Option4.Value Then
        ScheduleType = 1
    End If
End Property
Property Let ScheduleType(ByVal newVal As Integer)
    Select Case newVal
        Case 1
            Option4.Value = True
            Option6.Value = False
        Case 2
            Option4.Value = False
            Option6.Value = True
    End Select
End Property

Property Get ExecuteDate() As String
    ExecuteDate = Text2.Text
End Property
Property Let ExecuteDate(ByVal newVal As String)
    Text2.Text = newVal
End Property

Property Get ExecuteTime() As String
    ExecuteTime = Text1.Text
End Property
Property Let ExecuteTime(ByVal newVal As String)
    Text1.Text = newVal
End Property

Property Get IncrementType() As Integer
    If Option5.Value Then
        IncrementType = 2
    ElseIf Option7.Value Then
        IncrementType = 1
    ElseIf Option8.Value Then
        IncrementType = 0
    End If
End Property
Property Let IncrementType(ByVal newVal As Integer)
    Select Case newVal
        Case 0
            Option5.Value = False
            Option7.Value = False
            Option8.Value = True
        Case 1
            Option8.Value = False
            Option5.Value = False
            Option7.Value = True
        Case 2
            Option7.Value = False
            Option8.Value = False
            Option5.Value = True
    End Select
End Property

Property Get IncrementInterval() As Long
    If Option1.Value Then
        IncrementInterval = 1
    ElseIf Option2.Value Then
        If IsNumeric(txtDays.Text) Then
            If CInt(txtDays.Text) < 2 Then
                IncrementInterval = 2
            Else
                IncrementInterval = IIf((CInt(txtDays.Text) > 1), txtDays.Text, 2)
            End If
        Else
            IncrementInterval = 2
        End If
    End If
End Property
Property Let IncrementInterval(ByVal newVal As Long)
    If newVal = 1 Then
        Option1.Value = True
        Option2.Value = False
        txtDays.Text = ""
    ElseIf newVal > 1 Then
        Option1.Value = False
        Option2.Value = True
        txtDays.Text = newVal
    End If
End Property

Private Function SetFormText(ByVal nText As String)
    Select Case nText
        Case 1
            Option1.Caption = "Perform operation every day."
            Label3.Caption = "days"
        Case 2
            Option1.Caption = "Perform operation every hour."
            Label3.Caption = "hours"
        Case 3
            Option1.Caption = "Perform operation every minute."
            Label3.Caption = "minutes"
    End Select
End Function

Private Function EnableForm(ByVal nEnable As Long)
    Select Case nEnable
        Case 1
            
            Label2.Enabled = False
            Option5.Enabled = False
            Option7.Enabled = False
            Option8.Enabled = False
        
            Option1.Enabled = False
            Option2.Enabled = False
            txtDays.Enabled = False
        Case 2
            
            Label2.Enabled = True
            Option5.Enabled = True
            Option7.Enabled = True
            Option8.Enabled = True
        
            Option1.Enabled = True
            Option2.Enabled = True
            txtDays.Enabled = True
            
    End Select
        
    txtDays.Enabled = Option2.Value
End Function
Private Function GetTime(ByVal nTime As String) As String
    Dim nMinute As String
    Dim nHour As String
    Dim nM As String
    
    If Len(CStr(Minute(nTime))) = 1 Then
        nMinute = "0" & Minute(nTime)
    Else
        nMinute = Minute(nTime)
    End If
    nHour = Hour(nTime)
    If CInt(nHour) > 12 Then
        nHour = CStr(CInt(nHour) - 12)
        nM = " PM"
    Else
        nM = " AM"
    End If
    If nHour = "0" Then nHour = "12"
    
    GetTime = nHour & ":" & nMinute & nM
End Function
Private Sub Command1_Click()
    Load frmDate
    frmDate.Show 1
    If frmDate.IsOk Then
        Text2.Text = frmDate.MonthView1.Month & "/" & frmDate.MonthView1.Day & "/" & frmDate.MonthView1.Year
    End If
End Sub

Private Sub Option1_Click()
    txtDays.Enabled = False
End Sub

Private Sub Option2_Click()
    txtDays.Enabled = True
End Sub

Private Sub Option4_Click()
    EnableForm 2
End Sub

Private Sub Option5_Click()
    SetFormText 1
End Sub

Private Sub Option6_Click()
    EnableForm 1
End Sub

Private Sub Option7_Click()
    SetFormText 2
End Sub

Private Sub Option8_Click()
    SetFormText 3
End Sub

Private Sub Timer1_Timer()
    Label1.Caption = "Current Time: " & Now
End Sub

Private Sub UserControl_Initialize()
    Text2.Text = Date
    Text1.Text = GetTime(time)
    
    txtDays.Text = 2
        
    Text1.Tag = Hour(time) & ":" & Minute(time)
    
    VScroll1.Max = -1
    VScroll1.Min = 1
    VScroll1.Value = 0
    
    EnableForm 1
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 3600
    UserControl.Width = 4800

End Sub

Private Sub VScroll1_Change()
    Dim newTime As String
    
    newTime = DateAdd("n", VScroll1, Text1.Tag)
    Text1.Tag = Hour(newTime) & ":" & Minute(newTime)
    Text1.Text = GetTime(Text1.Tag)
    
    VScroll1.Value = 0
    Text1.SetFocus
End Sub
