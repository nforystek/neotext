VERSION 5.00
Begin VB.UserControl ScheduleProperties 
   ClientHeight    =   1770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7515
   LockControls    =   -1  'True
   ScaleHeight     =   1770
   ScaleWidth      =   7515
   ToolboxBitmap   =   "ScheduleProperties.ctx":0000
   Begin VB.Frame Frame4 
      Height          =   435
      Left            =   75
      TabIndex        =   21
      Top             =   -30
      Width           =   4050
      Begin VB.OptionButton Option3 
         Caption         =   "Manual"
         Height          =   240
         Left            =   555
         TabIndex        =   0
         Top             =   135
         Width           =   885
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Increment"
         Height          =   240
         Left            =   2925
         TabIndex        =   2
         Top             =   135
         Width           =   1035
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Set Time/Date"
         Height          =   240
         Left            =   1485
         TabIndex        =   1
         Top             =   135
         Value           =   -1  'True
         Width           =   1440
      End
      Begin VB.Label Label5 
         Caption         =   "Type:"
         Height          =   225
         Left            =   75
         TabIndex        =   22
         Top             =   150
         Width           =   435
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1305
      Left            =   4230
      TabIndex        =   19
      Top             =   405
      Width           =   3255
      Begin VB.OptionButton Option1 
         Caption         =   "Schedule every day."
         Height          =   195
         Left            =   375
         TabIndex        =   10
         Top             =   510
         Value           =   -1  'True
         Width           =   2355
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Schedule every"
         Height          =   210
         Left            =   375
         TabIndex        =   11
         Top             =   825
         Width           =   1425
      End
      Begin VB.TextBox txtInterval 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1845
         TabIndex        =   12
         Top             =   795
         Width           =   450
      End
      Begin VB.Label Label6 
         Caption         =   "Increment Value:"
         Height          =   240
         Left            =   105
         TabIndex        =   23
         Top             =   165
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "days"
         Height          =   225
         Left            =   2370
         TabIndex        =   20
         Top             =   825
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1305
      Left            =   75
      TabIndex        =   15
      Top             =   405
      Width           =   4050
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1365
         TabIndex        =   6
         Top             =   255
         Width           =   1470
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1365
         TabIndex        =   8
         Text            =   "12:00 PM"
         Top             =   600
         Width           =   1470
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Date"
         Height          =   285
         Left            =   2895
         TabIndex        =   7
         Top             =   255
         Width           =   570
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   300
         Left            =   2865
         TabIndex        =   9
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Start at Date:"
         Height          =   225
         Left            =   285
         TabIndex        =   18
         Top             =   300
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Start at Time:"
         Height          =   210
         Left            =   300
         TabIndex        =   17
         Top             =   645
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Height          =   270
         Left            =   105
         TabIndex        =   16
         Top             =   960
         Width           =   3795
      End
   End
   Begin VB.Frame Frame3 
      Height          =   435
      Left            =   4230
      TabIndex        =   13
      Top             =   -30
      Width           =   3240
      Begin VB.OptionButton Option8 
         Caption         =   "Minute"
         Height          =   255
         Left            =   2355
         TabIndex        =   5
         Top             =   135
         Width           =   855
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Hour"
         Height          =   255
         Left            =   1605
         TabIndex        =   4
         Top             =   135
         Width           =   780
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Day"
         Height          =   255
         Left            =   915
         TabIndex        =   3
         Top             =   135
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Increment:"
         Height          =   225
         Left            =   75
         TabIndex        =   14
         Top             =   150
         Width           =   765
      End
   End
End
Attribute VB_Name = "ScheduleProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Private WithEvents Timer1 As NTSchedule20.Timer
Attribute Timer1.VB_VarHelpID = -1


Property Get ScheduleType() As Integer
    If Option3.Value Then
        ScheduleType = 0
    ElseIf Option6.Value Then
        ScheduleType = 2
    ElseIf Option4.Value Then
        ScheduleType = 1
    End If
End Property
Property Let ScheduleType(ByVal newVal As Integer)
    Select Case newVal
        Case 0
            Option3.Value = True
            Option4.Value = False
            Option5.Value = False
        Case 1
            Option3.Value = False
            Option4.Value = True
            Option6.Value = False
        Case 2
            Option3.Value = False
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
        If IsNumeric(txtInterval.Text) Then
            If CInt(txtInterval.Text) < 2 Then
                IncrementInterval = 2
            Else
                IncrementInterval = IIf((CInt(txtInterval.Text) > 1), txtInterval.Text, 2)
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
        txtInterval.Text = ""
    ElseIf newVal > 1 Then
        Option1.Value = False
        Option2.Value = True
        txtInterval.Text = newVal
    End If
End Property



Private Function SetFormText(ByVal nText As Integer)
    Select Case nText
        Case 1
            Option1.Caption = "Schedule every day."
            Label3.Caption = "days"
        Case 2
            Option1.Caption = "Schedule every hour."
            Label3.Caption = "hours"
        Case 3
            Option1.Caption = "Schedule every minute."
            Label3.Caption = "minutes"
    End Select
End Function
    
    

Private Function EnableForm(ByVal nEnable As Long)
    Select Case nEnable
        Case 0
            Label7.Enabled = False
            Label4.Enabled = False
            Label1.Enabled = False
            Text2.Enabled = False
            Text1.Enabled = False
            VScroll1.Enabled = False
            Command1.Enabled = False
            
            
            Label2.Enabled = False
            Option5.Enabled = False
            Option7.Enabled = False
            Option8.Enabled = False
        
            Option1.Enabled = False
            Option2.Enabled = False
            txtInterval.Enabled = False
            Label6.Enabled = False
            Label3.Enabled = False
            
        Case 1
            Label7.Enabled = True
            Label4.Enabled = True
            Label1.Enabled = True
            Text2.Enabled = True
            Text1.Enabled = True
            VScroll1.Enabled = True
            Command1.Enabled = True
            
            Label2.Enabled = False
            Option5.Enabled = False
            Option7.Enabled = False
            Option8.Enabled = False
        
            Option1.Enabled = False
            Option2.Enabled = False
            txtInterval.Enabled = False
            Label6.Enabled = False
            Label3.Enabled = False
        Case 2
            Label7.Enabled = True
            Label4.Enabled = True
            Label1.Enabled = True
            Text2.Enabled = True
            Text1.Enabled = True
            VScroll1.Enabled = True
            Command1.Enabled = True
            
            Label2.Enabled = True
            Option5.Enabled = True
            Option7.Enabled = True
            Option8.Enabled = True
        
            Option1.Enabled = True
            Option2.Enabled = True
            txtInterval.Enabled = Option2.Value
            Label6.Enabled = True
            Label3.Enabled = True
            
    End Select
        
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
    Dim frmPickDate As New frmDate
    Load frmPickDate
    
    frmPickDate.SelectedDate = Text2.Text
    
    frmPickDate.Show 1
    If frmPickDate.IsOk Then
        Text2.Text = frmPickDate.SelectedDate
    End If
    
    Unload frmPickDate
End Sub

Private Sub Option1_Click()
    txtInterval.Enabled = False
End Sub

Private Sub Option2_Click()
    txtInterval.Enabled = True
End Sub

Private Sub Option3_Click()
    EnableForm 0
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


Private Sub Timer1_OnTicking()
    Label1.Caption = "Current Time: " & Now
End Sub


Private Sub UserControl_Initialize()
    Set Timer1 = New NTSchedule20.Timer
    Timer1.Interval = 1000
    Timer1.Enabled = True
    
    Text2.Text = Date
    Text1.Text = GetTime(Time)
    txtInterval.Text = 2
    
    Text1.Tag = Hour(Time) & ":" & Minute(Time)
    
    VScroll1.Max = -1
    VScroll1.Min = 1
    VScroll1.Value = 0
    
    EnableForm 1
End Sub


Private Sub UserControl_Resize()
    UserControl.Width = 7515
    UserControl.Height = 1770
End Sub

Private Sub UserControl_Terminate()
    Timer1.Enabled = False
    Set Timer1 = Nothing
End Sub

Private Sub VScroll1_Change()
    Dim newTime As String
    
    newTime = DateAdd("n", VScroll1, Text1.Tag)
    Text1.Tag = Hour(newTime) & ":" & Minute(newTime)
    Text1.Text = GetTime(Text1.Tag)
    
    VScroll1.Value = 0
    Text1.SetFocus
End Sub
