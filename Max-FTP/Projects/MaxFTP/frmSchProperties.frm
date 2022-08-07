VERSION 5.00
Begin VB.Form frmSchProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Schedule Properties"
   ClientHeight    =   3720
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7560
   Icon            =   "frmSchProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   435
      Left            =   4215
      TabIndex        =   26
      Top             =   1500
      Width           =   3240
      Begin VB.OptionButton Option5 
         Caption         =   "Day"
         Height          =   255
         Left            =   915
         TabIndex        =   29
         Top             =   135
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Hour"
         Height          =   255
         Left            =   1605
         TabIndex        =   28
         Top             =   135
         Width           =   780
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Minute"
         Height          =   255
         Left            =   2355
         TabIndex        =   27
         Top             =   135
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Increment:"
         Height          =   225
         Left            =   75
         TabIndex        =   30
         Top             =   150
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1305
      Left            =   60
      TabIndex        =   18
      Top             =   1935
      Width           =   4050
      Begin VB.VScrollBar VScroll1 
         Height          =   300
         Left            =   2865
         TabIndex        =   22
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Date"
         Height          =   285
         Index           =   2
         Left            =   2895
         TabIndex        =   21
         Top             =   255
         Width           =   570
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1365
         TabIndex        =   20
         Text            =   "12:00 PM"
         Top             =   600
         Width           =   1470
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1365
         TabIndex        =   19
         Top             =   255
         Width           =   1470
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Height          =   270
         Left            =   105
         TabIndex        =   25
         Top             =   960
         Width           =   3795
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Start at Time:"
         Height          =   210
         Left            =   300
         TabIndex        =   24
         Top             =   645
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Start at Date:"
         Height          =   225
         Left            =   285
         TabIndex        =   23
         Top             =   300
         Width           =   990
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1305
      Left            =   4215
      TabIndex        =   12
      Top             =   1935
      Width           =   3255
      Begin VB.TextBox txtInterval 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1845
         TabIndex        =   15
         Top             =   795
         Width           =   450
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Schedule every"
         Height          =   210
         Left            =   375
         TabIndex        =   14
         Top             =   825
         Width           =   1425
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Schedule every day."
         Height          =   195
         Left            =   375
         TabIndex        =   13
         Top             =   510
         Value           =   -1  'True
         Width           =   2355
      End
      Begin VB.Timer Timer1 
         Left            =   0
         Top             =   30
      End
      Begin VB.Label Label3 
         Caption         =   "days"
         Height          =   225
         Left            =   2370
         TabIndex        =   17
         Top             =   825
         Width           =   480
      End
      Begin VB.Label Label6 
         Caption         =   "Increment Value:"
         Height          =   240
         Left            =   105
         TabIndex        =   16
         Top             =   165
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Height          =   435
      Left            =   60
      TabIndex        =   7
      Top             =   1500
      Width           =   4050
      Begin VB.OptionButton Option6 
         Caption         =   "Set Time/Date"
         Height          =   240
         Left            =   1485
         TabIndex        =   10
         Top             =   135
         Value           =   -1  'True
         Width           =   1440
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Increment"
         Height          =   240
         Left            =   2925
         TabIndex        =   9
         Top             =   135
         Width           =   1035
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Manual"
         Height          =   240
         Left            =   555
         TabIndex        =   8
         Top             =   135
         Width           =   885
      End
      Begin VB.Label Label5 
         Caption         =   "Type:"
         Height          =   225
         Left            =   75
         TabIndex        =   11
         Top             =   150
         Width           =   435
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1500
      Left            =   75
      TabIndex        =   4
      Top             =   0
      Width           =   7395
      Begin VB.TextBox Text3 
         Height          =   300
         Left            =   1515
         TabIndex        =   0
         Top             =   315
         Width           =   5550
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Public"
         Height          =   210
         Left            =   315
         TabIndex        =   1
         Top             =   675
         Width           =   1770
      End
      Begin VB.Label Label9 
         Caption         =   "Schedule Name:"
         Height          =   270
         Left            =   285
         TabIndex        =   6
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label8 
         Caption         =   $"frmSchProperties.frx":08CA
         Height          =   450
         Left            =   600
         TabIndex        =   5
         Top             =   960
         Width           =   6525
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cacnel"
      Height          =   345
      Index           =   1
      Left            =   6450
      TabIndex        =   3
      Top             =   3300
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   5295
      TabIndex        =   2
      Top             =   3300
      Width           =   1035
   End
End
Attribute VB_Name = "frmSchProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private Sid As Long
Private OldName As String
Private SchNew As Boolean

Property Get ScheduleType() As Integer
    If Option3.Value Then
        ScheduleType = 0
    ElseIf Option6.Value Then
        ScheduleType = 2
    ElseIf Option4.Value Then
        ScheduleType = 1
    End If
End Property
Property Let ScheduleType(ByVal newval As Integer)
    Select Case newval
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
Property Let ExecuteDate(ByVal newval As String)
    Text2.Text = newval
End Property

Property Get ExecuteTime() As String
    ExecuteTime = Text1.Text
End Property
Property Let ExecuteTime(ByVal newval As String)
    Text1.Text = newval
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
Property Let IncrementType(ByVal newval As Integer)
    Select Case newval
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
Property Let IncrementInterval(ByVal newval As Long)
    If newval = 1 Then
        Option1.Value = True
        Option2.Value = False
        txtInterval.Text = ""
    ElseIf newval > 1 Then
        Option1.Value = False
        Option2.Value = True
        txtInterval.Text = newval
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
            Label7.enabled = False
            Label4.enabled = False
            Label1.enabled = False
            Text2.enabled = False
            Text1.enabled = False
            VScroll1.enabled = False
            Command1(2).enabled = False
            
            Label2.enabled = False
            Option5.enabled = False
            Option7.enabled = False
            Option8.enabled = False
        
            Option1.enabled = False
            Option2.enabled = False
            txtInterval.enabled = False
            Label6.enabled = False
            Label3.enabled = False
            
        Case 1
            Label7.enabled = True
            Label4.enabled = True
            Label1.enabled = True
            Text2.enabled = True
            Text1.enabled = True
            VScroll1.enabled = True
            Command1(2).enabled = True
            
            Label2.enabled = False
            Option5.enabled = False
            Option7.enabled = False
            Option8.enabled = False
        
            Option1.enabled = False
            Option2.enabled = False
            txtInterval.enabled = False
            Label6.enabled = False
            Label3.enabled = False
        Case 2
            Label7.enabled = True
            Label4.enabled = True
            Label1.enabled = True
            Text2.enabled = True
            Text1.enabled = True
            VScroll1.enabled = True
            Command1(2).enabled = True
            
            Label2.enabled = True
            Option5.enabled = True
            Option7.enabled = True
            Option8.enabled = True
        
            Option1.enabled = True
            Option2.enabled = True
            txtInterval.enabled = Option2.Value
            Label6.enabled = True
            Label3.enabled = True
            
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

Public Property Get ID() As Long
    ID = Sid
End Property
Public Sub ShowProperties(ByVal ScheduleID As Long, ByVal IsNew As Boolean)
    Dim dbSchedule As New clsDBSchedule
    
    With dbSchedule
        Sid = ScheduleID
        OldName = .GetScheduleValue(Sid, "ScheduleName")
        SchNew = IsNew
        If IsNew Then
            .SetScheduleValue Sid, "ScheduleName", ""
        End If
        
        Text3.Text = .GetScheduleValue(Sid, "ScheduleName")
        
        Check2.Value = BoolToCheck(.GetScheduleValue(Sid, "IsPublic"))

        ScheduleType = .GetScheduleValue(Sid, "ScheduleType")
        IncrementType = .GetScheduleValue(Sid, "IncrementType")
        IncrementInterval = .GetScheduleValue(Sid, "IncrementInterval")
        
        If Not (ScheduleType = 0) Then
            ExecuteTime = .GetScheduleValue(Sid, "ExecuteTime")
            ExecuteDate = .GetScheduleValue(Sid, "ExecuteDate")
        End If

        Check2.enabled = IIf(dbSettings.CurrentUserAccessRights = ar_Administrator, True, (.GetScheduleOwner(Sid) = dbSettings.GetUserLoginName))

    End With
    
    Set dbSchedule = Nothing
    
    Me.Left = ((Screen.Width / 2) - (Me.Width / 2))
    Me.Top = ((Screen.Height / 2) - (Me.Height / 2))
    Dim frm As Form
    For Each frm In Forms
        If (TypeName(frm) = TypeName(Me)) Then
            If Not (frm.hwnd = Me.hwnd) Then
                Me.Left = frm.Left + (Screen.TwipsPerPixelX * 32)
                Me.Top = frm.Top + (Screen.TwipsPerPixelY * 32)
            End If
        End If
    Next
    
    If ((Me.Left + Me.Width) > Screen.Width) Or _
        ((Me.Top + Me.Height) > Screen.Height) Then
        Me.Left = (32 * Screen.TwipsPerPixelX)
        Me.Top = (32 * Screen.TwipsPerPixelY)
    End If
    
    Me.Show
    
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            
            Dim testInt As Integer
        
            If ScheduleType = 1 Or ScheduleType = 2 Then
            
                If Not IsDate(ExecuteDate) Then
                    MsgBox "You must enter a valid starting date.   MM/DD/YYYY", vbInformation, AppName
                    Exit Sub
                End If
                On Error Resume Next
                testInt = Hour(ExecuteTime)
                testInt = Minute(ExecuteTime)
                If Err > 0 Then
                    Err.Clear
                    MsgBox "You must enter a valid operaion time.   HH:MM PM/AM", vbInformation, AppName
                    On Error GoTo 0
                    Exit Sub
                Else
                    On Error GoTo 0
                End If
                
            End If
            
            Dim dbSchedule As New clsDBSchedule
            
            If Trim(Text3.Text) = "" Then
                MsgBox "You must enter a name for this schedule.", vbInformation, AppName
            Else
                If dbSchedule.ScheduleNameExists(dbSettings.CurrentUserID, Text3.Text) And (Text3.Text <> OldName) Then
                    MsgBox "The schedule name you supplied already exists.  Please choose another.", vbInformation, AppName
                    
                Else
            
                    With dbSchedule

                        .SetScheduleValue Sid, "ScheduleName", Text3.Text
                        .SetScheduleValue Sid, "IsPublic", CBool(Check2.Value)

                        .SetScheduleValue Sid, "ScheduleType", ScheduleType
                        .SetScheduleValue Sid, "IncrementType", IncrementType
                        .SetScheduleValue Sid, "IncrementInterval", IncrementInterval
                        .SetScheduleValue Sid, "ExecuteTime", ExecuteTime
                        .SetScheduleValue Sid, "ExecuteDate", ExecuteDate
                            
                    End With
                    
                    Set dbSchedule = Nothing
                

                    If frmSchManager.Visible = True Then
                        frmSchManager.RefreshSchedules
                    End If

                    If Not (ProcessRunning(ServiceFileName) = 0) Then
                        MessageQueueAdd ServiceFileName, "/loadschedule " & Sid
                    End If
                    
                    Unload Me

                End If
                
            End If
            
            Set dbSchedule = Nothing
        Case 1
            
            If SchNew Then
                
                Dim dbConn As New clsDBConnection
                
                dbConn.dbQuery "DELETE FROM Operations WHERE ParentID=" & Sid & ";"
                dbConn.dbQuery "DELETE FROM Schedules WHERE ID=" & Sid & ";"
                
                Set dbConn = Nothing
            
            End If
            
            Dim frm As Form
            For Each frm In Forms
                If TypeName(frm) = "frmSchManager" Then
                    frm.RefreshSchedules
                End If
            Next
            Unload Me
        Case 2
        
            Dim frmPickDate As New frmDate
            Load frmPickDate
            
            frmPickDate.SelectedDate = Text2.Text
            
            frmPickDate.Show 1
            If frmPickDate.IsOk Then
                Text2.Text = frmPickDate.SelectedDate
            End If
            
            Unload frmPickDate
    End Select
    
End Sub
Private Sub Option1_Click()
    txtInterval.enabled = False
End Sub

Private Sub Option2_Click()
    txtInterval.enabled = True
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

Private Sub Timer1_Timer()
    Label1.Caption = "Current Time: " & Now
End Sub

Private Sub Form_Load()
    Timer1.Interval = 1000
    
    Text2.Text = Date
    Text1.Text = GetTime(time)
    txtInterval.Text = 2
    
    Text1.Tag = Hour(time) & ":" & Minute(time)
    
    VScroll1.Max = -1
    VScroll1.Min = 1
    VScroll1.Value = 0
    
    EnableForm 1
    
    Timer1.enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.enabled = False
End Sub

Private Sub VScroll1_Change()
    Dim newTime As String
    
    newTime = DateAdd("n", VScroll1, Text1.Tag)
    Text1.Tag = Hour(newTime) & ":" & Minute(newTime)
    Text1.Text = GetTime(Text1.Tag)
    
    VScroll1.Value = 0
    Text1.SetFocus
End Sub
