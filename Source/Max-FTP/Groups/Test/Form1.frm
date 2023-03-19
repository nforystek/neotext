VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   5160
      Top             =   2160
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Destroy MaxService Object2"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6360
      TabIndex        =   14
      Top             =   4800
      Width           =   2565
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Create MaxService Object1"
      Height          =   375
      Left            =   6240
      TabIndex        =   13
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Call Quit App Object1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Call Quit App Object2"
      Enabled         =   0   'False
      Height          =   420
      Left            =   240
      TabIndex        =   2
      Top             =   5400
      Width           =   2535
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Calle New Window Object2"
      Enabled         =   0   'False
      Height          =   465
      Left            =   240
      TabIndex        =   9
      Top             =   4920
      Width           =   2535
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Destroy Max-FTP Object2"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   5880
      Width           =   2475
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Create Max-FTP Object2"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   4440
      Width           =   2535
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Call New Window Object1"
      Enabled         =   0   'False
      Height          =   345
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Destroy MaxService Object1"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6240
      TabIndex        =   4
      Top             =   480
      Width           =   2565
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Create Max-FTP Object1"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create MaxService Object2"
      Height          =   345
      Left            =   6360
      TabIndex        =   1
      Top             =   4440
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Destroy Max-FTP Object1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   2355
   End
   Begin VB.Label Label7 
      Height          =   855
      Left            =   6360
      TabIndex        =   18
      Top             =   5280
      Width           =   2535
   End
   Begin VB.Label Label6 
      Height          =   1695
      Left            =   2880
      TabIndex        =   17
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Label Label5 
      Height          =   735
      Left            =   6240
      TabIndex        =   16
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label4 
      Height          =   1335
      Left            =   2760
      TabIndex        =   15
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   $"Form1.frx":0000
      Height          =   2535
      Left            =   3480
      TabIndex        =   8
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":00E0
      Height          =   2535
      Left            =   5880
      TabIndex        =   7
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":02DB
      Height          =   2655
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Dim ClientWin1 As MaxFTP.Application
Dim ClientWin2 As MaxFTP.Application
Dim MaxService1 As MaxService.Application
Dim MaxService2 As MaxService.Application

Private Sub SetEnabled()
    
    Command4.Enabled = (ClientWin1 Is Nothing)
    Command6.Enabled = Not (ClientWin1 Is Nothing)
    Command10.Enabled = Not (ClientWin1 Is Nothing)
    Command1.Enabled = Not (ClientWin1 Is Nothing)
    If Not ClientWin1 Is Nothing Then
        On Error Resume Next
        Label4.Caption = "ClientWin1: " & ClientWin1.AppTitle & " ProcessID: " & ClientWin1.AppProcessId & vbCrLf & _
                            "ThreadID: " & ClientWin1.AppThreadId & " hInstance: " & ClientWin1.AppHInstance
        If Err.Number <> 0 Then
            Label4.Caption = Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Else
        Label4.Caption = "ClientWin1 Is Nothing"
    End If
    
    Command8.Enabled = (ClientWin2 Is Nothing)
    Command7.Enabled = Not (ClientWin2 Is Nothing)
    Command3.Enabled = Not (ClientWin2 Is Nothing)
    Command9.Enabled = Not (ClientWin2 Is Nothing)
    If Not ClientWin2 Is Nothing Then
        On Error Resume Next
        Label6.Caption = "ClientWin2: " & ClientWin2.AppTitle & " ProcessID: " & ClientWin2.AppProcessId & vbCrLf & _
                                        "ThreadID: " & ClientWin2.AppThreadId & " hInstance: " & ClientWin2.AppHInstance
        If Err.Number <> 0 Then
            Label4.Caption = Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Else
    
        Label6.Caption = "ClientWin2 Is Nothing"
    End If
    
    Command11.Enabled = (MaxService1 Is Nothing)
    Command5.Enabled = Not (MaxService1 Is Nothing)
    If Not MaxService1 Is Nothing Then
        On Error Resume Next
        Label5.Caption = "MaxService1: " & MaxService1.AppTitle & " ProcessID: " & MaxService1.AppProcessId & vbCrLf & _
                            "ThreadID: " & MaxService1.AppThreadId & " hInstance: " & MaxService1.AppHInstance
        If Err.Number <> 0 Then
            Label4.Caption = Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Else
        Label5.Caption = "MaxService1 Is Nothing"
    End If
    
    Command2.Enabled = (MaxService2 Is Nothing)
    Command12.Enabled = Not (MaxService2 Is Nothing)
    If Not MaxService2 Is Nothing Then
        On Error Resume Next
        Label7.Caption = "MaxService2: " & MaxService2.AppTitle & " ProcessID: " & MaxService2.AppProcessId & vbCrLf & _
                            "ThreadID: " & MaxService2.AppThreadId & " hInstance: " & MaxService2.AppHInstance
        If Err.Number <> 0 Then
            Label4.Caption = Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Else
        Label7.Caption = "MaxService2 Is Nothing"
    End If
End Sub


Private Sub Command1_Click()
    Set ClientWin1 = Nothing
End Sub

Private Sub Command10_Click()
    ClientWin1.QuitApplication
End Sub

Private Sub Command11_Click()
    Set MaxService1 = New MaxService.Application
End Sub

Private Sub Command12_Click()
    Set MaxService2 = Nothing
End Sub

Private Sub Command2_Click()
    Set MaxService2 = New MaxService.Application
End Sub

Private Sub Command3_Click()
    ClientWin2.QuitApplication
End Sub

Private Sub Command4_Click()
    Set ClientWin1 = New MaxFTP.Application
End Sub

Private Sub Command5_Click()
    Set MaxService1 = Nothing
End Sub

Private Sub Command6_Click()
    ClientWin1.WindowCreate
End Sub

Private Sub Command7_Click()
    ClientWin2.WindowCreate
End Sub

Private Sub Command8_Click()
    Set ClientWin2 = New MaxFTP.Application
End Sub

Private Sub Command9_Click()
    Set ClientWin2 = Nothing
End Sub

Private Sub Timer1_Timer()
    SetEnabled
End Sub
