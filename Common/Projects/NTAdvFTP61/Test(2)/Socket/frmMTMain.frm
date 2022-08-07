VERSION 5.00
Begin VB.Form frmMTMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmMTMain"
   ClientHeight    =   1200
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   2535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1080
      Top             =   480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Server"
      Height          =   492
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1092
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Client"
      Height          =   492
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   1092
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   252
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2052
   End
End
Attribute VB_Name = "frmMTMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Code for the form frmMTMain.
Public MainApp As MainApp

Private Sub Command1_Click()
    Debug.Print "SERVER"
    Dim tw As ThreadedWindow
    If IsCompiled Then
        Set tw = CreateObject("ThreadDemo.ThreadedWindow")
    Else
        Set tw = New ThreadedWindow
    End If
   ' Tell the new object to show its form, and
   '   pass it a reference to the main
   '   application object.
   Call tw.Initialize(Me.MainApp, True)
End Sub

Private Sub Command2_Click()
Debug.Print "CLIENT"
   Dim tw As ThreadedWindow
    If IsCompiled Then
        Set tw = CreateObject("ThreadDemo.ThreadedWindow")
    Else
        Set tw = New ThreadedWindow
    End If
   ' Tell the new object to show its form, and
   '   pass it a reference to the main
   '   application object.
   Call tw.Initialize(Me.MainApp, False)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call MainApp.Closing
   Set MainApp = Nothing
End Sub

Private Sub Timer1_Timer()
    Label1.Caption = Forms.count
End Sub
