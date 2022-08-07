VERSION 5.00
Begin VB.Form frmWarning 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Warning"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "frmWarning.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Continue  Anyway"
      Height          =   375
      Left            =   4515
      TabIndex        =   0
      Top             =   1500
      Width           =   1665
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   825
      Left            =   135
      TabIndex        =   2
      Top             =   510
      Width           =   6030
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   45
      X2              =   6210
      Y1              =   345
      Y2              =   345
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   60
      X2              =   6225
      Y1              =   330
      Y2              =   330
   End
   Begin VB.Label Label1 
      Caption         =   "The fololowing applications should be shut down before the requested action is taken."
      Height          =   240
      Left            =   90
      TabIndex        =   1
      Top             =   75
      Width           =   6270
   End
End
Attribute VB_Name = "frmWarning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private Sub Command1_Click()
    Me.Hide
End Sub

Private Sub Form_Load()

    TestApps
    
    If Label2.Caption = "" Then
        Me.Hide
    End If
    
End Sub

Public Sub TestApps()
    Dim tmp As String
    
    If (ProcessRunning(RemindMeFileName) > 0) Then
        tmp = tmp & RemindMeFileName & vbCrLf
    End If
    
    If (ProcessRunning(ServiceFileName) > 0) Then
        tmp = tmp & ServiceFileName & vbCrLf
    End If
           
    Label2.Caption = tmp
    
    If Label2.Caption = "" Then
        Me.Hide
    End If
    
End Sub
