


VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Date"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2700
   ControlBox      =   0   'False
   Icon            =   "frmDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Index           =   1
      Left            =   375
      TabIndex        =   1
      Top             =   2400
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   330
      Index           =   0
      Left            =   1365
      TabIndex        =   2
      Top             =   2400
      Width           =   945
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   48496641
      CurrentDate     =   37838
   End
End
Attribute VB_Name = "frmDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TOP DOWN
Option Explicit
'TOP DOWN

Option Compare Binary

Public IsOk As Boolean

Public Property Get SelectedDate() As Date
    SelectedDate = MonthView1.Month & "/" & MonthView1.Day & "/" & MonthView1.Year
End Property
Public Property Let SelectedDate(ByVal newval As Date)
    MonthView1.Month = DatePart("m", newval)
    MonthView1.Day = DatePart("d", newval)
    MonthView1.Year = DatePart("yyyy", newval)
End Property
Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            IsOk = False
            Me.Hide
        Case 1
            IsOk = True
            Me.Hide
    End Select
End Sub

Private Sub Form_Load()
    Me.Width = 2790
    Me.Height = 3135
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        IsOk = False
        Me.Hide
    End If
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    Command1(1).SetFocus

End Sub

Private Sub MonthView1_DateDblClick(ByVal DateDblClicked As Date)
    Command1(1).SetFocus

End Sub
