
VERSION 5.00
Begin VB.Form frmEntry 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entry"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   ControlBox      =   0   'False
   Icon            =   "frmEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.ComboBox cmbProduct 
      Height          =   315
      Left            =   1065
      TabIndex        =   10
      Text            =   "cmbProduct"
      Top             =   510
      Width           =   4005
   End
   Begin VB.ComboBox cmbStatus 
      Height          =   315
      Left            =   1065
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3615
      Width           =   4005
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      Left            =   1065
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   900
      Width           =   4005
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      Height          =   315
      Index           =   1
      Left            =   3030
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4050
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   315
      Index           =   0
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4050
      Width           =   975
   End
   Begin VB.TextBox txtComments 
      Height          =   2250
      Left            =   1065
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1290
      Width           =   4005
   End
   Begin VB.TextBox txtDateTime 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1065
      MaxLength       =   255
      TabIndex        =   0
      Top             =   150
      Width           =   4005
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Product:"
      Height          =   210
      Index           =   2
      Left            =   105
      TabIndex        =   11
      Top             =   570
      Width           =   885
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      Height          =   210
      Index           =   1
      Left            =   105
      TabIndex        =   8
      Top             =   3675
      Width           =   885
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Type:"
      Height          =   210
      Index           =   0
      Left            =   105
      TabIndex        =   6
      Top             =   960
      Width           =   885
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comments:"
      Height          =   210
      Left            =   105
      TabIndex        =   5
      Top             =   1305
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date/Time:"
      Height          =   210
      Left            =   105
      TabIndex        =   4
      Top             =   180
      Width           =   885
   End
End
Attribute VB_Name = "frmEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN
Option Compare Text
Public IsOk As Boolean
Public cDateTime As String
Public cProduct As String
Public cType As String
Public cComments As String
Public cStatus As String

Private Sub Action(ByVal aIsOk As Boolean)
    IsOk = aIsOk
    If IsOk Then
        cDateTime = txtDateTime.Text
        cProduct = cmbProduct.Text
        cType = cmbType.List(cmbType.ListIndex)
        cComments = txtComments.Text
        cStatus = cmbStatus.List(cmbStatus.ListIndex)
    Else
        cDateTime = ""
        cProduct = ""
        cType = ""
        cComments = ""
        cStatus = ""
    End If
    Me.Hide
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            Action False
        Case 1
            If Trim(txtComments.Text) <> "" Then
                Action True
            Else
                MsgBox "You must at least type comments in.", vbInformation
            End If
    End Select
End Sub

Private Sub Form_Load()
    
    Dim nProduct
    For Each nProduct In colProducts
        cmbProduct.AddItem nProduct
    Next
    
    cmbType.AddItem "Change"
    cmbType.AddItem "Fix"

    cmbStatus.AddItem "Open"
    cmbStatus.AddItem "Fixed"
    cmbStatus.AddItem "Re-Opened"
    cmbStatus.AddItem "Closed"
            
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Action False
        Cancel = True
    End If
End Sub

Private Sub txtComments_GotFocus()
    If (Len(txtComments.Tag) > 0) And (txtComments.SelStart < Len(txtComments.Tag)) Then
        txtComments.SelStart = Len(txtComments.Tag)
    End If
End Sub

Private Sub txtComments_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 46) Then
        If (Len(txtComments.Tag) > 0) And (txtComments.SelStart < Len(txtComments.Tag)) Then
            KeyCode = 0
        End If
    End If
End Sub

Private Sub txtComments_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 8) Then
        If (Len(txtComments.Tag) > 0) And (txtComments.SelStart <= Len(txtComments.Tag)) Then
            KeyAscii = 0
        End If
    ElseIf (Len(txtComments.Tag) > 0) And (txtComments.SelStart < Len(txtComments.Tag)) Then
        KeyAscii = 0
    End If
End Sub