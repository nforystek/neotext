VERSION 5.00
Begin VB.Form frmOverwrite 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "This file already exists.."
   ClientHeight    =   2172
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   7404
   Icon            =   "frmOverwrite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2172
   ScaleWidth      =   7404
   StartUpPosition =   2  'CenterScreen
   Tag             =   "askoverwrite"
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Auto All"
      Height          =   330
      Index           =   6
      Left            =   5205
      TabIndex        =   6
      Top             =   1680
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   330
      Index           =   5
      Left            =   6165
      TabIndex        =   0
      Top             =   1680
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Resume"
      Height          =   330
      Index           =   4
      Left            =   4230
      TabIndex        =   5
      Top             =   1680
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "N&o to All"
      Height          =   330
      Index           =   3
      Left            =   3255
      TabIndex        =   4
      Top             =   1680
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&No"
      Height          =   330
      Index           =   2
      Left            =   2280
      TabIndex        =   3
      Top             =   1680
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Y&es to All"
      Height          =   330
      Index           =   1
      Left            =   1305
      TabIndex        =   2
      Top             =   1680
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Yes"
      Height          =   330
      Index           =   0
      Left            =   330
      TabIndex        =   1
      Top             =   1680
      Width           =   945
   End
   Begin VB.TextBox SourceFile 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   405
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1110
      Width           =   6600
   End
   Begin VB.TextBox DestFile 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   405
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   420
      Width           =   6600
   End
   Begin VB.Label Label2 
      Caption         =   "With this file:"
      Height          =   225
      Left            =   165
      TabIndex        =   10
      Top             =   825
      Width           =   3150
   End
   Begin VB.Label Label1 
      Caption         =   "Do you wish to overwrite this file:"
      Height          =   210
      Left            =   120
      TabIndex        =   9
      Top             =   135
      Width           =   2505
   End
End
Attribute VB_Name = "frmOverwrite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private pPrevWndProc As Long
Private pParentHWnd As Long
Private pParentWindowState As Integer
Private pParentIsActive As Boolean
Private pVisible As Boolean

Public OptionClick As Integer

Public IsOk As Boolean

Public Property Get VisVar() As Boolean
    VisVar = pVisible
End Property
Public Property Let VisVar(ByVal newval As Boolean)
    pVisible = newval
End Property

Public Property Get PrevWndProc() As Long
    PrevWndProc = pPrevWndProc
End Property
Public Property Let PrevWndProc(ByVal newval As Long)
    pPrevWndProc = newval
End Property

Public Property Get ParentHWnd() As Long
    ParentHWnd = pParentHWnd
End Property
Public Property Let ParentHWnd(ByVal newval As Long)
    If pParentHWnd = 0 Then

        pParentHWnd = newval
        Hook Me
        
    End If
End Property

Public Property Get ParentWindowState() As Long
    ParentWindowState = pParentWindowState
End Property
Public Property Let ParentWindowState(ByVal newval As Long)
    If (Not pParentWindowState = newval) Then
        pParentWindowState = newval
        ParentChanged
    End If
End Property

Public Property Get ParentIsActive() As Boolean
    ParentIsActive = pParentIsActive
End Property
Public Property Let ParentIsActive(ByVal newval As Boolean)
    If (Not pParentIsActive = newval) Then
        pParentIsActive = newval
        ParentChanged
    End If
End Property

Friend Sub ParentChanged()
    If pVisible Then
        If pParentHWnd = 0 Then
            If Me.WindowState = vbMinimized Then
                Me.WindowState = vbNormal
            End If
            Me.Visible = True
            TopMostForm Me, True
        Else
            If pParentIsActive Then
                If pParentWindowState = vbMinimized Then
                    Me.Visible = False
                Else
                    If Me.WindowState = vbMinimized Then
                        Me.WindowState = vbNormal
                    End If
                    Me.Visible = True
                    TopMostForm Me, True
                End If
            Else
                Me.Visible = False
            End If
        End If
    End If
End Sub

Private Sub Command1_Click(Index As Integer)

    OptionClick = Index

    IsOk = True

    Me.Visible = False
    pVisible = False
    Unhook Me
    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        OptionClick = or_Cancel
        Cancel = True
    End If
    IsOk = False
    Me.Visible = False
    pVisible = False
    Unhook Me
End Sub

Private Sub Form_Resize()
    ParentChanged
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pVisible = False
    Unhook Me
End Sub
