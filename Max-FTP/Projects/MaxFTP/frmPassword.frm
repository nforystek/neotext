VERSION 5.00
Object = "{C98B112F-745F-4542-B5B3-DDFADF1F6E2F}#1021.0#0"; "NTControls22.ocx"
Begin VB.Form frmPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Username and Password Required"
   ClientHeight    =   1704
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7452
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1704
   ScaleWidth      =   7452
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin NTControls22.SiteInformation sInfo 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   2985
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   345
      Index           =   1
      Left            =   6075
      TabIndex        =   2
      Top             =   480
      Width           =   1305
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   6075
      TabIndex        =   1
      Top             =   90
      Width           =   1305
   End
End
Attribute VB_Name = "frmPassword"
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
    frmMain.ValidDataPortRange sInfo.sPortRange
    Select Case Index
        Case 0
            SaveCache sInfo
            IsOk = True
        Case 1
            IsOk = False
    End Select
    Me.Visible = False
    pVisible = False
    Unhook Me
End Sub

Private Sub Form_Activate()
    PreSetup
End Sub

Private Sub PreSetup()
    SetAutoTypeList Me, sInfo.AutoTypeCombo
    If dbSettings.RemoveProfile Then
        sInfo.sSavePass.Value = 1
        sInfo.sSavePass.Visible = False
    End If
    sInfo.ShowAdvSettings = dbSettings.GetProfileSetting("ShowAdvSettings")
    If Not sInfo.ShowAdvSettings Then
        sInfo.sPassive.Value = IIf((dbSettings.GetProfileSetting("ConnectionMode") = 0), 1, 0)
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
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
    Unhook Me
End Sub
