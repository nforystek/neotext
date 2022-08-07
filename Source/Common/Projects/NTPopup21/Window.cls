VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary

Private pPopup As New frmPopup

Private pPrevWndProc As Long
Private pParentHWnd As Long
Private pParentWindowState As Integer
Private pParentIsActive As Boolean

Private pIcon As Integer
Private pVisible As Boolean
Private pAlwaysOnTop As Boolean

Friend Property Get hWnd() As Long
    hWnd = pPopup.hWnd
End Property

Friend Property Get WindowState() As Long
    WindowState = pPopup.WindowState
End Property

Friend Property Get PrevWndProc() As Long
    PrevWndProc = pPrevWndProc
End Property
Friend Property Let PrevWndProc(ByVal newVal As Long)
    pPrevWndProc = newVal
End Property

Public Property Get AllowInput() As Boolean
    AllowInput = Not pPopup.txtMessage.Locked
End Property
Public Property Let AllowInput(ByVal newVal As Boolean)
    pPopup.txtMessage.Locked = Not newVal
End Property

Public Property Get ParentHWnd() As Long
    ParentHWnd = pParentHWnd
End Property
Public Property Let ParentHWnd(ByVal newVal As Long)
    If pParentHWnd = 0 Then
        pParentHWnd = newVal
        Hook Me
    End If
End Property

Friend Property Get ParentWindowState() As Long
    ParentWindowState = pParentWindowState
End Property
Friend Property Let ParentWindowState(ByVal newVal As Long)
    If (Not pParentWindowState = newVal) Then
        pParentWindowState = newVal
        ParentChanged
    End If
End Property

Friend Property Get ParentIsActive() As Boolean
    ParentIsActive = pParentIsActive
End Property
Friend Property Let ParentIsActive(ByVal newVal As Boolean)
    If (Not pParentIsActive = newVal) Then
        pParentIsActive = newVal
        ParentChanged
    End If
End Property

Friend Sub ParentChanged()
    If pVisible Then
        If pParentHWnd = 0 Then
            If pPopup.WindowState = vbMinimized Then
                pPopup.WindowState = vbNormal
            End If
            pPopup.Visible = True
            If pAlwaysOnTop Then TopMostForm pPopup, True
        Else
            If pParentIsActive And (pParentHWnd = GetActiveWindow) Then
                If pParentWindowState = vbMinimized Then
                    pPopup.Visible = False
                Else
                    If pPopup.WindowState = vbMinimized Then
                        pPopup.WindowState = vbNormal
                    End If
                    pPopup.Visible = True
                    If pAlwaysOnTop Then TopMostForm pPopup, True
                End If
            Else
                pPopup.Visible = False
            End If
        End If
    End If
End Sub

Public Property Get LinkText() As String
    LinkText = pPopup.Label2.Caption
End Property
Public Property Let LinkText(ByVal newVal As String)
    pPopup.Label2.Caption = newVal
End Property

Public Property Get LinkURL() As String
    LinkURL = pPopup.Label2.Tag
End Property
Public Property Let LinkURL(ByVal newVal As String)
    pPopup.Label2.Tag = newVal
End Property

Public Property Get Icon() As Integer
    Icon = pIcon
End Property

Public Property Let Icon(ByVal newVal As Integer)
    If newVal = vbInformation Or newVal = vbQuestion Or newVal = vbCritical Or newVal = vbExclamation Then
        pIcon = newVal
        Set pPopup.Image1.Picture = pPopup.ImageList1.ListImages("k" & newVal).Picture
        Set pPopup.Icon = pPopup.ImageList1.ListImages("k" & newVal).Picture
    ElseIf newVal = 0 Then
        pIcon = 0
        Set pPopup.Image1.Picture = LoadPicture("")
        Set pPopup.Icon = LoadPicture("")
    End If
End Property

Public Property Get Title() As String
    Title = pPopup.Caption
End Property
Public Property Let Title(ByVal newVal As String)
    pPopup.Caption = newVal
End Property

Public Property Get Message() As String
    Message = pPopup.txtMessage.Text
End Property
Public Property Let Message(ByVal newVal As String)
    pPopup.txtMessage.Text = newVal
End Property

Public Property Get AlwaysOnTop() As Boolean
    AlwaysOnTop = pAlwaysOnTop
End Property
Public Property Let AlwaysOnTop(ByVal newVal As Boolean)
    pAlwaysOnTop = newVal
End Property

Public Property Get Visible() As Boolean
    Visible = pVisible
End Property
Public Property Let Visible(ByVal newVal As Boolean)
    If (Not pVisible = newVal) Then
        pVisible = newVal
        ParentChanged
        If (Not pVisible) Then pPopup.Visible = False
    End If
End Property

Public Sub Show()
    pVisible = True
    pPopup.Visible = True
    If pAlwaysOnTop Then TopMostForm pPopup, True
End Sub
Public Sub Hide()
    pVisible = False
    pPopup.Visible = False
End Sub

Private Sub Class_Initialize()

    pParentHWnd = 0
    pParentIsActive = True
    pParentWindowState = vbNormal
    
    Load pPopup
    Set pPopup.MyParent = Me
    
End Sub
Private Sub Class_Terminate()

    If (Not pParentHWnd = 0) Then
        Unhook Me
    End If
    
    If Not (pPopup Is Nothing) Then
        Set pPopup.MyParent = Nothing
        Unload pPopup
    End If
End Sub

