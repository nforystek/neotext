VERSION 5.00
Begin VB.UserControl ctlButton 
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3105
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   207
   ToolboxBitmap   =   "ctlButton.ctx":0000
   Begin VB.Image img 
      Height          =   1050
      Left            =   135
      Top             =   210
      Width           =   2550
   End
End
Attribute VB_Name = "ctlButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private Const LOGPIXELSY = 90

Private Const LF_FACESIZE = 32

Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_SINGLELINE = &H20
Private Const DT_WORDBREAK = &H10
Private Const DT_CALCRECT = &H400
Private Const DT_END_ELLIPSIS = &H8000
Private Const DT_MODIFYSTRING = &H10000
Private Const DT_WORD_ELLIPSIS = &H40000

Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNFACE = 15
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_ACTIVEBORDER = 10
Private Const COLOR_WINDOWFRAME = 6
Private Const COLOR_3DDKSHADOW = 21
Private Const COLOR_3DLIGHT = 22
Private Const COLOR_INFOTEXT = 23
Private Const COLOR_INFOBK = 24

Private Const PATCOPY = &HF00021
Private Const SRCCOPY = &HCC0020

Private Const PS_SOLID = 0
Private Const PS_DASHDOT = 3
Private Const PS_DASHDOTDOT = 4
Private Const PS_DOT = 2
Private Const PS_DASH = 1
Private Const PS_ENDCAP_FLAT = &H200

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function SelectClipPath Lib "gdi32" (ByVal hdc As Long, ByVal iMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long

Enum PicStates
    picNothing = 0
    picDown = 1
    picHover = 2
    picNorm = 3
End Enum

Dim PicState As PicStates

Dim m_NormalPic As Variant
Dim m_HoverPic As Variant
Dim m_DownPic As Variant
Dim m_MouseInside As Boolean

Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseOut()
Public Event Resize()

Private Sub img_Click()
    RaiseEvent Click
End Sub

Private Sub img_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub img_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub img_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
    If Button = vbLeftButton Then
        Exit Sub
    End If
    
    If GetCapture() <> UserControl.hWnd Then
        SetCapture (UserControl.hWnd)
        If Not img.Picture = HoverPic Then
            img.Picture = HoverPic
        End If
    Else
        Dim pt As POINTAPI
        pt.X = X
        pt.Y = Y
        ClientToScreen UserControl.hWnd, pt
        If WindowFromPoint(pt.X, pt.Y) <> UserControl.hWnd Then
            Refresh
            If Button <> vbLeftButton Then
                ReleaseCapture
                img.Picture = NormalPic
                RaiseEvent MouseOut
            End If
        End If
    End If
End Sub

Private Sub img_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    RaiseEvent Click
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    If Not NormalPic Is Nothing Then
        Set img.Picture = NormalPic
    End If
    img.Top = 0
    img.Left = 0
End Sub

Private Sub usercontrol_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    img.Picture = DownPic
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub usercontrol_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
    If Button = vbLeftButton Then
        Exit Sub
    End If
    
    If GetCapture() <> UserControl.hWnd Then
        SetCapture (UserControl.hWnd)
        If Not img.Picture = HoverPic Then
            img.Picture = HoverPic
            m_MouseInside = True
        End If
    Else
        Dim pt As POINTAPI
        pt.X = X
        pt.Y = Y
        ClientToScreen UserControl.hWnd, pt
        If WindowFromPoint(pt.X, pt.Y) <> UserControl.hWnd Then
            Refresh
            If Button <> vbLeftButton Then
                ReleaseCapture
                img.Picture = NormalPic
                m_MouseInside = False
                RaiseEvent MouseOut
            End If
        End If
    End If
End Sub

Private Sub usercontrol_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    img.Picture = HoverPic
    RaiseEvent MouseUp(Button, Shift, X, Y)
    RaiseEvent Click
End Sub

Public Property Get NormalPic() As Variant
    Set NormalPic = m_NormalPic
End Property

Public Property Set NormalPic(vNewPic As Variant)
    Set m_NormalPic = vNewPic
    PropertyChanged "NormalPic"
    img.Picture = NormalPic
End Property

Public Property Get DownPic() As Variant
    Set DownPic = m_DownPic
End Property

Public Property Set DownPic(vNewPic As Variant)
    Set m_DownPic = vNewPic
    PropertyChanged "DownPic"
End Property

Public Property Get HoverPic() As Variant
    Set HoverPic = m_HoverPic
End Property

Public Property Set HoverPic(vNewPic As Variant)
    Set m_HoverPic = vNewPic
    PropertyChanged "HoverPic"
End Property

Private Sub UserControl_Paint()
    img.Picture = NormalPic
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set m_NormalPic = PropBag.ReadProperty("NormalPic", Nothing)
    Set m_DownPic = PropBag.ReadProperty("DownPic", Nothing)
    Set m_HoverPic = PropBag.ReadProperty("HoverPic", Nothing)
    img.Stretch = PropBag.ReadProperty("Stretch", False)
End Sub

Private Sub UserControl_Resize()
    RaiseEvent Resize
    img.Width = UserControl.ScaleWidth
    img.Height = UserControl.ScaleHeight
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "NormalPic", m_NormalPic, 0
    PropBag.WriteProperty "DownPic", m_DownPic, 0
    PropBag.WriteProperty "HoverPic", m_HoverPic, 0
    PropBag.WriteProperty "Stretch", img.Stretch
End Sub

Public Property Get Stretch() As Boolean
    Stretch = img.Stretch
End Property

Public Property Let Stretch(vNewValue As Boolean)
    img.Stretch = vNewValue
    PropertyChanged "Stretch"
End Property

Public Property Get CurPicture() As PicStates
    If img.Picture = 0 Then
        CurPicture = picNothing
    ElseIf img.Picture = NormalPic Then
        CurPicture = picNorm
    ElseIf img.Picture = DownPic Then
        CurPicture = picDown
    ElseIf img.Picture = HoverPic Then
        CurPicture = picHover
    End If
End Property

Public Property Get MouseInside() As Boolean
    MouseInside = m_MouseInside
End Property
