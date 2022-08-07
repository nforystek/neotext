VERSION 5.00
Begin VB.UserControl ctlDragger 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ctlDragger.ctx":0000
End
Attribute VB_Name = "ctlDragger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const BDR_SUNKENINNER = &H8
Private Const BF_LEFT As Long = &H1
Private Const BF_TOP As Long = &H2
Private Const BF_RIGHT As Long = &H4
Private Const BF_BOTTOM As Long = &H8
Private Const BF_RECT As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BDR_RAISED = &H5
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_TOOLWINDOW = &H80
Private Const VK_LBUTTON = &H1
Private Const PS_SOLID = 0
Private Const R2_NOTXORPEN = 10
Private Const BLACK_PEN = 7
Private Const SM_CYCAPTION = 4

Private Const SM_CYSMCAPTION = 51

Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As Any) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetParent Lib "user32" (ByVal HwndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long

#If UNICODE Then
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hwnd As Long, ByVal uMgs As Long, ByVal wParam As Long, lParam As Any) As Long
#Else
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal uMgs As Long, ByVal wParam As Long, lParam As Any) As Long
#End If

Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal lLeft As Long, ByVal lTop As Long) As Long
Private Declare Function SetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long

Private Const SWW_HPARENT = -8
Private Const HTRIGHT = 11
Private Const HTLEFT = 10
Private Const HTBOTTOM = 12
Private Const HTTOP = 15

Private mdifrm As MDIForm
Private WithEvents picbox As PictureBox
Attribute picbox.VB_VarHelpID = -1
Private WithEvents picbar As PictureBox
Attribute picbar.VB_VarHelpID = -1
Private WithEvents frm As Form
Attribute frm.VB_VarHelpID = -1

Private bMoving As Boolean
Private iAlign As Integer

Private lFloatingWidth As Long
Private lFloatingHeight As Long
Private lFloatingLeft As Long
Private lFloatingTop As Long

Private bDocked As Boolean
Private bDockable As Boolean
Private bResizable As Boolean
Private bMovable As Boolean

Private lDockedWidth As Long
Private lDockedHeight As Long

Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Resize(Left As Long, Top As Long, Width As Long, ByVal Height As Long)

Private Const m_def_RepositionForm = True
Private Const m_def_Caption = ""

Private m_RepositionForm As Boolean
Private m_Caption As String

Public Property Get Movable() As Boolean
    Movable = bMovable
End Property
Public Property Let Movable(ByVal newval As Boolean)
    bMovable = newval
End Property

Public Property Get Resizable() As Boolean
    Resizable = bResizable
End Property
Public Property Let Resizable(ByVal newval As Boolean)
    bResizable = newval
End Property

Public Property Get Dockable() As Boolean
    Dockable = bDockable
End Property
Public Property Let Dockable(ByVal newval As Boolean)
    bDockable = newval
End Property

Public Property Get Docked() As Boolean
    Docked = bDocked
End Property
Public Property Let Docked(ByVal newval As Boolean)
    bDocked = newval
End Property

Public Property Get DockedWidth() As Long
    DockedWidth = lDockedWidth - (8 * Screen.TwipsPerPixelX)
End Property
Public Property Let DockedWidth(ByVal newval As Long)
    lDockedWidth = newval
End Property
Public Property Get DockedHeight() As Long
    DockedHeight = lDockedHeight - (8 * Screen.TwipsPerPixelY)
End Property
Public Property Let DockedHeight(ByVal newval As Long)
    lDockedHeight = newval
End Property

Public Property Get FloatingTop() As Long
    FloatingTop = lFloatingTop
End Property
Public Property Let FloatingTop(ByVal newval As Long)
    lFloatingTop = newval
End Property
Public Property Get FloatingLeft() As Long
    FloatingLeft = lFloatingLeft
End Property
Public Property Let FloatingLeft(ByVal newval As Long)
    lFloatingLeft = newval
End Property
Public Property Get FloatingWidth() As Long
    FloatingWidth = lFloatingWidth
End Property
Public Property Let FloatingWidth(ByVal newval As Long)
    lFloatingWidth = newval
End Property
Public Property Get FloatingHeight() As Long
    FloatingHeight = lFloatingHeight
End Property
Public Property Let FloatingHeight(ByVal newval As Long)
    lFloatingHeight = newval
End Property

Public Sub SetupDockedForm(ByRef MDIParent As MDIForm, ByRef DockedForm As Form, ByVal DockAlign As Integer)
    
    On Error GoTo setuperr
    bMoving = True
    
    iAlign = DockAlign
    Set frm = DockedForm
    Set mdifrm = MDIParent

    Set picbar = mdifrm.Controls.Add("VB.PictureBox", "picbar" & frm.hwnd)
    Set picbox = mdifrm.Controls.Add("VB.PictureBox", "picbox" & frm.hwnd)
    picbar.AutoRedraw = True
    picbar.BorderStyle = 0
    picbar.Appearance = 0
    
    picbox.AutoRedraw = True
    picbox.BorderStyle = 0
    picbox.Appearance = 0
    
    picbar.Height = (3 * Screen.TwipsPerPixelY)
    picbar.Width = (3 * Screen.TwipsPerPixelX)
    
    picbar.BackColor = frm.BackColor
    picbox.BackColor = frm.BackColor
    
    bMoving = True
    picbox.Height = lDockedHeight
    picbox.Width = lDockedWidth
    bMoving = False

    picbox.Align = iAlign
    picbar.Align = iAlign
   
    Select Case iAlign
        Case 1
            picbar.MousePointer = 7
            picbar.Visible = True
            picbox.Visible = True
        Case 2
            picbar.MousePointer = 7
            picbox.Visible = True
            picbar.Visible = True
        Case 3
            picbar.MousePointer = 9
            picbar.Visible = True
            picbox.Visible = True
        Case 4
            picbar.MousePointer = 9
            picbox.Visible = True
            picbar.Visible = True
    End Select
   
    lDockedWidth = picbox.ScaleWidth + (8 * Screen.TwipsPerPixelX)
    lDockedHeight = picbox.ScaleHeight + (8 * Screen.TwipsPerPixelY)

    If Not bDocked Then

        frm.Visible = False
        bDocked = False
        SetParent frm.hwnd, 0
        frm.Move lFloatingLeft, lFloatingTop, lFloatingWidth, lFloatingHeight
        SetVisible False

        Call SetWindowWord(frm.hwnd, SWW_HPARENT, mdifrm.hwnd)
    Else

        bDocked = True
        SetParent frm.hwnd, picbox.hwnd
        frm.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, lDockedWidth, lDockedHeight
        SetVisible True
    End If

    bMoving = False
    On Error GoTo 0
    Exit Sub
setuperr:
    Err.Clear
    On Error GoTo 0
End Sub
Public Sub SetVisible(ByVal IsVisible As Boolean)
    picbar.Visible = False
    picbox.Visible = False
    Select Case iAlign
        Case 1, 3
            picbar.Visible = IsVisible And bDocked
            picbox.Visible = IsVisible And bDocked
        Case 2, 4
            picbox.Visible = IsVisible And bDocked
            picbar.Visible = IsVisible And bDocked
    End Select
End Sub
Private Sub frm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SetWindowWord(frm.hwnd, SWW_HPARENT, 0&)
End Sub

Private Sub frm_Resize()
    If (Not bResizable) And (Not bDocked) Then
        frm.Move lFloatingLeft, lFloatingTop, lFloatingWidth, lFloatingHeight
    ElseIf bDocked Then
    
        If (frm.WindowState <> vbMinimized) Then
            StoreFormDimensions
    
            RaiseEvent Resize(3 * Screen.TwipsPerPixelX, UserControl.Height + (3 * Screen.TwipsPerPixelY), frm.ScaleWidth - (7 * Screen.TwipsPerPixelX), frm.ScaleHeight - (UserControl.Height + (6 * Screen.TwipsPerPixelY)))
        End If
    End If
End Sub

Private Sub FormDropped(FormLeft As Long, FormTop As Long, FormWidth As Long, FormHeight As Long)
    If (Not bMovable) Then Exit Sub
    
    Dim rct As RECT

    GetWindowRect picbox.hwnd, rct

    With rct
        .Left = .Left - 4
        .Top = .Top - 4
        .Right = .Right + 4
        .Bottom = .Bottom + 4
    End With

    If PtInRect(rct, FormLeft, FormTop) Then
        bDocked = True
        SetParent frm.hwnd, picbox.hwnd
        frm.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, lDockedWidth, lDockedHeight
        SetVisible True
    Else
        frm.Visible = False
        bDocked = False
        SetParent frm.hwnd, 0
        frm.Move FormLeft * Screen.TwipsPerPixelX, FormTop * Screen.TwipsPerPixelY, lFloatingWidth, lFloatingHeight
        SetVisible False
        frm.Visible = True

        Call SetWindowWord(frm.hwnd, SWW_HPARENT, mdifrm.hwnd)
    End If

    bMoving = False
    StoreFormDimensions

End Sub

Private Sub FormMoved(FormLeft As Long, FormTop As Long, FormWidth As Long, FormHeight As Long)
    If (Not bMovable) Then Exit Sub
    
    Dim rct As RECT
    
    bMoving = True
    
    GetWindowRect picbox.hwnd, rct
    With rct
        .Left = .Left - 4
        .Top = .Top - 4
        .Right = .Right + 4
        .Bottom = .Bottom + 4
    End With
    
    If PtInRect(rct, FormLeft, FormTop) Then
        FormWidth = lDockedWidth / Screen.TwipsPerPixelX
        FormHeight = lDockedHeight / Screen.TwipsPerPixelY
    Else
        FormWidth = lFloatingWidth / Screen.TwipsPerPixelX
        FormHeight = lFloatingHeight / Screen.TwipsPerPixelY
    End If

End Sub

Private Sub picbox_Resize()
     On Error Resume Next
    If picbox.Width < 120 Then
        picbox.Width = 120
    End If
    If bDocked Then
        frm.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, picbox.ScaleWidth + (8 * Screen.TwipsPerPixelX), picbox.ScaleHeight + (8 * Screen.TwipsPerPixelY)
    End If
    If Err Then Err.Clear
End Sub

Private Sub StoreFormDimensions()

    If Not bMoving Then
        If bDocked Then
            lDockedWidth = frm.Width
            lDockedHeight = frm.Height
        Else
            lFloatingLeft = frm.Left
            lFloatingTop = frm.Top
            lFloatingWidth = frm.Width
            lFloatingHeight = frm.Height
        End If
    End If
End Sub

Private Sub picbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If picbox.Visible And bResizable Then
    
        ReleaseCapture
        
        Select Case iAlign
            Case 1
                SendMessage picbox.hwnd, WM_NCLBUTTONDOWN, HTTOP, ByVal 0&
            Case 2
                SendMessage picbox.hwnd, WM_NCLBUTTONDOWN, HTBOTTOM, ByVal 0&
            Case 3
                SendMessage picbox.hwnd, WM_NCLBUTTONDOWN, HTRIGHT, ByVal 0&
            Case 4
                SendMessage picbox.hwnd, WM_NCLBUTTONDOWN, HTLEFT, ByVal 0&
        End Select
            
    End If
End Sub

Private Sub UserControl_Initialize()
    m_Caption = m_def_Caption
    m_RepositionForm = m_def_RepositionForm
    bDockable = True
    bMovable = True
    bResizable = True
    bDocked = True
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If (Not Resizable) And (Not bMovable) Then Exit Sub
    
    Dim na As Long
    Dim pt As POINTAPI
    Dim frmHWnd As Long
    
    UserControl_Paint
    frmHWnd = UserControl.Extender.Parent.hwnd
    
    If Button = vbLeftButton And X >= 0 And X <= UserControl.ScaleWidth And Y >= 0 And Y <= UserControl.ScaleHeight Then
        ReleaseCapture
        DragObject frmHWnd
    End If

    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub DragObject(ByVal hwnd As Long)
    If bDocked Or Not bMovable Then Exit Sub
    
    Dim pt As POINTAPI
    Dim ptPrev As POINTAPI
    Dim objRect As RECT
    Dim DragRect As RECT
    Dim na As Long
    Dim lBorderWidth As Long
    Dim lObjWidth As Long
    Dim lObjHeight As Long
    Dim lXOffset As Long
    Dim lYOffset As Long
    Dim bMoved As Boolean
    
    ReleaseCapture
    GetWindowRect hwnd, objRect
    lObjWidth = objRect.Right - objRect.Left
    lObjHeight = objRect.Bottom - objRect.Top
    GetCursorPos pt

    ptPrev.X = pt.X
    ptPrev.Y = pt.Y

    lXOffset = pt.X - objRect.Left
    lYOffset = pt.Y - objRect.Top
    
    With DragRect
        .Left = pt.X - lXOffset
        .Top = pt.Y - lYOffset
        .Right = .Left + lObjWidth
        .Bottom = .Top + lObjHeight
    End With

    lBorderWidth = 3
    DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth

    Do While GetKeyState(VK_LBUTTON) < 0
        GetCursorPos pt
        If pt.X <> ptPrev.X Or pt.Y <> ptPrev.Y Then
            ptPrev.X = pt.X
            ptPrev.Y = pt.Y

            DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth

            FormMoved pt.X - lXOffset, pt.Y - lYOffset, lObjWidth, lObjHeight

            With DragRect
                .Left = pt.X - lXOffset
                .Top = pt.Y - lYOffset
                .Right = .Left + lObjWidth
                .Bottom = .Top + lObjHeight
            End With
            DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
            bMoved = True
        End If

    Loop

    DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth

    If bMoved Then
        If m_RepositionForm Then
            MoveWindow hwnd, DragRect.Left, DragRect.Top, DragRect.Right - DragRect.Left, DragRect.Bottom - DragRect.Top, True
        End If

        FormDropped DragRect.Left, DragRect.Top, DragRect.Right - DragRect.Left, DragRect.Bottom - DragRect.Top
    End If
    
End Sub

Private Sub DrawDragRectangle(ByVal X As Long, ByVal Y As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal lWidth As Long)

    Dim hDC As Long
    Dim hPen As Long
    hPen = CreatePen(PS_SOLID, lWidth, &HE0E0E0)
    hDC = GetDC(0)
    Call SelectObject(hDC, hPen)
    Call SetROP2(hDC, R2_NOTXORPEN)
    Call Rectangle(hDC, X, Y, x1, y1)
    Call SelectObject(hDC, GetStockObject(BLACK_PEN))
    Call DeleteObject(hPen)
    Call SelectObject(hDC, hPen)
    Call ReleaseDC(0, hDC)
    
End Sub

Private Sub UserControl_Paint()
    
    Dim lBackColor As Long
    Dim sCaption As String
       
    If Not mdifrm Is Nothing Then
        If mdifrm.Visible And mdifrm.WindowState = 0 Then mdifrm.Move mdifrm.Left, mdifrm.Top, mdifrm.Width, mdifrm.Height
    End If
        
    With UserControl
        .Cls
        .Extender.Align = vbAlignTop
        .Extender.Top = 0
        .Height = GetSystemMetrics(SM_CYCAPTION) * Screen.TwipsPerPixelY
        If bDocked Then
            .ForeColor = vbTitleBarText
            lBackColor = vbActiveTitleBar
        Else
            If GetActiveWindow = UserControl.Extender.Parent.hwnd Then
                .ForeColor = vbTitleBarText
                lBackColor = vbActiveTitleBar
            Else
                .ForeColor = vbInactiveTitleBarText
                lBackColor = vbInactiveTitleBar
            End If
        End If
        UserControl.Line (Screen.TwipsPerPixelX, Screen.TwipsPerPixelY)-(UserControl.ScaleWidth - (2 * Screen.TwipsPerPixelX), UserControl.ScaleHeight - Screen.TwipsPerPixelY), lBackColor, BF
        .CurrentX = 4 * Screen.TwipsPerPixelX
        .CurrentY = 3 * Screen.TwipsPerPixelY
        .Font.name = "MS Sans Serif"
        .Font.Bold = True
        
        sCaption = m_Caption
        If UserControl.TextWidth(sCaption) > (UserControl.ScaleWidth - (4 * Screen.TwipsPerPixelX)) Then
             Do While UserControl.TextWidth(sCaption & "...") > (UserControl.ScaleWidth - (4 * Screen.TwipsPerPixelX)) And Len(sCaption) > 0
                sCaption = Trim$(Left$(sCaption, Len(sCaption) - 1))
            Loop
            sCaption = sCaption & "..."
        End If
        UserControl.Print sCaption;
        
        UpdateWindow UserControl.hwnd
    End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_RepositionForm = PropBag.ReadProperty("RepositionForm", m_def_RepositionForm)
    bDockable = PropBag.ReadProperty("Dockable", bDockable)
    bResizable = PropBag.ReadProperty("Resizable", bResizable)
    bMovable = PropBag.ReadProperty("Movable", bMovable)
    bDocked = PropBag.ReadProperty("Docked", bDocked)
    UserControl_Paint
End Sub

Private Sub UserControl_Resize()
    UserControl_Paint
End Sub

Private Sub UserControl_Show()
    UserControl.Refresh
    UserControl_Paint
End Sub

Private Sub UserControl_Terminate()
    bDocked = False
    If Not frm Is Nothing Then
        SetParent frm.hwnd, 0
        SetVisible False
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    Call PropBag.WriteProperty("Dockable", bDockable, True)
        
    Call PropBag.WriteProperty("Resizable", bResizable, True)
    Call PropBag.WriteProperty("Movable", bMovable, True)
    Call PropBag.WriteProperty("Docked", bDocked, True)
    
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("RepositionForm", m_RepositionForm, m_def_RepositionForm)

End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    UserControl_Paint
End Property

Public Property Get RepositionForm() As Boolean
    RepositionForm = m_RepositionForm
End Property

Public Property Let RepositionForm(ByVal New_RepositionForm As Boolean)
    m_RepositionForm = New_RepositionForm
    PropertyChanged "RepositionForm"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Public Sub ToggleDocked(ByVal IsDocked As Boolean, Optional ByVal Override As Boolean = False)
    bMoving = True And (Not Override)
    bDocked = IsDocked
    If IsDocked Then
        frm.Visible = False
        bDocked = False
        SetParent frm.hwnd, 0
        frm.Move lFloatingLeft, lFloatingTop, lFloatingWidth, lFloatingHeight
        SetVisible False
        frm.Visible = True
        Call SetWindowWord(frm.hwnd, SWW_HPARENT, mdifrm.hwnd)
    Else
        bDocked = True
        SetParent frm.hwnd, picbox.hwnd
        frm.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, lDockedWidth, lDockedHeight
        SetVisible True
    End If
    bMoving = False
End Sub
Private Sub UserControl_DblClick()
    If bDockable Then ToggleDocked bDocked
End Sub

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub





