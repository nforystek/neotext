VERSION 5.00
Begin VB.UserControl Dragger 
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
   ToolboxBitmap   =   "Dragger.ctx":0000
End
Attribute VB_Name = "Dragger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary
Public GradClr1 As OLE_COLOR
Public GradClr2 As OLE_COLOR
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * 32
End Type

Private Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
End Type


Private Enum tdBorderStyles
    bdrNone = 0
    bdrRaisedOuter = 1
    bdrRaisedInner = 2
    bdrRaised = 3
    bdrSunkenOuter = 4
    bdrSunkenInner = 5
    bdrSunken = 6
    bdrEtched = 7
    bdrBump = 8
    bdrMono = 9
    bdrFlat = 10
    bdrSoft = 11
End Enum


Public Enum tdCaptionStyles
    tdCaptionNormal = 0
    tdCaptionEtched = 1
    tdCaptionSoft = 2
    tdCaptionRaised = 3
    tdCaptionRaisedInner = 4
    tdCaptionSunkenOuter = 5
    tdCaptionSunken = 6
    tdCaptionSingleRaisedBar = 7
    tdCaptionGradient = 8
    tdCaptionSingleRaisedInner = 9
    tdCaptionSingleSoft = 10
    tdCaptionSingleEtched = 11
    tdCaptionSingleSunken = 12
    tdCaptionSingleSunkenOuter = 13
    tdCaptionOfficeXP = 14
End Enum

Private Const DT_END_ELLIPSIS = &H8000&
Dim captionFont As LOGFONT

Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4

Private Const DT_CENTER = &H1
Private Const DT_BOTTOM = &H8

' System metrics constants
Private Const SM_CXMIN = 28
Private Const SM_CYMIN = 29
Private Const SM_CXSIZE = 30
Private Const SM_CXFRAME = 32
Private Const SM_CYFRAME = 33
Private Const SM_CYSIZE = 31
Private Const SM_CYCAPTION = 4
Private Const SM_CXBORDER = 5
Private Const SM_CYBORDER = 6
Private Const SM_CYMENU = 15
Private Const SM_CYSMCAPTION = 51 'height of windows 95 small caption

' These constants define the style of border to draw.
Private Const BDR_RAISED = &H5
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKEN = &HA
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2

Private Const BF_FLAT = &H4000
Private Const BF_MONO = &H8000
Private Const BF_SOFT = &H1000      ' For softer buttons

Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

' These constants define which sides to draw.
Private Const BF_BOTTOM = &H8
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const TRANSPARENT = 1
Private Const OPAQUE = 2

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_TOOLWINDOW = &H80
Private Const VK_LBUTTON = &H1
Private Const PS_SOLID = 0
Private Const R2_NOTXORPEN = 10
Private Const BLACK_PEN = 7


Private Const SPI_GETNONCLIENTMETRICS = 41
' align properties for each panel that is created
' for the docking engine
Public Enum tdAlignProperty
    tdAlignNone = 0     ' Floating host not implemented
    tdAlignTop = 1      ' Top host
    tdAlignBottom = 2   ' Bottom host
    tdAlignLeft = 3     ' Left Host
    tdAlignRight = 4    ' Right Host
End Enum
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
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
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal clr As Long, ByVal hpal As Long, ByRef lpcolorref As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Const CLR_INVALID = 0

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
Private WithEvents frmPar As Form
Attribute frmPar.VB_VarHelpID = -1

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
Event Resize(Left As Long, Top As Long, width As Long, ByVal Height As Long)

Private Const m_def_RepositionForm = True
Private Const m_def_Caption = "Dragger1"

Private m_RepositionForm As Boolean
Private m_Caption As String


Private Const DFC_CAPTION = 1
Private Const DFC_MENU = 2               'Menu
Private Const DFC_SCROLL = 3             'Scroll bar
Private Const DFC_BUTTON = 4             'Standard button



Private Const DFCS_CAPTIONCLOSE = &H0
Private Const DFCS_CAPTIONRESTORE = &H3
Private Const DFCS_FLAT = &H4000
Private Const DFCS_PUSHED = &H200
Private Const DFCS_MENUARROWRIGHT = &H4
Private Const DFCS_SCROLLUP = &H0
Private Const DFCS_SCROLLLEFT = &H2
'Private Const DFCS_FLAT = &H4000

Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Const SPLITTER_HEIGHT = 80
Private Const SPLITTER_WIDTH = 80
Private m_Align As tdAlignProperty

Public Property Get Align() As tdAlignProperty
    Align = m_Align
End Property

Public Property Let Align(New_Align As tdAlignProperty)
    Extender.Align = New_Align
    m_Align = New_Align
End Property
Public Sub gradateColors(Colors() As Long, ByVal color1 As Long, ByVal Color2 As Long)

'Alright, I admit -- this routine was
'taken from a VBPJ issue a few months back.

Dim i As Integer
Dim dblR As Double, dblG As Double, dblB As Double
Dim addR As Double, addG As Double, addB As Double
Dim bckR As Double, bckG As Double, bckB As Double

   dblR = CDbl(color1 And &HFF)
   dblG = CDbl(color1 And &HFF00&) / 255
   dblB = CDbl(color1 And &HFF0000) / &HFF00&
   bckR = CDbl(Color2 And &HFF&)
   bckG = CDbl(Color2 And &HFF00&) / 255
   bckB = CDbl(Color2 And &HFF0000) / &HFF00&
   
   addR = (bckR - dblR) / UBound(Colors)
   addG = (bckG - dblG) / UBound(Colors)
   addB = (bckB - dblB) / UBound(Colors)
   
   For i = 0 To UBound(Colors)
      dblR = dblR + addR
      dblG = dblG + addG
      dblB = dblB + addB
      If dblR > 255 Then dblR = 255
      If dblG > 255 Then dblG = 255
      If dblB > 255 Then dblB = 255
      If dblR < 0 Then dblR = 0
      If dblG < 0 Then dblG = 0
      If dblG < 0 Then dblB = 0
      Colors(i) = RGB(dblR, dblG, dblB)
   Next
End Sub

Private Sub drawGradient(captionRect As RECT, hDC As Long, captionText As String, bActive As Boolean, gradient As Boolean, Optional captionOrientation As Integer, Optional captionForm As Form)

    Dim hBr As Long
    Dim drawDC As Long
    Dim bar As Long
    Dim width As Long
    Dim pixelStep As Long
    Dim storedCaptionRect As RECT
    Dim tmpGradFont As Long
    Dim oldFont As Long
    Dim hDCTemp As Long
    
    hDCTemp = hDC
    
    'Debug.Print captionText, captionOrientation, hDC, hDCTemp
    
    storedCaptionRect = captionRect
    
    If captionOrientation <> tdAlignTop And captionOrientation <> tdAlignBottom Then
        width = captionRect.Right - captionRect.Left
    Else
        width = captionRect.Bottom - captionRect.Top
    End If
    
    pixelStep = width / 4
    
    ReDim Colors(pixelStep) As Long
    
    ' determine colors of gradient fill also determine if a gradient fill is required
    If bActive Then
        If gradient Then
            gradateColors Colors(), GradClr1, GradClr2
        Else
            gradateColors Colors(), TranslateColor(vbActiveTitleBar), TranslateColor(vbActiveTitleBar)
        End If
    Else
        If gradient Then
            gradateColors Colors(), TranslateColor(vbInactiveTitleBar), TranslateColor(vbButtonFace)
        Else
            gradateColors Colors(), TranslateColor(vbInactiveTitleBar), TranslateColor(vbInactiveTitleBar)
        End If
    End If
    
    For bar = 1 To pixelStep - 1
        hBr = CreateSolidBrush(Colors(bar))
        
        FillRect hDCTemp, captionRect, hBr
        
        If captionOrientation <> tdAlignTop And captionOrientation <> tdAlignBottom Then
            captionRect.Left = captionRect.Left + 4
        Else
            captionRect.Bottom = captionRect.Bottom - 4
        End If
        
        DeleteObject hBr
    Next bar
  
    'draw caption text
    'Use a white caption, since the background is black
    'on the left side
    
    'get caption font information
    getCapsFont
    
    'If getting the caption font failed, use the font
    'from the gradient caption form.
    tmpGradFont = 0
    
    If captionText = "Form6" Then
      '  Beep
    End If
    
    If tmpGradFont = 0 Then
    
        'tmpGradFont = CreateFontIndirect(captionFont)
        
        If captionOrientation = tdAlignTop Or captionOrientation = tdAlignBottom Then
            captionFont.lfEscapement = 900
            
            'hDCTemp = captionForm.hDC
            'debug.Print "gradient font hdc set"
        End If
        
        tmpGradFont = CreateFontIndirect(captionFont)
        oldFont = SelectObject(hDCTemp, tmpGradFont)
    End If
    
    SetBkMode hDCTemp, TRANSPARENT
    
    If (bActive) Then
       SetTextColor hDCTemp, TranslateColor(vbActiveTitleBarText)
    Else
       SetTextColor hDCTemp, TranslateColor(vbInactiveTitleBarText)
    End If
    
    'move text a wee bit to the right
    If captionOrientation = tdAlignTop Or captionOrientation = tdAlignBottom Then
        'captionForm.CurrentX = 50
        'captionForm.CurrentY = captionForm.ScaleHeight - 100
        'captionForm.Print captionText
        'Debug.Print "caption text drawn", captionForm.CurrentX
        storedCaptionRect.Right = storedCaptionRect.Bottom - 40
        storedCaptionRect.Bottom = 8 + (captionForm.Height / Screen.TwipsPerPixelY)
        'Debug.Print "pixel height = "; captionForm.height / Screen.TwipsPerPixelY
        
        DrawText hDCTemp, captionText, Len(captionText), storedCaptionRect, DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_BOTTOM
    Else
        storedCaptionRect.Left = storedCaptionRect.Left + 2
        storedCaptionRect.Right = storedCaptionRect.Right - 40
        DrawText hDCTemp, captionText, Len(captionText), storedCaptionRect, DT_SINGLELINE Or DT_END_ELLIPSIS 'Or DT_HCENTER
    End If
    
    SelectObject hDCTemp, oldFont
    DeleteObject tmpGradFont
    tmpGradFont = 0

End Sub
Private Sub drawOfficeXP(captionRect As RECT, hDC As Long, captionText As String, bActive As Boolean, gradient As Boolean, Optional captionOrientation As Integer, Optional captionForm As Form)

    Dim hBr As Long
    Dim drawDC As Long
    Dim bar As Long
    Dim width As Long
    Dim pixelStep As Long
    Dim storedCaptionRect As RECT
    Dim tmpGradFont As Long
    Dim oldFont As Long
    Dim hDCTemp As Long
    Dim colorOutline As Long
    Dim colorInline As Long
    
    hDCTemp = hDC
    
    'Debug.Print captionText, captionOrientation, hDC, hDCTemp
    
    storedCaptionRect = captionRect
    
    If captionOrientation <> tdAlignTop And captionOrientation <> tdAlignBottom Then
        width = captionRect.Right - captionRect.Left
    Else
        width = captionRect.Bottom - captionRect.Top
    End If
        
    ' determine colors of gradient fill also determine if a gradient fill is required
    If bActive Then
        colorOutline = TranslateColor(vbActiveTitleBar)
        colorInline = TranslateColor(vbActiveTitleBar)
    Else
        colorOutline = TranslateColor(vbInactiveTitleBar)
        colorInline = TranslateColor(vbButtonFace)
    End If
    

    hBr = CreateSolidBrush(colorOutline)
        
    FillRect hDCTemp, captionRect, hBr
    
    With captionRect
        .Top = .Top + 1
        .Left = .Left + 1
        .Right = .Right - 1
        .Bottom = .Bottom - 1
    End With
        
    hBr = CreateSolidBrush(colorInline)
        
    FillRect hDCTemp, captionRect, hBr
    
    DeleteObject hBr
  
    'draw caption text
    'Use a white caption, since the background is black
    'on the left side
    
    'get caption font information
    getCapsFont
    
    'If getting the caption font failed, use the font
    'from the gradient caption form.
    tmpGradFont = 0
    
    If captionText = "Form6" Then
      '  Beep
    End If
    
    If tmpGradFont = 0 Then
    
        'tmpGradFont = CreateFontIndirect(captionFont)
        
        If captionOrientation = tdAlignTop Or captionOrientation = tdAlignBottom Then
            captionFont.lfEscapement = 900
            
            'hDCTemp = captionForm.hDC
            'debug.Print "gradient font hdc set"
        End If
        
        tmpGradFont = CreateFontIndirect(captionFont)
        oldFont = SelectObject(hDCTemp, tmpGradFont)
    End If
    
    SetBkMode hDCTemp, TRANSPARENT
    
    If (bActive) Then
       SetTextColor hDCTemp, TranslateColor(vbActiveTitleBarText)
    Else
       SetTextColor hDCTemp, TranslateColor(vbInactiveTitleBarText)
    End If
    
    'move text a wee bit to the right
    If captionOrientation = tdAlignTop Or captionOrientation = tdAlignBottom Then
        'captionForm.CurrentX = 50
        'captionForm.CurrentY = captionForm.ScaleHeight - 100
        'captionForm.Print captionText
        'Debug.Print "caption text drawn", captionForm.CurrentX
        storedCaptionRect.Right = storedCaptionRect.Bottom - 40
        storedCaptionRect.Bottom = 8 + (captionForm.Height / Screen.TwipsPerPixelY)
        'Debug.Print "pixel height = "; captionForm.height / Screen.TwipsPerPixelY
        
        DrawText hDCTemp, captionText, Len(captionText), storedCaptionRect, DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_BOTTOM
    Else
        storedCaptionRect.Left = storedCaptionRect.Left + 2
        storedCaptionRect.Right = storedCaptionRect.Right - 40
        DrawText hDCTemp, captionText, Len(captionText), storedCaptionRect, DT_SINGLELINE Or DT_END_ELLIPSIS 'Or DT_HCENTER
    End If
    
    SelectObject hDCTemp, oldFont
    DeleteObject tmpGradFont
    tmpGradFont = 0

End Sub

Private Sub drawGripper(captionRect As RECT, hDC As Long, gripStyle As Long, gripSides As Long, oneBar As Boolean, captionHeight As Long, Optional captionOrientation As Integer, Optional maximiseButton As Boolean)
    
    Dim numOfButtons As Integer
    
    If maximiseButton Then
        numOfButtons = 2
    Else
        numOfButtons = 1
    End If
    
    If oneBar Then
    
        If captionOrientation <> tdAlignTop And captionOrientation <> tdAlignBottom Then
            With captionRect
                .Top = .Top + ((captionHeight - 11) / 2)
                .Left = .Left + 1
                .Right = .Right - (captionHeight * numOfButtons) + 5
                .Bottom = .Top + 4
            End With
        Else
            With captionRect
                .Top = .Top + (captionHeight * numOfButtons) - 4
                .Left = .Left + ((captionHeight - 14) / 2)
                .Right = .Left + 4
                .Bottom = .Bottom - 2
            End With
        End If
        
        DrawEdge hDC, captionRect, gripStyle, gripSides
    
    Else
    
        If captionOrientation <> tdAlignTop And captionOrientation <> tdAlignBottom Then
            With captionRect
                .Top = .Top + ((captionHeight - 16) / 2)
                .Left = .Left + 1
                .Right = .Right - (captionHeight * numOfButtons) + 5
                .Bottom = .Top + 4
            End With
        Else
            With captionRect
                .Top = .Top + (captionHeight * numOfButtons) - 4
                .Left = .Left + ((captionHeight - 20) / 2) + 1
                .Right = .Left + 4
                .Bottom = .Bottom - 2
            End With
        End If
        
        DrawEdge hDC, captionRect, gripStyle, gripSides
        
        If captionOrientation <> tdAlignTop And captionOrientation <> tdAlignBottom Then
            With captionRect
                .Top = .Bottom + 1
                .Bottom = .Bottom + 5
            End With
        Else
            With captionRect
                .Left = .Right + 1
                .Right = .Left + 4
            End With
        End If
        
        DrawEdge hDC, captionRect, gripStyle, gripSides
        
    End If

End Sub

Private Sub getCapsFont()

    Dim NCM As NONCLIENTMETRICS
    Dim lfNew As LOGFONT

    NCM.cbSize = Len(NCM)
    Call SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, NCM, 0)
    
    If NCM.iCaptionHeight = 0 Then
       captionFont.lfHeight = 0
    Else
       captionFont = NCM.lfSMCaptionFont
       'If captionFont.lfHeight < 10 Then
       ' captionFont.lfHeight = 14
       'End If
    End If
    
End Sub

Private Function getCaptionButtonHeight() As Long
    
    Dim NCM As NONCLIENTMETRICS
    Dim lfNew As LOGFONT

    NCM.cbSize = Len(NCM)
    Call SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, NCM, 0)
    
    If NCM.iCaptionHeight = 0 Then
       'captionFont.lfHeight = 0
       getCaptionButtonHeight = 14
    Else
       'captionFont = NCM.lfSMCaptionFont
       getCaptionButtonHeight = NCM.iSMCaptionHeight
    End If
    
End Function
Private Function getCaptionHeight() As Long
        
    getCaptionHeight = GetSystemMetrics(SM_CYSMCAPTION)
    
    'If getCaptionHeight < 20 Then getCaptionHeight = 15
    
End Function
Private Property Get frm() As Object
    On Error Resume Next
    
    If Not UserControl Is Nothing Then
        Set frmPar = UserControl.Parent
        Set frm = UserControl.Parent
    End If
    
    If Err Then Err.Clear
    
    
End Property

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

Public Property Get MDIParent() As Object
    Set MDIParent = mdifrm
End Property
Public Property Set MDIParent(ByRef RHS As Object)
    Set mdifrm = RHS
    SetParentUp
End Property

Public Sub SetVisible(ByVal IsVisible As Boolean)
  '  picbar.Visible = False
  '  picbox.Visible = False
    Select Case iAlign
        Case 1, 3
          '  picbar.Visible = IsVisible And bDocked
    '        picbox.Visible = IsVisible And bDocked
        Case 2, 4
         '   picbox.Visible = IsVisible And bDocked
    '        picbar.Visible = IsVisible And bDocked
    End Select
End Sub
Private Sub frmPar_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SetWindowWord(frm.hwnd, SWW_HPARENT, 0&)
End Sub

Private Sub frmPar_Resize()
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

        If Not mdifrm Is Nothing Then
            Call SetWindowWord(frm.hwnd, SWW_HPARENT, mdifrm.hwnd)
        End If
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
    If picbox.width < 120 Then
        picbox.width = 120
    End If
    If bDocked Then
        frm.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, picbox.ScaleWidth + (8 * Screen.TwipsPerPixelX), picbox.ScaleHeight + (8 * Screen.TwipsPerPixelY)
    End If
    If Err Then Err.Clear
End Sub

Private Sub StoreFormDimensions()

    If Not bMoving Then
        If bDocked Then
            lDockedWidth = frm.width
            lDockedHeight = frm.Height
        Else
            lFloatingLeft = frm.Left
            lFloatingTop = frm.Top
            lFloatingWidth = frm.width
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

Private Sub UserControl_InitProperties()
    SetParentUp
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

' ******************************************************************************
' Routine       : DrawBorder
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 02/10/0010:20:55
' Inputs        :
' Outputs       :
' Credits       : linda.69@mailcity.com (Color your Border demo)
' Modifications : Color translation OLE_COLOR to RGB
' Description   : draw a user defined color border
' ******************************************************************************
Private Sub DrawBorder(frmTarget As Form, Color As OLE_COLOR)
    Dim hWindowDC As Long
    Dim hOldPen As Long
    Dim nLeft As Long
    Dim nRight As Long
    Dim nTop As Long
    Dim nBottom As Long
    Dim Ret As Long
    Dim hMyPen As Long
    Dim WidthX As Long
    Dim rgbColor As Long
    
    ' translate
    rgbColor = TranslateColor(Color)
    ' border width
    WidthX = GetSystemMetrics(SM_CYBORDER) * 5
    ' get window DC
    hWindowDC = GetWindowDC(frmTarget.hwnd)   'this is outside the form
    ' create a pen
    hMyPen = CreatePen(PS_SOLID, WidthX, rgbColor)
    ' Initialize misc variables
    nLeft = 0: nTop = 0
    nRight = frmTarget.width / Screen.TwipsPerPixelX
    nBottom = frmTarget.Height / Screen.TwipsPerPixelY
    ' select border pen
    hOldPen = SelectObject(hWindowDC, hMyPen)
    ' draw color around the border
    Ret = LineTo(hWindowDC, nLeft, nBottom)
    Ret = LineTo(hWindowDC, nRight, nBottom)
    Ret = LineTo(hWindowDC, nRight, nTop)
    Ret = LineTo(hWindowDC, nLeft, nTop)
    ' select old pen
    Ret = SelectObject(hWindowDC, hOldPen)
    Ret = DeleteObject(hMyPen)
    Ret = ReleaseDC(frmTarget.hwnd, hWindowDC)
End Sub

' ******************************************************************************
' Routine       : TranslateColor
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 02/10/0010:20:19
' Inputs        :
' Outputs       :
' Credits       : Extracted from VB KB Article
' Modifications :
' Description   : Converts an OLE_COLOR to RGB color
' ******************************************************************************
Private Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hpal As Long = 0) As Long
    If OleTranslateColor(clr, hpal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function
Private Sub UserControl_Paint()
'    Dim Rc As RECT
'    Dim bdrStyle As Long
'    Dim bdrSides As Long
'    Dim BorderStyle  As tdBorderStyles
'    Dim CaptionStyle As tdCaptionStyles
'    Dim hDC As Long
'    Dim captionHeight As Long
'
'    If UserControl.Extender.Name = "Form6" Then
'        'debug.Print Align, Me.Extender.Name, Me.State
'    End If
'
'
'    ' draw a custom border based on parante's color
'    DrawBorder Parent, Parent.BackColor
'    ' retrieve TabDock border style
'    BorderStyle = Parent.BorderStyle
'    'CaptionStyle = Parent.CaptionStyle
'    ' all sides must be updated
'    bdrSides = BF_RECT
'    ' update border styles
'    If BorderStyle = bdrFlat Then bdrSides = bdrSides Or BF_FLAT
'    If BorderStyle = bdrMono Then bdrSides = bdrSides Or BF_MONO
'    If BorderStyle = bdrSoft Then bdrSides = bdrSides Or BF_SOFT
'    Select Case BorderStyle
'        Case bdrRaisedOuter: bdrStyle = BDR_RAISEDOUTER
'        Case bdrRaisedInner: bdrStyle = BDR_RAISEDINNER
'        Case bdrRaised: bdrStyle = EDGE_RAISED
'        Case bdrSunkenOuter: bdrStyle = BDR_SUNKENOUTER
'        Case bdrSunkenInner: bdrStyle = BDR_SUNKENINNER
'        Case bdrSunken: bdrStyle = EDGE_SUNKEN
'        Case bdrEtched: bdrStyle = EDGE_ETCHED
'        Case bdrBump: bdrStyle = EDGE_BUMP
'        Case bdrFlat: bdrStyle = BDR_SUNKEN
'        Case bdrMono: bdrStyle = BDR_SUNKEN
'        Case bdrSoft: bdrStyle = BDR_RAISED
'    End Select
'
'    ' get a window rect by hand
'    ' GetWindowRect will not work here!
'    Rc.Left = 0
'    Rc.Top = 0
'    Rc.Bottom = Extender.Height / Screen.TwipsPerPixelY
'    Rc.Right = Extender.width / Screen.TwipsPerPixelY
'    ' First get the window DC
'    hDC = GetWindowDC(hwnd)
'    ' Simply call the API and draw the edge.
'    DrawEdge hDC, Rc, bdrStyle, bdrSides
'
'    ' get a window rect by hand
'    ' GetWindowRect will not work here!
'    Rc.Left = 0
'    Rc.Top = 0
'    Rc.Bottom = Height / Screen.TwipsPerPixelY + 1
'    Rc.Right = width / Screen.TwipsPerPixelX
'
'    If Align = tdAlignLeft Then
'        Rc.Right = (width - SPLITTER_WIDTH) / Screen.TwipsPerPixelX + 1
'        Rc.Bottom = Height / Screen.TwipsPerPixelY
'    End If
'
'    If Align = tdAlignRight Then
'        Rc.Left = SPLITTER_WIDTH / Screen.TwipsPerPixelX
'        Rc.Bottom = Height / Screen.TwipsPerPixelY
'    End If
'
'    If Align = tdAlignTop Then
'        Rc.Bottom = (Height - SPLITTER_HEIGHT) / Screen.TwipsPerPixelY + 1
'    End If
'
'    If Align = tdAlignBottom Then
'        Rc.Top = SPLITTER_HEIGHT / Screen.TwipsPerPixelY
'        Rc.Bottom = Height / Screen.TwipsPerPixelY
'    End If
'
'    ' First get the window DC
'    hDC = GetWindowDC(UserControl.hwnd)
'    ' Simply call the API and draw the edge.
'    DrawEdge hDC, Rc, bdrStyle, bdrSides
'
'    Rc.Left = 0
'    Rc.Top = 0
'    Rc.Bottom = Extender.Height / Screen.TwipsPerPixelY
'    Rc.Right = Extender.width / Screen.TwipsPerPixelY
'
'    hDC = GetWindowDC(hwnd)
'    '************************************************
'
'    If UserControl.BorderStyle = vbBSNone Or Align = tdAlignTop Or Align = tdAlignBottom Then
'        'draw custom caption here!!!!
'
'        Dim captionRect As RECT
'        Dim frameRect As RECT
'        Dim hBr As Long
'
'        captionHeight = getCaptionHeight + 6
'
'        hBr = CreateSolidBrush(TranslateColor(vbButtonFace))
'
'        With captionRect
'            .Top = Rc.Top + 3
'            .Left = Rc.Left + 3
'            .Right = Rc.Right - 2
'            .Bottom = captionHeight - 4
'
'           If Align = tdAlignTop Or Align = tdAlignBottom Then
'                .Top = .Top
'                .Bottom = Rc.Bottom - 3
'                .Right = captionHeight - 4
'            End If
'
'        End With
'
'
'        'blank out current caption
'
'        FillRect hDC, captionRect, hBr
'        DeleteObject hBr
'
''         Select Case CaptionStyle
''            Case tdCaptionNormal ' = 0
'                drawGradient captionRect, hDC, UserControl.Extender.Caption, GetActiveWindow() = UserControl.hwnd, False, Align, Parent
''            Case tdCaptionEtched ' = 1
''                drawGripper captionRect, hDC, EDGE_ETCHED, BF_RECT, False, captionHeight, Align, Parent.MaximizeButton
''            Case tdCaptionSoft ' = 2
''                drawGripper captionRect, hDC, BDR_RAISED, BF_RECT Or BF_SOFT, False, captionHeight, Align, Parent.MaximizeButton
''            Case tdCaptionRaised ' = 3
''                drawGripper captionRect, hDC, EDGE_RAISED, BF_RECT, False, captionHeight, Align, Parent.MaximizeButton
''            Case tdCaptionRaisedInner ' = 4
''                drawGripper captionRect, hDC, BDR_RAISEDINNER, BF_RECT, False, captionHeight, Align, Parent.MaximizeButton
''            Case tdCaptionSunkenOuter ' = 5
''                drawGripper captionRect, hDC, BDR_SUNKENOUTER, BF_RECT, False, captionHeight, Align, Parent.MaximizeButton
''            Case tdCaptionSunken ' = 6
''                drawGripper captionRect, hDC, EDGE_SUNKEN, BF_RECT, False, captionHeight, Align, Parent.MaximizeButton
''            Case tdCaptionSingleRaisedBar ' = 7
''                drawGripper captionRect, hDC, EDGE_RAISED, BF_RECT, True, captionHeight, Align, Parent.MaximizeButton
''            Case tdCaptionSingleRaisedInner ' = 9
''                drawGripper captionRect, hDC, BDR_RAISEDINNER, BF_RECT, True, captionHeight, Align, Parent.MaximizeButton
''            Case tdCaptionSingleSoft ' = 10
''                drawGripper captionRect, hDC, BDR_RAISED, BF_RECT Or BF_SOFT, True, captionHeight, Align, Parent.MaximizeButton
''            Case tdCaptionSingleEtched ' = 11
''                drawGripper captionRect, hDC, EDGE_ETCHED, BF_RECT, True, captionHeight, Align, Parent.MaximizeButton
''            Case tdCaptionSingleSunken ' = 12
''                drawGripper captionRect, hDC, EDGE_SUNKEN, BF_RECT, True, captionHeight, Align, Parent.MaximizeButton
''            Case tdCaptionSingleSunkenOuter ' = 13
''                drawGripper captionRect, hDC, BDR_SUNKENOUTER, BF_RECT, True, captionHeight, Align, Parent.MaximizeButton
''            Case tdCaptionGradient ' = 8
''                drawGradient captionRect, hDC, UserControl.Extender.Caption, GetActiveWindow() = UserControl.hwnd, True, Align, UserControl.Extender
''            Case tdCaptionOfficeXP ' = 14
''                drawOfficeXP captionRect, hDC, UserControl.Extender.Caption, GetActiveWindow() = UserControl.hwnd, True, Align, UserControl.Extender
''
''        End Select
'
'        'draw close box on form
'
'        With frameRect
'            .Top = Rc.Top + 5
'            .Left = Rc.Right - captionHeight + 6
'            .Right = Rc.Right - 4
'            .Bottom = captionHeight - 5
'
'            If Align = tdAlignTop Or Align = tdAlignBottom Then
'                .Right = Rc.Left + captionHeight - 6
'                .Left = Rc.Left + 4
'            End If
'
'        End With
'
'        If UserControl.Extender.Name = "Form6" Then
'            'debug.Print frameRect.Top, frameRect.Left, frameRect.Right, frameRect.Bottom
'        End If
'
'        DrawFrameControl hDC, frameRect, DFC_CAPTION, DFCS_CAPTIONCLOSE 'Or DFCS_FLAT
'
'        If Parent.MaxButton Then
'            If Align = tdAlignTop Or Align = tdAlignBottom Then
'                With frameRect
'                    .Top = .Bottom + 2
'                    .Bottom = .Top + captionHeight - 10
'                    '.Right = .Right + 1
'                    '.Left = .Left - 1
'                End With
'                DrawFrameControl hDC, frameRect, DFC_CAPTION, DFCS_CAPTIONRESTORE 'Or DFCS_FLAT
'
'            Else
'                With frameRect
'                    .Right = .Left - 2
'                    .Left = .Right - captionHeight + 10
'                End With
'                DrawFrameControl hDC, frameRect, DFC_CAPTION, DFCS_CAPTIONRESTORE 'Or DFCS_FLAT
'
'            End If
'
'        End If
'
'    End If
'    '*********************************************************
'    ' release it
'
'    ReleaseDC hwnd, hDC
    
    
    
    
    
    Dim lBackColor As Long
    Dim sCaption As String

    If Not mdifrm Is Nothing Then
        If mdifrm.Visible And mdifrm.WindowState = 0 Then mdifrm.Move mdifrm.Left, mdifrm.Top, mdifrm.width, mdifrm.Height
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
        sCaption = m_Caption
        .CurrentX = (UserControl.ScaleHeight / 2) - (UserControl.TextHeight(sCaption) / 2)
        .CurrentY = .CurrentX

        If Not Parent Is Nothing Then
            Set .Font = Parent.Font
            .FontBold = Parent.Font.Bold
            .FontItalic = Parent.Font.Italic
            .FontName = Parent.Font.Name
            .FontSize = Parent.Font.Size
            .FontStrikethru = Parent.Font.Strikethrough
            .FontUnderline = Parent.Font.Underline

        End If


        If UserControl.TextWidth(sCaption) > (UserControl.ScaleWidth - (4 * Screen.TwipsPerPixelX)) Then
             Do While UserControl.TextWidth(sCaption & "...") > (UserControl.ScaleWidth - (4 * Screen.TwipsPerPixelX)) And Len(sCaption) > 0
                sCaption = Trim$(Left$(sCaption, Len(sCaption) - 1))
            Loop
            sCaption = sCaption & "..."
        End If
        UserControl.Print sCaption;

        Dim rec As RECT
        rec.Top = (Screen.TwipsPerPixelY * 4)
        rec.Bottom = UserControl.ScaleHeight - (Screen.TwipsPerPixelY * 6)
        rec.Left = UserControl.ScaleWidth - (Screen.TwipsPerPixelX * 24)

        rec.Right = UserControl.ScaleWidth - (Screen.TwipsPerPixelX * 7)

       ' UserControl.Line (rec.Left, rec.Top)-(rec.Right, rec.Bottom), vbBlack, BF

        UpdateWindow UserControl.hwnd


    End With
End Sub

Private Sub SetParentUp()
    If Not frm Is Nothing Then
        If frm.Caption <> "" Then
            Me.Caption = frm.Caption
            m_Caption = frm.Caption
            frm.Caption = ""

        End If
        
        'frm.BorderStyle = 0
       ' frm.MaxButton = False
       ' frm.MinButton = False
       ' frm.ControlBox = False
      
        Set UserControl.Font = frm.Font

        DragObject Me.hwnd
        
    End If


    On Error GoTo setuperr
    bMoving = True

    If Not mdifrm Is Nothing And Not frm Is Nothing Then
        Set picbar = mdifrm.Controls.Add("VB.PictureBox", "picbar" & frm.hwnd)
        Set picbox = mdifrm.Controls.Add("VB.PictureBox", "picbox" & frm.hwnd)
    End If
    picbar.AutoRedraw = True
    picbar.BorderStyle = 0
    picbar.Appearance = 0
    
    picbox.AutoRedraw = True
    picbox.BorderStyle = 0
    picbox.Appearance = 0
    
    picbar.Height = (3 * Screen.TwipsPerPixelY)
    picbar.width = (3 * Screen.TwipsPerPixelX)
    
    picbar.BackColor = frm.BackColor
    picbox.BackColor = frm.BackColor
    
    bMoving = True
    picbox.Height = lDockedHeight
    picbox.width = lDockedWidth
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

        
        If Not mdifrm Is Nothing Then
            Call SetWindowWord(frm.hwnd, SWW_HPARENT, mdifrm.hwnd)
        End If
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

Private Function FindMDIWindow() As Long

End Function
Public Sub ToggleDocked(ByVal IsDocked As Boolean, Optional ByVal Override As Boolean = False)
    bMoving = True And (Not Override)
    bDocked = IsDocked
    If Not mdifrm Is Nothing Then
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





