VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Field"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'a stored energy between a space
'and a vision, projecting render
'purpose like a viewport a draw
Private myBackBuffer As Long
Private myBufferBMP As Long

Private pInterest As Stars
Private pWaypoint As Vision
    
Private hDC As Long
Private hpen As Long
Private hbrush As Long
Private Brush As LOGBRUSH
Private Pen As LOGPEN
    
Friend Property Get Interest() As Stars
    Set Interest = pInterest
End Property
Friend Property Set Interest(ByRef RHS As Stars)
    Set pInterest = RHS
End Property
Friend Property Get Waypoint() As Vision
    Set Waypoint = pWaypoint
End Property
Friend Property Set Waypoint(ByRef RHS As Vision)
    Set pWaypoint = RHS
End Property

Public Property Get Stars() As Stars
    Set Stars = pInterest
End Property
Public Property Set Stars(ByRef RHS As Stars)
    Set pInterest = RHS
End Property

Public Property Get Vision() As Vision
    Set Vision = pWaypoint
End Property
Public Property Set Vision(ByRef RHS As Vision)
    Set pWaypoint = RHS
End Property

Public Property Get hWnd() As Long
    hWnd = pWaypoint.Dtp
End Property
Friend Property Let hWnd(ByRef RHS As Long)
    pWaypoint.Dtp = RHS
    If (IsWindow(RHS) = 1) Then
        Dim rct As RECT
        If (hWnd <> 0) Then

            ReleaseDC hWnd, hDC
            DeleteObject myBufferBMP
            DeleteDC myBackBuffer

            HookObj Me
            If (RHS = 0) Then
                rct.Bottom = Screen.Height
                rct.Right = Screen.Width
            End If
        End If
        pWaypoint.Dtp = RHS
        If (RHS <> 0) Then
           'HookObj Me
            GetWindowRect RHS, rct
        End If

        myBackBuffer = CreateCompatibleDC(GetDC(0))
        myBufferBMP = CreateCompatibleBitmap(GetDC(0), rct.Right, rct.Bottom)
        SelectObject myBackBuffer, myBufferBMP

        hDC = GetDC(hWnd)

        hpen = CreatePenIndirect(Pen)
        SelectObject hDC, hpen

        hbrush = CreateBrushIndirect(Brush)
        SelectObject hDC, hbrush
    End If
End Property

Friend Function ControlWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Debug.Print TypeName(HookObj(hWnd)) & ", " & hWnd & ", " & uMsg & ", " & wParam & ", " & lParam

End Function

Friend Sub Frame()
    Dim rct As RECT
    GetWindowRect hWnd, rct
    Dim hDC As Long
    hDC = CreateCompatibleDC(GetDC(hWnd))
    BitBlt hDC, 0, 0, rct.Right, rct.Bottom, myBackBuffer, 0, 0, vbSrcCopy
    DeleteDC hDC
End Sub

Friend Sub Clear(Optional ByRef Color As Long = vbBlack)
    If Not (Pen.lopnColor = Color) Then Pen.lopnColor = Color
    If Not (Brush.lbColor = Color) Then Brush.lbColor = Color
    Dim rct As RECT
    GetWindowRect hWnd, rct
    Rectangle hDC, 0, 0, rct.Right, rct.Bottom
End Sub

Friend Function Equals(ByRef SameAs) As Boolean
    Equals = ((pInterest = SameAs.Interest) And pWaypoint.Equals(SameAs.Waypoint))
End Function

Friend Function Clone() As Field
    Set Clone = New Field
    Set Clone.Interest = pInterest.Clone
    Set Clone.Waypoint = pWaypoint.Clone
End Function

Public Property Get Serialize() As String
    Serialize = pInterest.Serialize() & "$" & pWaypoint.Serialize()
End Property
Public Property Let Serialize(ByVal RHS As String)
    If (pInterest Is Nothing) Then Set pInterest = New Stars
    pInterest.Serialize = NextArg(RHS, "$")
    If (pWaypoint Is Nothing) Then Set pWaypoint = New Vision
    pWaypoint.Serialize = RemoveArg(RHS, "$")
End Property

Private Sub Class_Initialize()

    With Pen
        .lopnColor = vbBlack
        .lopnStyle = PS_SOLID
        .lopnWidth.X = 1
    End With

    With Brush
        .lbColor = vbBlack
        .lbStyle = PS_SOLID
        .lbHatch = HS_SOLID
    End With

End Sub

Private Sub Class_Terminate()


    DeleteObject hbrush
    DeleteObject hpen

    ReleaseDC hWnd, hDC
    DeleteObject myBufferBMP
    DeleteDC myBackBuffer

    If (hWnd <> 0) Then HookObj Me
    
    Set pWaypoint.Interest = Nothing
    
    Set pInterest = Nothing
    Set pWaypoint = Nothing
    
End Sub

