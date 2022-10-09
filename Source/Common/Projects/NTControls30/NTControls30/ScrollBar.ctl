VERSION 5.00
Begin VB.UserControl ScrollBar 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   645
   MousePointer    =   1  'Arrow
   ScaleHeight     =   570
   ScaleWidth      =   645
   ToolboxBitmap   =   "ScrollBar.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   105
      Top             =   60
   End
End
Attribute VB_Name = "ScrollBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum vbOrientation
    OrientationAuto = -1
    OrientationHorizontal = 1
    OrientationVertical = 0
End Enum

Public Event Change()
Public Event Scroll()
Public Event Click()
Public Event DblClick()

Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLECompleteDrag(Effect As Long)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Public Event Show()
Public Event Hide()
Public Event Resize()
Public Event Paint()

'drawing rectangles
Private rCanvas As Rect 'whole control
Private rButton1 As Rect 'left or top button
Private rButton2 As Rect 'right or bottom button
Private rSlider As Rect 'just the thumb bar part
Private rScroll As Rect 'everything but buttons

'internal systemmetric
Private keyDelay As Long
Private keySpeed As Long

'what part of the scrollbar
'mouse events are occuring
Private pHitRegion As Long

'values to save the initial mouse
'events we'll analyze or repeat
Private pPressed As Integer
Private pShift As Integer
Private pEventX As Single
Private pEventY As Single

'specific values for the thumb bar
'during events of the mouse/repeat
Private pThumbValue As Long
Private pThumbX As Single
Private pThumbY As Single

'the containers for public properties
Private pOrientation As vbOrientation
'Private pProportionalThumb As Boolean
Private pIncremental As Boolean
Private pSmallChange As Long
Private pLargeChange As Long
Private pMin As Long
Private pMax As Long
Private pValue As Long
Private pTag

'Public Property Get ProportionalThumb() As Boolean
'    ProportionalThumb = pProportionalThumb
'End Property
'Public Property Let ProportionalThumb(ByVal RHS As Boolean)
'    pProportionalThumb = RHS
'End Property
Public Property Get Incremental() As Boolean ' _
Gets whether or not the scroll bar's value is forced to SmallChange increments only.
Attribute Incremental.VB_Description = "Gets whether or not the scroll bar's value is forced to SmallChange increments only."
    Incremental = pIncremental
End Property
Public Property Let Incremental(ByVal RHS As Boolean) ' _
Sets whether or not the scoll bar's value should be forced to SmallChange increments only.
Attribute Incremental.VB_Description = "Sets whether or not the scoll bar's value should be forced to SmallChange increments only."
    pIncremental = RHS
End Property
Private Function IsHorizontal() As Boolean
    IsHorizontal = ((pOrientation = vbOrientation.OrientationAuto And UserControl.Height < UserControl.Width) Or pOrientation = vbOrientation.OrientationHorizontal)
End Function
Private Function IsVertical() As Boolean
    IsVertical = ((pOrientation = vbOrientation.OrientationAuto And UserControl.Width < UserControl.Height) Or pOrientation = vbOrientation.OrientationVertical)
End Function

Friend Sub UpdateRects()
    rCanvas.Right = (UserControl.Width / Screen.TwipsPerPixelX)
    rCanvas.Bottom = (UserControl.Height / Screen.TwipsPerPixelY)
    If IsHorizontal Then
        rButton1.Bottom = rCanvas.Bottom
        rButton2.Bottom = rButton1.Bottom
        rButton1.Right = rButton1.Bottom
        rButton2.Right = rCanvas.Right
        rButton2.Left = rButton2.Right - rButton1.Right
        rButton1.Top = 0
        rButton2.Top = 0
        rScroll.Left = rButton1.Right
        rScroll.Top = 0
        rScroll.Right = rButton2.Left
        rScroll.Bottom = rButton1.Bottom
    ElseIf IsVertical Then
        rButton1.Right = rCanvas.Right
        rButton2.Right = rButton1.Right
        rButton1.Bottom = rButton1.Right
        rButton2.Bottom = rCanvas.Bottom
        rButton2.Top = rButton2.Bottom - rButton1.Bottom
        rButton1.Left = 0
        rButton2.Left = 0
        rScroll.Right = rButton1.Bottom
        rScroll.Left = 0
        rScroll.Bottom = rButton2.Top
        rScroll.Right = rButton1.Right
    End If
    
    rSlider = GetSliderRect
End Sub

Private Function GetSliderRect() As Rect
        
    If IsHorizontal Then

        With GetSliderRect
            .Left = rCanvas.Bottom
            .Right = rCanvas.Right - .Left
            .Bottom = .Left
            
            If pThumbValue <> 0 Then 'draw at thumb locatoin
                If .Left + (pThumbValue / Screen.TwipsPerPixelX) > .Right - rCanvas.Bottom Then
                    .Left = (.Right - rCanvas.Bottom)
                ElseIf Not (.Left + (pThumbValue / Screen.TwipsPerPixelX) < .Left) Then
                    .Left = (.Left + (pThumbValue / Screen.TwipsPerPixelX))
                End If
                
            ElseIf ScrollAmount > 0 Then 'draw at value location
                .Left = .Left + (((((.Right * Screen.TwipsPerPixelX) - ((rCanvas.Bottom * _
                    Screen.TwipsPerPixelX) * 2)) / ScrollAmount) * pValue) / Screen.TwipsPerPixelX)
            End If
                
            .Right = .Left + rCanvas.Bottom
            
        End With
    ElseIf IsVertical Then

        With GetSliderRect
            .Top = rCanvas.Right
            .Right = .Top
            .Bottom = rCanvas.Bottom - .Top
            
            If pThumbValue <> 0 Then 'draw at thumb locatoin
                If .Top + (pThumbValue / Screen.TwipsPerPixelY) > .Bottom - rCanvas.Right Then
                    .Top = (.Bottom - rCanvas.Right)
                ElseIf Not (.Top + (pThumbValue / Screen.TwipsPerPixelY) < .Top) Then
                    .Top = (.Top + (pThumbValue / Screen.TwipsPerPixelY))
                End If
                
            ElseIf ScrollAmount > 0 Then 'draw at value location
                .Top = .Top + (((((.Bottom * Screen.TwipsPerPixelY) - ((rCanvas.Right * _
                    Screen.TwipsPerPixelY) * 2)) / ScrollAmount) * pValue) / Screen.TwipsPerPixelY)
            End If
                
            .Bottom = .Top + rCanvas.Right
        End With
    End If

End Function
Private Function ScrollAmount(Optional ByVal InScreen As Boolean = False) As Single
    If Not InScreen Then
        If pMax > pMin Then
            ScrollAmount = (pMax - pMin)
        End If
    Else
        If IsHorizontal Then
            ScrollAmount = (rScroll.Right - rScroll.Left)
        ElseIf IsVertical Then
            ScrollAmount = (rScroll.Bottom - rScroll.Top)
        End If
    End If

End Function

Private Function ScrollableSpace(Optional ByVal ExcludeSlider As Boolean = False) As Single
    If IsHorizontal Then
        ScrollableSpace = ((rCanvas.Right - (rCanvas.Bottom * 2)) * Screen.TwipsPerPixelX)
    ElseIf IsVertical Then
        ScrollableSpace = ((rCanvas.Bottom - (rCanvas.Right * 2)) * Screen.TwipsPerPixelY)
    End If
    If Not ExcludeSlider Then
        If IsHorizontal Then
            ScrollableSpace = ScrollableSpace - ((rSlider.Right - rSlider.Left) * Screen.TwipsPerPixelX)
        ElseIf IsVertical Then
            ScrollableSpace = ScrollableSpace - ((rSlider.Bottom - rSlider.Top) * Screen.TwipsPerPixelY)
        End If
    End If
End Function

Public Property Get AutoRedraw() As Boolean ' _
Gets whether or not the scroll bar automatically redraws itself.
Attribute AutoRedraw.VB_Description = "Gets whether or not the scroll bar automatically redraws itself."
    AutoRedraw = UserControl.AutoRedraw
End Property
Public Property Let AutoRedraw(ByVal RHS As Boolean) ' _
Sets whether or not the scroll bar automaticall redraws itself.
Attribute AutoRedraw.VB_Description = "Sets whether or not the scroll bar automaticall redraws itself."
    UserControl.AutoRedraw = RHS
    UserControl_Paint
End Property

Public Property Get hWnd() As Long ' _
Returns the standard windows handle to the control.
Attribute hWnd.VB_Description = "Returns the standard windows handle to the control."
    hWnd = UserControl.hWnd
End Property

Public Property Get SmallChange() As Long ' _
Gets value at which the scroll bar increments when an arrow button is clicked.
Attribute SmallChange.VB_Description = "Gets value at which the scroll bar increments when an arrow button is clicked."
    SmallChange = pSmallChange
End Property
Public Property Let SmallChange(ByVal RHS As Long) ' _
Sets the value at which the scroll bar increments when an arrow button is clicked.
Attribute SmallChange.VB_Description = "Sets the value at which the scroll bar increments when an arrow button is clicked."
    If pSmallChange <> RHS And RHS <> 0 Then
        pSmallChange = RHS
        If pLargeChange < pSmallChange Then pLargeChange = pSmallChange
        UserControl_Paint
    End If
End Property
Public Property Get LargeChange() As Long ' _
Gets the value at which the scroll bar increments when clicking in the scrollable area.
Attribute LargeChange.VB_Description = "Gets the value at which the scroll bar increments when clicking in the scrollable area."
    LargeChange = pLargeChange
End Property
Public Property Let LargeChange(ByVal RHS As Long) ' _
Sets the value at which the scroll bar increments when clicking in the scrollable area.
Attribute LargeChange.VB_Description = "Sets the value at which the scroll bar increments when clicking in the scrollable area."
    If pLargeChange <> RHS And RHS <> 0 Then
        pLargeChange = RHS
        If pLargeChange < pSmallChange Then pSmallChange = pLargeChange
        UserControl_Paint
    End If
End Property
Public Property Get Min() As Long ' _
Gets the minimum value restriction of the scroll bar.
Attribute Min.VB_Description = "Gets the minimum value restriction of the scroll bar."
    Min = pMin
End Property
Public Property Let Min(ByVal RHS As Long) ' _
Sets the minimum value restriction of the scroll bar.
Attribute Min.VB_Description = "Sets the minimum value restriction of the scroll bar."
    If pMin <> RHS Then
        pMin = RHS
        If pValue < pMin Then
            Value = pMin
        Else
            UserControl_Paint
        End If
    End If
End Property
Public Property Get Max() As Long ' _
Gets the maximum value restriction of the scroll bar.
Attribute Max.VB_Description = "Gets the maximum value restriction of the scroll bar."
    Max = pMax
End Property
Public Property Let Max(ByVal RHS As Long) ' _
Sets the maximum value restriction of the scroll bar.
Attribute Max.VB_Description = "Sets the maximum value restriction of the scroll bar."
    If pMax <> RHS Then
        pMax = RHS
        If pValue > pMax Then
            Value = pMax
        Else
            UserControl_Paint
        End If
    End If
End Property

Public Property Get Value() As Long ' _
Gets the current value of the scroll bar's slider.
Attribute Value.VB_Description = "Gets the current value of the scroll bar's slider."
    If pThumbValue <> 0 Then
        Value = ((pThumbValue / (ScrollableSpace / ScrollAmount)))
        If Value < pMin Then
            Value = pMin
        End If
        If Value > pMax Then
            Value = pMax
        End If
    Else
        Value = pValue
    End If
    If pIncremental Then
        Value = (Value \ pSmallChange) * pSmallChange
    End If
End Property
Public Property Let Value(ByVal RHS As Long) ' _
Sets the current value of the scroll bar's slider.
Attribute Value.VB_Description = "Sets the current value of the scroll bar's slider."
    Static lastValue As Long
    If RHS < pMin Then
        RHS = pMin
    End If
    If RHS > pMax Then
        RHS = pMax
    End If
    If RHS <> lastValue Then
        lastValue = RHS
        pValue = RHS
        RaiseEvent Change
        UserControl_Paint
    End If
End Property

Public Property Get Orientation() As vbOrientation ' _
Gets the orientation of the scroll bar, whether vertical, horizontal or automatic.
Attribute Orientation.VB_Description = "Gets the orientation of the scroll bar, whether vertical, horizontal or automatic."
    Orientation = pOrientation
End Property
Public Property Let Orientation(ByRef RHS As vbOrientation) ' _
Sets the orientation of the scroll bar, whehter vertical, horizontal or automatic.
Attribute Orientation.VB_Description = "Sets the orientation of the scroll bar, whehter vertical, horizontal or automatic."
    If pOrientation <> RHS Then
        pOrientation = RHS
        UserControl_Paint
    End If
End Property

Public Property Get Enabled() As Boolean ' _
Gets whether or not the control is greyed out, disallowing user interactions.
Attribute Enabled.VB_Description = "Gets whether or not the control is greyed out, disallowing user interactions."
    Enabled = UserControl.Enabled And ScrollAmount > 0
End Property
Public Property Let Enabled(ByVal RHS As Boolean) ' _
Sets whether or not the control is greyed out, disallowing user interactions.
Attribute Enabled.VB_Description = "Sets whether or not the control is greyed out, disallowing user interactions."
    If UserControl.Enabled <> RHS Then
        UserControl.Enabled = RHS
        UserControl_Paint
    End If
End Property

Public Property Get Tag() ' _
Gets the non specific user defined datatype belonging to this control instance.
Attribute Tag.VB_Description = "Gets the non specific user defined datatype belonging to this control instance."
    If IsObject(pTag) Then
        Set Tag = pTag
    Else
        Tag = pTag
    End If
End Property
Public Property Let Tag(ByVal RHS) ' _
Sets the non specific user defined datatype belonging to this control instance.
Attribute Tag.VB_Description = "Sets the non specific user defined datatype belonging to this control instance."
    pTag = RHS
End Property
Public Property Set Tag(ByRef RHS) ' _
Sets the non specific user defined datatype belonging to this control instance.
Attribute Tag.VB_Description = "Sets the non specific user defined datatype belonging to this control instance."
    Set pTag = RHS
End Property

Private Sub Timer1_Timer()

    Timer1.Interval = keySpeed
    
    UserControl_MouseDown pPressed, pShift, pEventX, pEventY
    
    Timer1.Enabled = (pPressed <> 0)
    If Not Timer1.Enabled Then Timer1.Interval = keySpeed * 10
    
    UserControl_Paint
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Hide()
    RaiseEvent Hide
End Sub

Private Sub UserControl_Initialize()
    SystemParametersInfo SPI_GETKEYBOARDSPEED, 0, keySpeed, 0
    Timer1.Interval = keySpeed * 10
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Enabled Then

       
        pPressed = Button
        pShift = Shift
        pEventX = X
        pEventY = Y

        If Button = 1 Then
            
            rSlider = GetSliderRect
            
            If PtInRect(rButton1, X / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY) Then
                pHitRegion = 1 'top or left most button / smallchange
                If pValue > pMin And pValue - pSmallChange >= pMin Then
                    Value = pValue - pSmallChange
                Else
                    Value = pMin
                End If
            ElseIf PtInRect(rButton2, X / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY) Then
                pHitRegion = 2 'bottom or right most button / smallchange
                If pValue < pMax And pValue + pSmallChange <= pMax Then
                    Value = pValue + pSmallChange
                Else
                    Value = pMax
                End If
            ElseIf PtInRect(rSlider, X / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY) Then
                If pHitRegion = 3 Then
                    'this is the very end of holding the mouse until silder contacts it
                    If IsHorizontal Then
                        pThumbValue = ((ScrollableSpace / ScrollAmount) * pValue)
                    ElseIf IsVertical Then
                        pThumbValue = ((ScrollableSpace / ScrollAmount) * pValue)
                    End If
                End If
                
                pHitRegion = 4 'the slider

                If pThumbValue = 0 Then
                    pThumbValue = ((ScrollableSpace(True) / ScrollAmount) * pValue)
                    
                    pThumbX = X
                    pThumbY = Y
                End If
                
                
            ElseIf PtInRect(rScroll, X / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY) Then
                pHitRegion = 3 'scrollable space
            
                Dim tmpVal As Single
                If IsHorizontal Then
                    If ScrollAmount > 0 Then
                        If (ScrollableSpace / ScrollAmount) > 0 Then
                            tmpVal = (((X - UserControl.Height) / (ScrollableSpace / ScrollAmount)))
                            If ((tmpVal < pValue) And (tmpVal < pMin)) Then tmpVal = pMin
                            If ((tmpVal > pValue) And (tmpVal > pMax)) Then tmpVal = pMax
                            If X / Screen.TwipsPerPixelX < rSlider.Left And Y / Screen.TwipsPerPixelY > rSlider.Top And Y / Screen.TwipsPerPixelY < rSlider.Bottom Then
                                'left side of scrollable space
                                If pValue > tmpVal And pValue - pLargeChange >= tmpVal Then
                                    Value = pValue - pLargeChange
                                Else
                                    Value = tmpVal
                                End If
                            ElseIf X / Screen.TwipsPerPixelX > rSlider.Right And Y / Screen.TwipsPerPixelY > rSlider.Top And Y / Screen.TwipsPerPixelY < rSlider.Bottom Then
                                'right side of scrollable space
                                If pValue < tmpVal And pValue + pLargeChange <= tmpVal Then
                                    Value = pValue + pLargeChange
                                Else
                                    Value = tmpVal
                                End If
                            End If
                        End If
                    End If
                ElseIf IsVertical Then
                    If ScrollAmount > 0 Then
                        If (ScrollableSpace / ScrollAmount) > 0 Then
                            tmpVal = (((Y - UserControl.Width) / (ScrollableSpace / ScrollAmount)))
                            If (tmpVal < pValue) And (tmpVal < pMin) Then tmpVal = pMin
                            If (tmpVal > pValue) And (tmpVal > pMax) Then tmpVal = pMax
                            If Y / Screen.TwipsPerPixelY < rSlider.Top And X / Screen.TwipsPerPixelX > rSlider.Left And X / Screen.TwipsPerPixelX < rSlider.Right Then
                                'top of scrollable space
                                If pValue > tmpVal And pValue - pLargeChange >= tmpVal Then
                                    Value = pValue - pLargeChange
                                Else
                                    Value = tmpVal
                                End If
                            ElseIf Y / Screen.TwipsPerPixelY > rSlider.Bottom And X / Screen.TwipsPerPixelX > rSlider.Left And X / Screen.TwipsPerPixelX < rSlider.Right Then
                                'bottom of scrollable space
                                If pValue < tmpVal And pValue + pLargeChange <= tmpVal Then
                                    Value = pValue + pLargeChange
                                Else
                                    Value = tmpVal
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        
        End If

        Timer1.Enabled = ((pPressed <> 0) And (pHitRegion <> 4))
       
    End If
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Enabled Then
        pPressed = Button
        If pPressed = 0 Then
            If Not pHitRegion = 0 Then
                pHitRegion = 0
                Timer1.Enabled = False
                UserControl_Paint
            End If
        ElseIf Button = 1 Then
        
            Dim tmpRct As Rect
            tmpRct = GetSliderRect
            If tmpRct.Left <> rSlider.Left Or tmpRct.Top <> rSlider.Top Or tmpRct.Bottom <> rSlider.Bottom Or tmpRct.Right <> rSlider.Right Then
                rSlider = tmpRct
                UserControl_Paint
                RaiseEvent Scroll
            End If

            
            If pHitRegion = 4 Then
                If ScrollAmount > 0 Then
                    If IsHorizontal Then
                        pThumbValue = (X - pThumbX) + ((ScrollableSpace(True) / ScrollAmount) * pValue)
                    Else
                        pThumbValue = (Y - pThumbY) + ((ScrollableSpace(True) / ScrollAmount) * pValue)
                    End If
                End If
            End If
           
        End If
    End If
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Enabled Then
        pPressed = Button

        If pThumbValue <> 0 Then
            'commit to change
            rSlider = GetSliderRect
            If ScrollAmount > 0 Then
                Value = ((pThumbValue / (ScrollableSpace / ScrollAmount)))
            End If
            pThumbValue = 0
        End If

        If pPressed = 0 Then
            If Not pHitRegion = 0 Then
                pHitRegion = 0
                Timer1.Enabled = False
                UserControl_Paint
            End If
        End If
    End If
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_Paint()
    UpdateRects
    If AutoRedraw Then
        If IsHorizontal Then
            DrawhScrollBar
        ElseIf IsVertical Then
            DrawvScrollBar
        End If
    End If
    RaiseEvent Paint
End Sub

Public Sub Paint() ' _
Redraws the control, this is preformed automatically when AutoRedraw is true.
Attribute Paint.VB_Description = "Redraws the control, this is preformed automatically when AutoRedraw is true."
    UpdateRects
    If IsHorizontal Then
        DrawhScrollBar
    ElseIf IsVertical Then
        DrawvScrollBar
    End If
End Sub

Private Sub DrawvScrollBar()

    'draw scrollable area backdrop
    If Enabled Then
        If rSlider.Top - rButton1.Bottom > 0 Then
            DrawFrameControl hdc, Rect(rCanvas.Left - 1, rButton1.Bottom - 1, rCanvas.Right + 1, rSlider.Top + 1), DFC_BUTTON, DFCS_BUTTONCHECK Or Not DFCS_MONO
        End If
        If rButton2.Top - rSlider.Bottom > 0 Then
            DrawFrameControl hdc, Rect(rCanvas.Left - 1, rSlider.Bottom - 1, rCanvas.Right + 1, rButton2.Top + 1), DFC_BUTTON, DFCS_BUTTONCHECK Or Not DFCS_MONO
        End If
    Else
        DrawFrameControl hdc, Rect(rCanvas.Left - 1, rButton1.Bottom - 1, rCanvas.Right + 1, rButton2.Top + 1), DFC_BUTTON, DFCS_BUTTONCHECK Or Not DFCS_MONO
    End If
    
    'draw the arrow buttons
    If rCanvas.Bottom - rCanvas.Top > rCanvas.Right * 2 Then
        DrawFrameControl hdc, rButton1, DFC_SCROLL, DFCS_SCROLLUP Or IIf(pHitRegion = 1 And Enabled, DFCS_PUSHED, 0) Or IIf(Enabled, 0, DFCS_INACTIVE)
        DrawFrameControl hdc, rButton2, DFC_SCROLL, DFCS_SCROLLDOWN Or IIf(pHitRegion = 2 And Enabled, DFCS_PUSHED, 0) Or IIf(Enabled, 0, DFCS_INACTIVE)
    ElseIf rCanvas.Bottom - rCanvas.Top > 0 Then
        DrawFrameControl hdc, rButton1, DFC_SCROLL, DFCS_SCROLLUP Or IIf(pHitRegion = 1 And Enabled, DFCS_PUSHED, 0) Or IIf(Enabled, 0, DFCS_INACTIVE)
        DrawFrameControl hdc, rButton2, DFC_SCROLL, DFCS_SCROLLDOWN Or IIf(pHitRegion = 2 And Enabled, DFCS_PUSHED, 0) Or IIf(Enabled, 0, DFCS_INACTIVE)
    End If

    If ScrollAmount > 0 And Enabled Then
        'draw the slider bar
        If rCanvas.Bottom - rCanvas.Top > rCanvas.Right * 3 - (rCanvas.Right / 2) Then
            DrawFrameControl hdc, rSlider, DFC_BUTTON, DFCS_BUTTONPUSH
        ElseIf rCanvas.Bottom - rCanvas.Top > rCanvas.Right * 2 Then
            DrawFrameControl hdc, Rect(rCanvas.Left, rCanvas.Top + rCanvas.Right, rCanvas.Right, rCanvas.Bottom - rCanvas.Right), DFC_BUTTON, DFCS_BUTTONPUSH
        End If
    End If

End Sub

Private Sub DrawhScrollBar()

    'draw scrollable area backdrop
    If Enabled Then
        If rSlider.Left - rButton1.Right > 0 Then
            DrawFrameControl hdc, Rect(rButton1.Right - 1, rCanvas.Top - 1, rSlider.Left + 1, rCanvas.Bottom + 1), DFC_BUTTON, DFCS_BUTTONCHECK Or Not DFCS_MONO
        End If
        If rButton2.Left - rSlider.Right > 0 Then
            DrawFrameControl hdc, Rect(rSlider.Right - 1, rCanvas.Top - 1, rButton2.Left + 1, rCanvas.Bottom + 1), DFC_BUTTON, DFCS_BUTTONCHECK Or Not DFCS_MONO
        End If
    Else
        DrawFrameControl hdc, Rect(rButton1.Right - 1, rCanvas.Top - 1, rButton2.Left + 1, rCanvas.Bottom + 1), DFC_BUTTON, DFCS_BUTTONCHECK Or Not DFCS_MONO
    End If
    
    'draw the arrow buttons
    If rCanvas.Right - rCanvas.Left > rCanvas.Bottom * 2 Then
        DrawFrameControl hdc, rButton1, DFC_SCROLL, DFCS_SCROLLLEFT Or IIf(pHitRegion = 1 And Enabled, DFCS_PUSHED, 0) Or IIf(Enabled, 0, DFCS_INACTIVE)
        DrawFrameControl hdc, rButton2, DFC_SCROLL, DFCS_SCROLLRIGHT Or IIf(pHitRegion = 2 And Enabled, DFCS_PUSHED, 0) Or IIf(Enabled, 0, DFCS_INACTIVE)
    ElseIf rCanvas.Right - rCanvas.Left > 0 Then
        DrawFrameControl hdc, rButton1, DFC_SCROLL, DFCS_SCROLLLEFT Or IIf(pHitRegion = 1 And Enabled, DFCS_PUSHED, 0) Or IIf(Enabled, 0, DFCS_INACTIVE)
        DrawFrameControl hdc, rButton2, DFC_SCROLL, DFCS_SCROLLRIGHT Or IIf(pHitRegion = 2 And Enabled, DFCS_PUSHED, 0) Or IIf(Enabled, 0, DFCS_INACTIVE)
    End If

    If ScrollAmount > 0 And Enabled Then
        'draw the slider bar
        If rCanvas.Right - rCanvas.Left > rCanvas.Bottom * 3 - (rCanvas.Bottom / 2) Then
            DrawFrameControl hdc, rSlider, DFC_BUTTON, DFCS_BUTTONPUSH
        ElseIf rCanvas.Right - rCanvas.Left > rCanvas.Bottom * 2 Then
            DrawFrameControl hdc, Rect(rCanvas.Left + rCanvas.Bottom, rCanvas.Top, rCanvas.Right - rCanvas.Bottom, rCanvas.Bottom), DFC_BUTTON, DFCS_BUTTONPUSH
        End If
    End If
    
End Sub

Private Sub UserControl_Resize()
    UserControl_Paint
    RaiseEvent Resize
End Sub

Private Sub UserControl_Show()
    RaiseEvent Show
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Orientation = PropBag.ReadProperty("Orientation", vbOrientation.OrientationAuto)
    Enabled = PropBag.ReadProperty("Enabled", True)
    Tag = PropBag.ReadProperty("Tag", Empty)
    
    Value = PropBag.ReadProperty("Value", 0)
    Min = PropBag.ReadProperty("Min", 0)
    Max = PropBag.ReadProperty("Max", 100)
    SmallChange = PropBag.ReadProperty("SmallChange", 1)
    LargeChange = PropBag.ReadProperty("LargeChange", 4)
    AutoRedraw = PropBag.ReadProperty("AutoRedraw", True)
    Incremental = PropBag.ReadProperty("Incremental", True)
    'ProportionalThumb = PropBag.ReadProperty("ProportionalThumb", True)
    UpdateRects
    Paint
End Sub

Private Sub UserControl_InitProperties()
    Orientation = vbOrientation.OrientationAuto
    Enabled = True
    Max = 100
    SmallChange = 1
    LargeChange = 4
    AutoRedraw = True
    Incremental = True
    'ProportionalThumb = True
    UpdateRects
    Paint
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    PropBag.WriteProperty "Orientation", Orientation, vbOrientation.OrientationAuto
    PropBag.WriteProperty "Enabled", Enabled, True
    PropBag.WriteProperty "Tag", Tag, Empty
    
    PropBag.WriteProperty "Value", Value, 0
    PropBag.WriteProperty "Min", Min, 0
    PropBag.WriteProperty "Max", Max, 100
    PropBag.WriteProperty "SmallChange", SmallChange, 1
    PropBag.WriteProperty "LargeChange", LargeChange, 4
    PropBag.WriteProperty "AutoRedraw", AutoRedraw, True
    PropBag.WriteProperty "Incremental", Incremental, True
    'PropBag.WriteProperty "ProportionalThumb", ProportionalThumb
    
End Sub
