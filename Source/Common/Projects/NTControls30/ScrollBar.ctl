VERSION 5.00
Begin VB.UserControl ScrollBar 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   645
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MousePointer    =   1  'Arrow
   ScaleHeight     =   38
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   43
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

Implements IControl

Public Enum vbOrientation
    OrientationAuto = -1
    OrientationHorizontal = 1
    OrientationVertical = 0
End Enum

Private Enum HitRegions
    None = 0
    ScrollButton1 = 1
    ScrollButton2 = 2
    SliderButton = 3
    ScrollableArea = 4
End Enum

Public Event Change() ' _
Event that is fired whenever the scroll Value has changed.

Public Event Scroll() ' _
Event that is fired only when the slider has been moved ny the user.

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
Private rCanvas As RECT 'whole control boundry
Private rButton1 As RECT 'left or top 1st button
Private rButton2 As RECT 'right or bottom 2nd button
Private rSlider As RECT 'thumb bar or middle 3rd button
Private rScroll As RECT 'thumb bar boundary

'internal systemmetric
Private keyDelay As Long
Private keySpeed As Long

'what part of the scrollbar
'mouse events are occuring
Private pHitRegion As HitRegions

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
Private pProportionalThumb As Boolean

Private pSmallChange As Long
Private pLargeChange As Long
Private pMin As Long
Private pMax As Long
Private pValue As Long
Private pTag
Private pOldProc As Long

Private pBackBuffer As Backbuffer


Public Property Get Backbuffer() As Backbuffer
    Set Backbuffer = pBackBuffer
End Property
Public Property Set Backbuffer(ByRef RHS As Backbuffer)
    Set pBackBuffer = RHS
End Property

Friend Property Get hProc() As Long
    hProc = pOldProc
End Property
Friend Property Let hProc(ByVal RHS As Long)
    pOldProc = RHS
End Property

Public Property Get ProportionalThumb() As Boolean
    ProportionalThumb = pProportionalThumb
End Property
Public Property Let ProportionalThumb(ByVal RHS As Boolean)
    pProportionalThumb = RHS
End Property

Private Function IsHorizontal() As Boolean
    IsHorizontal = ((pOrientation = vbOrientation.OrientationAuto And UserControl.ScaleHeight < UserControl.ScaleWidth) Or pOrientation = vbOrientation.OrientationHorizontal)
End Function
Private Function IsVertical() As Boolean
    IsVertical = ((pOrientation = vbOrientation.OrientationAuto And UserControl.ScaleWidth < UserControl.ScaleHeight) Or pOrientation = vbOrientation.OrientationVertical)
End Function

Friend Sub UpdateRects()

    rCanvas.Left = 0
    rCanvas.Top = 0
    rCanvas.Right = UserControl.ScaleWidth
    rCanvas.Bottom = UserControl.ScaleHeight
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

Private Function GetSliderRect() As RECT
    Dim tmpRct As RECT
    Dim tmpVal As Single
    Dim fullamt As Single

    tmpRct = rSlider
    
    If IsHorizontal Then
        
        With tmpRct
            .Left = rButton1.Right
            .Top = 0
            .Bottom = rButton1.Bottom
            
            If ScrollAmount(True) > 0 Then
                fullamt = ScrollAmount
                If fullamt > 0 Then
                    If pValue > 0 Then
                        .Left = rButton1.Right + ((pValue / fullamt) * ScrollableSpace)
                    End If
                    
                    If pProportionalThumb Then
                        'Debug.Print (ScrollAmount(False) / ScrollableSpace(True)); UserControl.Width / Screen.Width
                        
                        .Right = (.Left + ((UserControl.ScaleWidth * UserControl.Width / Screen.Width)))
                    Else
                        .Right = (.Left + (rButton1.Bottom - rButton1.Top))
                    End If
        
                    If (.Right - .Left) < ((rButton1.Bottom - rButton1.Top) / 2) Then
                        .Right = .Left + ((rButton1.Bottom - rButton1.Top) / 2)
                    End If
                    If (.Left < rButton1.Left) Then .Left = rCanvas.Left
                    If (.Right > rButton2.Left) Then .Left = rButton2.Left - (.Right - .Left)
            
                Else
                    .Left = 0
                    .Right = 0
                End If
            End If

        End With
        
    ElseIf IsVertical Then
        
        With tmpRct
          
            .Top = rButton1.Bottom
            .Left = 0
            .Right = rButton1.Right

            If ScrollAmount(True) > 0 Then
                fullamt = ScrollAmount
                If fullamt > 0 Then
                    If pValue > 0 Then
                        .Top = (rButton1.Bottom + ((pValue / fullamt) * ScrollableSpace))
                    End If
                    
                    If pProportionalThumb Then
                        Debug.Print (ScrollAmount(False) / ScrollableSpace(True)); UserControl.Height / Screen.Height
                        
                        .Bottom = (.Top + ((UserControl.ScaleHeight * UserControl.Height / Screen.Height)))
                    Else
                        .Bottom = (.Top + (rButton1.Right - rButton1.Left))
                    End If
        
                    If (.Bottom - .Top) < ((rButton1.Right - rButton1.Left) / 2) Then
                        .Bottom = .Top + ((rButton1.Right - rButton1.Left) / 2)
                    End If
                    If (.Top < rButton1.Top) Then .Top = rCanvas.Top
                    If (.Bottom > rButton2.Top) Then .Top = rButton2.Top - (.Bottom - .Top)
                Else
                    .Top = 0
                    .Bottom = 0
                End If

            End If

        End With
        
    End If
    
    GetSliderRect = tmpRct

End Function
            
'            If ScrollAmount > 0 Then
'                tmpVal = ((ScrollableSpace / ScrollAmount) * pValue)
'
'                ratioMoveable = (UserControl.ScaleWidth / ScrollAmount)
'
'
'                If pValue > 0 Then
'                    If (ScrollAmount(False) / pValue) > 0 Then
'                        .Left = ScrollableSpace / (ScrollAmount(False) / pValue)
'                    End If
'                End If
'
'            End If
'
'
'            ratioToScreen = (Screen.Width / UserControl.ScaleWidth)
'
'            Debug.Print ratioMoveable; ratioToScreen

            
'            If (ScrollAmount(True) - (.Bottom - .Top)) > 0 Then
'                .Left = tmpVal
'            End If

            
'           If pProportionalThumb And ScrollAmount(True) <> 0 Then
'               If (ScrollAmount(False) / ScrollAmount(True)) <> 0 Then
'                    fullamt = (1 - (ScrollableSpace / (ScrollableSpace / (ScrollAmount(False) / ScrollAmount(True)))))
'                    tmpVal = (rScroll.Right - rScroll.Left) - ((rScroll.Right - rScroll.Left) * fullamt)
'                Else
'                    tmpVal = 0
'                End If
'                If tmpVal < rCanvas.Bottom And tmpVal > 0 Then tmpVal = rScroll.Bottom - rScroll.Top
'            Else
'                tmpVal = rCanvas.Bottom
'            End If
'            If pThumbValue <> 0 Then 'draw at thumb locatoin
'                If .Left + pThumbValue > .Right - tmpVal Then
'                    .Left = (.Right - .Left)
'                ElseIf Not (.Left + pThumbValue < .Left) Then
'                    .Left = (.Left + pThumbValue )
'                End If
'            ElseIf ScrollAmount > 0 Then 'draw at value location
'                .Left = .Left + (((.Right - (rCanvas.Bottom * 2)) / ScrollAmount) * ((pValue + pMin) * fullamt))
'                If .Left + pThumbValue  > .Right - tmpVal Then .Left = (.Right - tmpVal)
'            End If
'
'            .Right = .Left + tmpVal
            
          '  If ScrollAmount > 0 Then
                'tmpVal = ((ScrollableSpace / ScrollAmount) * pValue)
                
                'ratioMoveable = (UserControl.ScaleHeight / ScrollAmount) / 100


'                    If UserControl.ScaleHeight > 0 Then
'                        .Top = UserControl.ScaleHeight * ((ScrollAmount(False) / UserControl.ScaleHeight) * pValue)
'                    End If

                
          '  End If
            
            'Debug.Print UserControl.ScaleHeight * ((ScrollAmount(False) / UserControl.ScaleHeight ) * pValue)
            
            
            
           ' ratioToScreen = (Screen.Height / UserControl.ScaleHeight)
            
           ' Debug.Print ratioMoveable; ratioToScreen
            

            

            
'            If pProportionalThumb And ScrollAmount(True) <> 0 Then
'                If (ScrollAmount(False) / ScrollAmount(True)) <> 0 Then
'                    fullamt = (1 - (ScrollableSpace / (UserControl.ScaleHeight / (ScrollAmount(False) / ScrollAmount(True)))))
'                    tmpVal = (rScroll.Bottom - rScroll.Top) - ((rScroll.Bottom - rScroll.Top) * fullamt)
'                Else
'                    tmpVal = 0
'                End If
'                'If tmpVal < rCanvas.Right And tmpVal > 0 Then tmpVal = rScroll.Right - rScroll.Left
'
'            Else
'                tmpVal = rCanvas.Right
'            End If
''
'            If pThumbValue <> 0 Then 'draw at thumb locatoin
'                If .Top + pThumbValue > .Bottom - tmpVal Then
'                    .Top = (.Bottom - .Top)
'                ElseIf Not (.Top + pThumbValue < .Top) Then
'                    .Top = (.Top + pThumbValue)
'                End If
'            ElseIf ScrollAmount > 0 Then 'draw at value location
'                 .Top = .Top + (((.Bottom - (rCanvas.Right * 2)) / ScrollAmount) * ((pValue + pMin) * fullamt))
'                 If .Top + pThumbValue > .Bottom - tmpVal Then .Top = (.Bottom - tmpVal)
'            End If
'            .Bottom = .Top + tmpVal

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

Private Function ScrollableSpace(Optional ByVal ExcludeSlider As Boolean = True) As Single
    If IsHorizontal Then
        ScrollableSpace = (UserControl.ScaleWidth - (rCanvas.Bottom * 2))
    ElseIf IsVertical Then
        ScrollableSpace = (UserControl.ScaleHeight - (rCanvas.Right * 2))
    End If
    If ExcludeSlider Then
        If IsHorizontal Then
            ScrollableSpace = ScrollableSpace - (rSlider.Right - rSlider.Left)
        ElseIf IsVertical Then
            ScrollableSpace = ScrollableSpace - (rSlider.Bottom - rSlider.Top)
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
    If Enabled Then

        If pThumbValue <> 0 Then
            Value = ((((pThumbValue / (ScrollableSpace / ScrollAmount))) \ pSmallChange) * pSmallChange)
        Else
            Value = pValue
        End If
        If Value < pMin Then
            Value = pMin
        End If
        If Value > pMax Then
            Value = pMax
        End If
    Else
        Value = pMin
    End If
End Property
Public Property Let Value(ByVal RHS As Long) ' _
Sets the current value of the scroll bar's slider.
Attribute Value.VB_Description = "Sets the current value of the scroll bar's slider."
    If Not Enabled Then Exit Property

    If RHS < pMin Then
        RHS = pMin
    End If
    If RHS > pMax Then
        RHS = pMax
    End If
    If RHS <> pValue Then
                    
        If RHS = pMin Then
            pValue = RHS
            SendScrollBarValue SB_TOP
        ElseIf RHS = pMax Then
            pValue = RHS
            SendScrollBarValue SB_BOTTOM
        ElseIf RHS = pValue - pSmallChange Then
            pValue = RHS
            SendScrollBarValue SB_LINEUP
        ElseIf RHS = pValue - pLargeChange Then
            pValue = RHS
            SendScrollBarValue SB_PAGEUP
        ElseIf RHS = pValue + pSmallChange Then
            pValue = RHS
            SendScrollBarValue SB_LINEDOWN
        ElseIf RHS = pValue + pLargeChange Then
            pValue = RHS
            SendScrollBarValue SB_PAGEDOWN
        Else
            pValue = RHS
        End If
        
        UserControl_Paint
        
        Static lastRaise As Single
        If lastRaise <> pValue And lastRaise <> 0 Then
            RaiseEvent Change
        End If
        lastRaise = pValue
        

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
    Enabled = UserControl.Enabled
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

Private Property Let IControl_hProc(ByVal RHS As Long)
    Me.hProc = RHS
End Property

Private Property Get IControl_hProc() As Long
    IControl_hProc = Me.hProc
End Property

Private Property Get IControl_hWnd() As Long
    IControl_hWnd = Me.hWnd
End Property

Private Sub Timer1_Timer()
    
    Timer1.Interval = keySpeed

    UserControl_MouseMove pPressed, pShift, pEventX, pEventY
    UserControl_Paint
    Timer1.Enabled = (pPressed <> 0)

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
    Timer1.Tag = ""
    Set pBackBuffer = New Backbuffer
    pBackBuffer.hWnd = UserControl.hWnd
    pBackBuffer.Forecolor = vbBlack
    pBackBuffer.BackColor = vbWhite
    Set pBackBuffer.Font = UserControl.Font
    
    Hook Me
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

        If Button = 1 And (X >= 0 And Y >= 0 And X <= UserControl.ScaleWidth And Y <= UserControl.ScaleHeight) Then
       
            pPressed = Button
            pShift = Shift
            pEventX = X
            pEventY = Y
        
            rSlider = GetSliderRect
            'Debug.Print rSlider.Top; rSlider.Bottom; rSlider.Left; rSlider.Right
            
            If PtInRect(rButton1, X, Y) Then
                pHitRegion = ScrollButton1 'top or left most button / smallchange
                If pValue > pMin And pValue - pSmallChange >= pMin Then
                    Value = pValue - pSmallChange
                Else
                    Value = pMin
                End If
            ElseIf PtInRect(rButton2, X, Y) Then
                pHitRegion = ScrollButton2 'bottom or right most button / smallchange
                If pValue < pMax And pValue + pSmallChange <= pMax Then
                    Value = pValue + pSmallChange
                Else
                    Value = pMax
                End If
            ElseIf PtInRect(rSlider, X, Y) Then

                If pHitRegion = SliderButton Then
                    'this is the very end of holding the mouse until silder contacts it
                    If IsHorizontal Then
                        pThumbValue = ((ScrollableSpace / ScrollAmount) * pValue)
                    ElseIf IsVertical Then
                        pThumbValue = ((ScrollableSpace / ScrollAmount) * pValue)
                    End If
                End If
                
                pHitRegion = ScrollableArea

                If pThumbValue = 0 Then
                    pThumbValue = ((ScrollableSpace(True) / ScrollAmount) * pValue)
                    
                    pThumbX = X
                    pThumbY = Y
                End If
                
                    
            ElseIf PtInRect(rScroll, X, Y) Then
                pHitRegion = SliderButton
 
                Dim tmpVal As Single
                If IsHorizontal Then
                    If ScrollAmount > 0 Then
                        If (ScrollableSpace / ScrollAmount) > 0 Then
                            tmpVal = (((X - UserControl.ScaleHeight) / (ScrollableSpace / ScrollAmount)))
                            If ((tmpVal < pValue) And (tmpVal < pMin)) Then tmpVal = pMin
                            If ((tmpVal > pValue) And (tmpVal > pMax)) Then tmpVal = pMax
                            If X < rSlider.Left And Y > rSlider.Top And Y < rSlider.Bottom Then
                                'left side of scrollable space
                                If pValue > tmpVal And pValue - pLargeChange >= tmpVal Then
                                    Value = pValue - pLargeChange
                                Else
                                    Value = tmpVal
                                End If
                            ElseIf X > rSlider.Right And Y > rSlider.Top And Y < rSlider.Bottom Then
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
                            tmpVal = (((Y - UserControl.ScaleWidth) / (ScrollableSpace / ScrollAmount)))
                            If (tmpVal < pValue) And (tmpVal < pMin) Then tmpVal = pMin
                            If (tmpVal > pValue) And (tmpVal > pMax) Then tmpVal = pMax
                            If Y < rSlider.Top And X > rSlider.Left And X < rSlider.Right Then
                                'top of scrollable space
                                If pValue > tmpVal And pValue - pLargeChange >= tmpVal Then
                                    Value = pValue - pLargeChange
                                Else
                                    Value = tmpVal
                                End If
                            ElseIf Y > rSlider.Bottom And X > rSlider.Left And X < rSlider.Right Then
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


        Timer1.Enabled = (pPressed <> 0)

    End If
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub SendScrollBarValue(ByVal SBConst As Integer)
    Dim lword As Long
    LoWord(lword) = SBConst
    If Value <= modBitValue.IntBound Then HiWord(lword) = Value
    If IsVertical Then
        SendMessage hWnd, WM_VSCROLL, lword, ByVal 0&
        SendMessage GetParent(hWnd), WM_VSCROLL, lword, hWnd
    Else
        SendMessage hWnd, WM_HSCROLL, lword, ByVal 0&
        SendMessage GetParent(hWnd), WM_HSCROLL, lword, hWnd
    End If
End Sub

Private Sub StopLongClick()
    If Not pHitRegion = HitRegions.None Then
        'Debug.Print "StopLongClick"
        pHitRegion = HitRegions.None
        pThumbValue = 0
        Timer1.Interval = keySpeed * 10
        Timer1.Enabled = False
        UserControl_Paint
        SendScrollBarValue SB_ENDSCROLL
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Enabled Then
        If pPressed <> 0 Then
            If IsHorizontal Then
                Y = (rCanvas.Bottom - rCanvas.Top) / 2

            ElseIf IsVertical Then
                X = (rCanvas.Right - rCanvas.Left) / 2

            End If
        End If
        
        pPressed = Button

        If pPressed = 0 Then
            StopLongClick
        ElseIf Button = 1 Then


            If Not pHitRegion = HitRegions.None Then
                Dim tmpRct As RECT
                tmpRct = GetSliderRect
                    
                If tmpRct.Left <> rSlider.Left Or tmpRct.Top <> rSlider.Top Or tmpRct.Bottom <> rSlider.Bottom Or tmpRct.Right <> rSlider.Right Then
                    rSlider = tmpRct
                    UserControl_Paint
                    SendScrollBarValue SB_THUMBTRACK
                    
                    RaiseEvent Scroll
                End If
                                

                If pHitRegion = ScrollableArea Then
                    If ScrollAmount > 0 Then
                        If IsHorizontal Then
                            pThumbValue = pThumbValue + (X - pThumbX) '+ ((ScrollableSpace(True) / ScrollAmount) * pValue)
                            pThumbX = X
                        Else
                            pThumbValue = pThumbValue + (Y - pThumbY) '+ ((ScrollableSpace(True) / ScrollAmount) * pValue)
                            pThumbY = Y
                        End If

                    End If

                End If
                
                If pThumbValue <> 0 Then

                    If ScrollAmount > 0 Then
                        Value = (((pThumbValue / (ScrollableSpace / ScrollAmount))) \ pSmallChange) * pSmallChange
                    End If

                End If

           End If
           
           If pHitRegion = SliderButton Then

               If ScrollAmount > 0 Then
                   UserControl_MouseDown Button, Shift, X, Y
                   
                   If IsHorizontal Then
                       If (Not (rSlider.Right < X)) And (Not (rSlider.Left > X)) Then
                           pHitRegion = HitRegions.None
                       End If

                   Else
                       If (Not (rSlider.Bottom < Y)) And (Not (rSlider.Top > Y)) Then
                           pHitRegion = HitRegions.None
                       End If
                   End If

               End If

           ElseIf pHitRegion = ScrollButton1 Or pHitRegion = ScrollButton2 Or pHitRegion = ScrollableArea Then
               UserControl_MouseDown Button, Shift, X, Y
           End If
           
        End If
    End If
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Enabled Then
        pPressed = Button

        If pPressed = 0 Then
            StopLongClick
        ElseIf Button = 1 Then

            If Not pHitRegion = HitRegions.None Then
                If pThumbValue <> 0 Then

                    SendScrollBarValue SB_THUMBPOSITION

                    rSlider = GetSliderRect
                    
                    If ScrollAmount > 0 Then
                        Value = (((pThumbValue / (ScrollableSpace / ScrollAmount))) \ pSmallChange) * pSmallChange
                    End If

                    pThumbValue = 0
                End If
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
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
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

    If AutoRedraw Then
        Refresh
        PaintBuffer
    End If
    RaiseEvent Paint
End Sub
Public Sub Refresh()
    If UserControl.Height > 0 And UserControl.Width > 0 Then
        UpdateRects
        If IsHorizontal Then
            DrawhScrollBar
        ElseIf IsVertical Then
            DrawvScrollBar
        End If
    End If
End Sub
Public Sub Paint() ' _
Redraws the control, this is preformed automatically when AutoRedraw is true.
Attribute Paint.VB_Description = "Redraws the control, this is preformed automatically when AutoRedraw is true."
    
    Refresh
    PaintBuffer
End Sub

Friend Sub PaintBuffer()
    If UserControl.Height > 0 And UserControl.Width > 0 Then
        pBackBuffer.Paint rCanvas.Left, rCanvas.Top, rCanvas.Right, rCanvas.Bottom
    End If

End Sub
Private Sub DrawvScrollBar()

    'draw scrollable area backdrop
    If Enabled Then
        If rSlider.Top - rButton1.Bottom > 0 Then
            pBackBuffer.DrawFrame rCanvas.Left - 2, rButton1.Bottom - 2, rCanvas.Right + 1, rSlider.Top + 1, DFC_BUTTON, DFCS_BUTTONCHECK Or Not DFCS_MONO
        End If
        If rButton2.Top - rSlider.Bottom > 0 Then
            pBackBuffer.DrawFrame rCanvas.Left - 2, rSlider.Bottom - 2, rCanvas.Right + 1, rButton2.Top + 1, DFC_BUTTON, DFCS_BUTTONCHECK Or Not DFCS_MONO
        End If
    Else
        pBackBuffer.DrawFrame rCanvas.Left - 2, rButton1.Bottom - 2, rCanvas.Right + 1, rButton2.Top + 1, DFC_BUTTON, DFCS_BUTTONCHECK Or Not DFCS_MONO
    End If

    If Enabled And ScrollableSpace > 0 And rSlider.Top <> rSlider.Bottom Then
        'draw the slider bar
        If rCanvas.Bottom - rCanvas.Top > rCanvas.Right * 3 - (rCanvas.Right / 2) Then
            pBackBuffer.DrawFrame rSlider.Left, rSlider.Top, rSlider.Right, rSlider.Bottom, DFC_BUTTON, DFCS_BUTTONPUSH
        ElseIf rCanvas.Bottom - rCanvas.Top > rCanvas.Right * 2 Then
            pBackBuffer.DrawFrame rCanvas.Left, rCanvas.Top + rCanvas.Right, rCanvas.Right, rCanvas.Bottom - rCanvas.Right, DFC_BUTTON, DFCS_BUTTONPUSH
        End If
    End If
    
    'draw the arrow buttons
    If rCanvas.Bottom - rCanvas.Top > rCanvas.Right * 2 Then
        pBackBuffer.DrawFrame rButton1.Left, rButton1.Top, rButton1.Right, rButton1.Bottom, DFC_SCROLL, DFCS_SCROLLUP Or IIf(pHitRegion = ScrollButton1 And Enabled, DFCS_PUSHED, 0) Or IIf(Enabled, 0, DFCS_INACTIVE)
        pBackBuffer.DrawFrame rButton2.Left, rButton2.Top, rButton2.Right, rButton2.Bottom, DFC_SCROLL, DFCS_SCROLLDOWN Or IIf(pHitRegion = ScrollButton2 And Enabled, DFCS_PUSHED, 0) Or IIf(Enabled, 0, DFCS_INACTIVE)
    ElseIf rCanvas.Bottom - rCanvas.Top > 0 Then
        pBackBuffer.DrawFrame rButton1.Left, rButton1.Top, rButton1.Right, rButton1.Bottom, DFC_SCROLL, DFCS_SCROLLUP Or IIf(pHitRegion = ScrollButton1 And Enabled, DFCS_PUSHED, 0) Or IIf(Enabled, 0, DFCS_INACTIVE)
        pBackBuffer.DrawFrame rButton2.Left, rButton2.Top, rButton2.Right, rButton2.Bottom, DFC_SCROLL, DFCS_SCROLLDOWN Or IIf(pHitRegion = ScrollButton2 And Enabled, DFCS_PUSHED, 0) Or IIf(Enabled, 0, DFCS_INACTIVE)
    End If
 
End Sub

Private Sub DrawhScrollBar()

    'draw scrollable area backdrop
    If Enabled Then
        If rSlider.Left - rButton1.Right > 0 Then
            pBackBuffer.DrawFrame rButton1.Right - 2, rCanvas.Top - 2, rSlider.Left + 1, rCanvas.Bottom + 1, DFC_BUTTON, DFCS_BUTTONCHECK Or Not DFCS_MONO
        End If
        If rButton2.Left - rSlider.Right > 0 Then
            pBackBuffer.DrawFrame rSlider.Right - 2, rCanvas.Top - 2, rButton2.Left + 1, rCanvas.Bottom + 1, DFC_BUTTON, DFCS_BUTTONCHECK Or Not DFCS_MONO
        End If
    Else
        pBackBuffer.DrawFrame rButton1.Right - 2, rCanvas.Top - 2, rButton2.Left + 1, rCanvas.Bottom + 1, DFC_BUTTON, DFCS_BUTTONCHECK Or Not DFCS_MONO
    End If

    If Enabled And ScrollableSpace > 0 And rSlider.Left <> rSlider.Right Then
        'draw the slider bar
        If rCanvas.Right - rCanvas.Left > rCanvas.Bottom * 3 - (rCanvas.Bottom / 2) Then
            pBackBuffer.DrawFrame rSlider.Left, rSlider.Top, rSlider.Right, rSlider.Bottom, DFC_BUTTON, DFCS_BUTTONPUSH
        ElseIf rCanvas.Right - rCanvas.Left > rCanvas.Bottom * 2 Then
            pBackBuffer.DrawFrame rCanvas.Left + rCanvas.Bottom, rCanvas.Top, rCanvas.Right - rCanvas.Bottom, rCanvas.Bottom, DFC_BUTTON, DFCS_BUTTONPUSH
        End If
    End If
    
    'draw the arrow buttons
    If rCanvas.Right - rCanvas.Left > rCanvas.Bottom * 2 Then
        pBackBuffer.DrawFrame rButton1.Left, rButton1.Top, rButton1.Right, rButton1.Bottom, DFC_SCROLL, DFCS_SCROLLLEFT Or IIf(pHitRegion = ScrollButton1 And Enabled, DFCS_PUSHED, 0) Or IIf(Enabled, 0, DFCS_INACTIVE)
        pBackBuffer.DrawFrame rButton2.Left, rButton2.Top, rButton2.Right, rButton2.Bottom, DFC_SCROLL, DFCS_SCROLLRIGHT Or IIf(pHitRegion = ScrollButton2 And Enabled, DFCS_PUSHED, 0) Or IIf(Enabled, 0, DFCS_INACTIVE)
    ElseIf rCanvas.Right - rCanvas.Left > 0 Then
        pBackBuffer.DrawFrame rButton1.Left, rButton1.Top, rButton1.Right, rButton1.Bottom, DFC_SCROLL, DFCS_SCROLLLEFT Or IIf(pHitRegion = ScrollButton1 And Enabled, DFCS_PUSHED, 0) Or IIf(Enabled, 0, DFCS_INACTIVE)
        pBackBuffer.DrawFrame rButton2.Left, rButton2.Top, rButton2.Right, rButton2.Bottom, DFC_SCROLL, DFCS_SCROLLRIGHT Or IIf(pHitRegion = ScrollButton2 And Enabled, DFCS_PUSHED, 0) Or IIf(Enabled, 0, DFCS_INACTIVE)
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

    ProportionalThumb = PropBag.ReadProperty("ProportionalThumb", True)

    Paint
End Sub

Private Sub UserControl_InitProperties()
    Orientation = vbOrientation.OrientationAuto
    Enabled = True
    Max = 100
    SmallChange = 1
    LargeChange = 4
    AutoRedraw = True
    ProportionalThumb = False

    Paint
End Sub

Private Sub UserControl_Terminate()
    Unhook Me
    
    Set pBackBuffer = Nothing
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
    
    PropBag.WriteProperty "ProportionalThumb", ProportionalThumb
    
End Sub
