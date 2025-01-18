Attribute VB_Name = "modCmds"


#Const modCmds = -1
Option Explicit
'TOP DOWN

Option Compare Binary


'Public Type InfoType
'    name As String
'    Address1 As String
'    Address2 As String
'    City As String
'    State As String
'    ZipCode As String
'End Type
'
'Public Company As InfoType
'Public Artist As InfoType


Private DInput As DirectInput8
Private DIKeyBoardDevice As DirectInputDevice8
Private DIKEYBOARDSTATE As DIKEYBOARDSTATE

Private DIMouseDevice As DirectInputDevice8
Private DIMOUSESTATE As DIMOUSESTATE

Private TogglePress1 As Long
Private TogglePress2 As Long
Private TogglePress3 As Long
'Private TogglePress4 As Long

Private Type KeyState
    VKState As Integer
    VKToggle As Boolean
    VKPressed As Boolean
    VKLatency As Single
End Type

Private IdleInput As Single

Private KeyState(255) As KeyState
Private KeyChars(255) As String

Private lX As Integer
Private lY As Integer
Private lZ As Integer
Private norepeat As String

Private SkyPlaq(0 To 35) As TVERTEX2
Private SkySkin(0 To 5) As Direct3DTexture8
Private SkyVBuf As Direct3DVertexBuffer8
Private SkyCRot As Single

Public Property Get TextHeight() As Single
    TextHeight = frmMain.TextHeight("A")
End Property

Public Property Get TextSpace() As Single
    TextSpace = 2
End Property

Private Function Toggled(ByVal vkCode As Long) As Boolean
    Toggled = KeyState(vkCode).VKToggle
End Function
Private Function Pressed(ByVal vkCode As Long) As Boolean
    Pressed = KeyState(vkCode).VKPressed
End Function
Private Sub DrawEdgeCheck()

End Sub

Public Sub InputScene()

    On Error GoTo pausing
    NotFocused = (GetActiveWindow <> frmStudio.hwnd) And (GetActiveWindow <> frmMain.hwnd)
    If Not NotFocused Then
        'Static lastfocus As Boolean
        If (GetActiveWindow <> frmStudio.hwnd) And (GetActiveWindow <> frmMain.hwnd) Then
            DoNotFocused
        End If
        
    End If
    
    If Not NotFocused Then
        TrapMouse = 0
        DIKeyBoardDevice.GetDeviceStateKeyboard DIKEYBOARDSTATE
    
        ConsoleInput DIKEYBOARDSTATE
    
        If DIKEYBOARDSTATE.Key(DIK_ESCAPE) Then
            If (Not TogglePress1 = DIK_ESCAPE) Then
                TogglePress1 = DIK_ESCAPE
    
                GotoCenter
    
            End If
        ElseIf (DIKEYBOARDSTATE.Key(DIK_LCONTROL) Or DIKEYBOARDSTATE.Key(DIK_RCONTROL)) And DIKEYBOARDSTATE.Key(DIK_Z) Then
    
            If (Not TogglePress1 = DIK_Z) Then
                TogglePress1 = DIK_Z
    
                frmStudio.mnuUndo_Click
    
            End If
        ElseIf (DIKEYBOARDSTATE.Key(DIK_LCONTROL) Or DIKEYBOARDSTATE.Key(DIK_RCONTROL)) And DIKEYBOARDSTATE.Key(DIK_Y) Then
    
            If (Not TogglePress1 = DIK_Y) Then
                TogglePress1 = DIK_Y
    
                frmStudio.mnuRedo_Click
    
            End If
    
        ElseIf DIKEYBOARDSTATE.Key(DIK_TAB) Then
            If (Not TogglePress1 = DIK_TAB) Then
                TogglePress1 = DIK_TAB
    
                Select Case GetActiveWindow
                    Case frmMain.hwnd, frmMain.Picture1.hwnd, frmStudio.Designer.hwnd
                        frmStudio.TabStrip1.SetFocus
                End Select
    
            End If
    
        ElseIf DIKEYBOARDSTATE.Key(DIK_BACKSPACE) Then
            If (Not TogglePress1 = DIK_BACKSPACE) Then
                TogglePress1 = DIK_BACKSPACE
                
                If frmStudio.mnuAdvanced.Checked Then
                    If frmStudio.RemovalTool.Value Then
                        frmStudio.Precision_Click
                    Else
                        frmStudio.RemovalTool_Click
                    End If
                Else
                    If frmStudio.RemovalTool.Value Then
                        frmStudio.Precision_Click
                    Else
                        frmStudio.RemovalTool_Click
                    End If
                End If
            End If
        ElseIf DIKEYBOARDSTATE.Key(DIK_SPACE) Then
            If (Not TogglePress1 = DIK_SPACE) Then
                TogglePress1 = DIK_SPACE
    
                If frmStudio.RemovalTool.Value Then
                    frmStudio.Precision_Click
                ElseIf frmStudio.Precision.Value Then
                    frmStudio.LinkedLines_Click
                End If
            End If
        ElseIf DIKEYBOARDSTATE.Key(DIK_0) Then
            If (Not TogglePress1 = DIK_0) Then
                TogglePress1 = DIK_0
                frmStudio.RemovalTool_Click
                frmStudio.VScroll2.Value = -0
            End If
        ElseIf DIKEYBOARDSTATE.Key(DIK_1) Then 'KeyCode: 191/
            If (Not TogglePress1 = DIK_1) Then
                TogglePress1 = DIK_1
                frmStudio.Precision_Click
                frmStudio.VScroll2.Value = -1
            End If
        ElseIf DIKEYBOARDSTATE.Key(DIK_2) Then 'KeyCode: 220\
            If (Not TogglePress1 = DIK_2) Then
                TogglePress1 = DIK_2
                frmStudio.Precision_Click
                frmStudio.VScroll2.Value = -2
            End If
        ElseIf DIKEYBOARDSTATE.Key(DIK_3) Then 'KeyCode: 189-
            If (Not TogglePress1 = DIK_3) Then
                TogglePress1 = DIK_3
                frmStudio.Precision_Click
                frmStudio.VScroll2.Value = -3
            End If
        ElseIf DIKEYBOARDSTATE.Key(DIK_4) Then 'KeyCode: 187=
            If (Not TogglePress1 = DIK_4) Then
                TogglePress1 = DIK_4
                frmStudio.Precision_Click
                frmStudio.VScroll2.Value = -4
            End If
        ElseIf DIKEYBOARDSTATE.Key(DIK_5) Then 'KeyCode: 219[
            If (Not TogglePress1 = DIK_5) Then
                TogglePress1 = DIK_5
    
                frmStudio.Precision_Click
                frmStudio.VScroll2.Value = -5
            End If
    '    ElseIf DIKEYBOARDSTATE.Key(DIK_RBRACKET) Then 'KeyCode: 221]
    '        If (Not TogglePress1 = DIK_RBRACKET) Then
    '            TogglePress1 = DIK_RBRACKET
    '
    '            frmStudio.Precision_Click
    '            frmStudio.VScroll2.Value = -4
    '        End If
        ElseIf DIKEYBOARDSTATE.Key(DIK_BACKSPACE) Then 'backspace
            If (Not TogglePress1 = DIK_BACKSPACE) Then
                TogglePress1 = DIK_BACKSPACE
    
                frmStudio.RemovalTool_Click
                frmStudio.VScroll2.Value = 0
    
            End If
        ElseIf Not (TogglePress1 = 0) Then
            TogglePress1 = 0
        End If
    
        If frmStudio.Designer.Visible Then
    
            Dim rec  As RECT
            GetWindowRect frmStudio.Designer.hwnd, rec
    
            DIMouseDevice.GetDeviceStateMouse DIMOUSESTATE
    
            Dim mX As Integer
            Dim mY As Integer
            Dim mZ As Integer
    
            Dim mloc As POINTAPI
            GetCursorPos mloc
    
            mX = DIMOUSESTATE.lX
            mY = DIMOUSESTATE.lY
            mZ = DIMOUSESTATE.lZ
    
            MouseLook 0, 0, mZ
    
    
           ' If (GetActiveWindow = frmMain.hWNd) Then
            If mloc.X > rec.Left And mloc.X < rec.Right And mloc.Y > rec.Top And mloc.Y < rec.Bottom Then
    
                If (DIKEYBOARDSTATE.Key(DIK_LCONTROL) Or DIKEYBOARDSTATE.Key(DIK_RCONTROL)) Then
                    TrapMouse = 1
                    If DIMOUSESTATE.Buttons(1) Then
    
                        MouseLook mX, mY, mZ
    
    '                    If Not (frmMain.MousePointer = 99) Then
    '                        frmMain.MousePointer = 99
    '                        frmMain.MouseIcon = LoadPicture(AppPath & "Base\mouse.cur")
    '                    End If
    
                        If (mY <> lY) Then
                            Player.Location.Y = Player.Location.Y + (Cos(D720) * Player.MoveSpeed)
                        End If
    
                        If (mY <> lY) Then
                            Player.Location.Y = Player.Location.Y - (Cos(D720) * Player.MoveSpeed)
                        End If
    
                        If (mX <> lX) Then
                            Player.Location.X = Player.Location.X + (Sin((D720) - D180) * Player.MoveSpeed)
                        End If
    
                        If (mX <> lX) Then
                            Player.Location.X = Player.Location.X + (Sin((D720) + D180) * Player.MoveSpeed)
                        End If
    
                        GetWindowRect frmStudio.Designer.hwnd, rec
    
                        SetCursorPos (rec.Right + ((rec.Left - rec.Right) / 2)) - (lX - mX), (rec.Top + ((rec.Bottom - rec.Top) / 2)) - (lY - mY)
    
                    End If
                ElseIf DIMOUSESTATE.Buttons(0) Then
                    Dim undoBuff As String
                    
    
                    
    
                   ' Debug.Print norepeat
        
                    If BlockSetX <= UBound(ProjGrid, 1) And BlockSetX > 0 And BlockSetY <= UBound(ProjGrid, 2) And BlockSetY > 0 Then
                        Static bX As Long
                        Static bY As Long
                        Static bN As Long
    
                        If frmStudio.RemovalTool.Value Then
    
                            If BlockSetX & "," & BlockSetY <> NextArg(norepeat, "|") Then norepeat = ""
                            
                            If frmStudio.VScroll2.Value = 0 Then
    
        
                                If frmStudio.mnuAdvanced.Checked Then
                                    If DetailSetY >= 30 Then
                                        If CheckStitch(StitchBit.TopEdgeThin) Then
                                            If norepeat = "" Then
                                                norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                                CheckStitch(StitchBit.TopEdgeThin) = False
                                                
                                            End If
                                        ElseIf CheckStitch(StitchBit.TopEdgeThick) And (frmStudio.DoubleThick.Value = 1) Then
                                            If norepeat = "" Then
                                                norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                                CheckStitch(StitchBit.TopEdgeThick) = False
                                            End If
                                        End If
                                    End If
                                    If DetailSetY < 30 Then
                                        If CheckStitch(StitchBit.BottomEdgeThin) Then
                                            If norepeat = "" Then
                                                norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                                CheckStitch(StitchBit.BottomEdgeThin) = False
                                            End If
                                        ElseIf CheckStitch(StitchBit.BottomEdgeThick) And (frmStudio.DoubleThick.Value = 1) Then
                                            If norepeat = "" Then
                                                norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                                CheckStitch(StitchBit.BottomEdgeThick) = False
                                            End If
                                        End If
                                    End If
                                    If DetailSetX >= 30 Then
                                        If CheckStitch(StitchBit.RightEdgeThin) Then
                                            If norepeat = "" Then
                                                norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                                CheckStitch(StitchBit.RightEdgeThin) = False
                                            End If
                                        ElseIf CheckStitch(StitchBit.RightEdgeThick) And (frmStudio.DoubleThick.Value = 1) Then
                                            If norepeat = "" Then
                                                norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                                CheckStitch(StitchBit.RightEdgeThick) = False
                                            End If
                                        End If
                                    End If
                                    If DetailSetX < 30 Then
                                        If CheckStitch(StitchBit.LeftEdgeThin) Then
                                            If norepeat = "" Then
                                                norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                                CheckStitch(StitchBit.LeftEdgeThin) = False
                                            End If
                                        ElseIf CheckStitch(StitchBit.LeftEdgeThick) And (frmStudio.DoubleThick.Value = 1) Then
                                            If norepeat = "" Then
                                                norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                                CheckStitch(StitchBit.LeftEdgeThick) = False
                                            End If
                                        End If
                                    End If
                                    
                                    If ((DetailSetX >= 30 And DetailSetY < 30) Or (DetailSetX < 30 And DetailSetY >= 30)) Then
            
                                        If CheckStitch(StitchBit.ForwardSlashThin) Then
                                            If norepeat = "" Then
                                                norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                                CheckStitch(StitchBit.ForwardSlashThin) = False
                                            End If
                                        ElseIf CheckStitch(StitchBit.ForwardSlashThick) And (frmStudio.DoubleThick.Value = 1) Then
                                            If norepeat = "" Then
                                                norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                                CheckStitch(StitchBit.ForwardSlashThick) = False
                                            End If
                                        End If
                                    End If
            
                                    If ((DetailSetX < 30 And DetailSetY < 30) Or (DetailSetX >= 30 And DetailSetY >= 30)) Then
            
                                        If CheckStitch(StitchBit.BackSlashThin) Then
                                            If norepeat = "" Then
                                                norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                                CheckStitch(StitchBit.BackSlashThin) = False
                                            End If
                                        ElseIf CheckStitch(StitchBit.BackSlashThick) And (frmStudio.DoubleThick.Value = 1) Then
                                            If norepeat = "" Then
                                                norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                                CheckStitch(StitchBit.BackSlashThick) = False
                                            End If
                                        End If
                                    End If
                                
                                Else
                                
                                    If CheckStitch(StitchBit.TopEdgeThin) Then
                                        If norepeat = "" Then
                                            norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                            CheckStitch(StitchBit.TopEdgeThin) = False
                                        End If
                                    End If
        
                                    If CheckStitch(StitchBit.RightEdgeThin) Then
                                        If norepeat = "" Then
                                            norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                            CheckStitch(StitchBit.RightEdgeThin) = False
                                        End If
        
                                    End If
        
                                    If CheckStitch(StitchBit.ForwardSlashThin) Then
                                        If norepeat = "" Then
                                            norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                            CheckStitch(StitchBit.ForwardSlashThin) = False
                                        End If
                                    End If
        
                                    If CheckStitch(StitchBit.BackSlashThin) Then
                                        If norepeat = "" Then
                                            norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                            CheckStitch(StitchBit.BackSlashThin) = False
                                        End If
                                    End If
                                
                                End If
                            End If
                            
                        ElseIf frmStudio.Precision.Value Then
    
                            If BlockSetX & "," & BlockSetY <> NextArg(norepeat, "|") Then norepeat = ""
    
                            If frmStudio.mnuAdvanced.Checked Then
                                If DetailSetY >= 30 And frmStudio.VScroll2.Value = -2 Then
                                    If (Not CheckStitch(StitchBit.TopEdgeThin)) Then
                                        If norepeat = "" Then
                                            norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                            CheckStitch(StitchBit.TopEdgeThin) = True
                                            If frmStudio.DoubleThick.Value = 1 Then
                                                CheckStitch(StitchBit.TopEdgeThick) = True
                                            End If
                                        End If
                                    ElseIf (Not CheckStitch(StitchBit.TopEdgeThick)) And (frmStudio.DoubleThick.Value = 1) Then
                                        If norepeat = "" Then
                                            norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                            CheckStitch(StitchBit.TopEdgeThick) = True
                                        End If
                                    End If
                                End If
                                If DetailSetY < 30 And frmStudio.VScroll2.Value = -2 Then
                                    If (Not CheckStitch(StitchBit.BottomEdgeThin)) Then
                                        If norepeat = "" Then
                                            norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                            CheckStitch(StitchBit.BottomEdgeThin) = True
                                            If frmStudio.DoubleThick.Value = 1 Then
                                                CheckStitch(StitchBit.BottomEdgeThick) = True
                                            End If
                                        End If
                                    ElseIf (Not CheckStitch(StitchBit.BottomEdgeThick)) And (frmStudio.DoubleThick.Value = 1) Then
                                        If norepeat = "" Then
                                            norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                            CheckStitch(StitchBit.BottomEdgeThick) = True
                                        End If
                                    End If
                                End If
                                If DetailSetX >= 30 And frmStudio.VScroll2.Value = -1 Then
                                    If (Not CheckStitch(StitchBit.RightEdgeThin)) Then
                                        If norepeat = "" Then
                                            norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                            CheckStitch(StitchBit.RightEdgeThin) = True
                                            If frmStudio.DoubleThick.Value = 1 Then
                                                CheckStitch(StitchBit.RightEdgeThick) = True
                                            End If
                                        End If
                                    ElseIf (Not CheckStitch(StitchBit.RightEdgeThick)) And (frmStudio.DoubleThick.Value = 1) Then
                                        If norepeat = "" Then
                                            norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                            CheckStitch(StitchBit.RightEdgeThick) = True
                                        End If
                                    End If
                                End If
                                If DetailSetX < 30 And frmStudio.VScroll2.Value = -1 Then
                                    If (Not CheckStitch(StitchBit.LeftEdgeThin)) Then
                                        If norepeat = "" Then
                                            norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                            CheckStitch(StitchBit.LeftEdgeThin) = True
                                            If frmStudio.DoubleThick.Value = 1 Then
                                                CheckStitch(StitchBit.LeftEdgeThick) = True
                                            End If
                                        End If
                                    ElseIf (Not CheckStitch(StitchBit.LeftEdgeThick)) And (frmStudio.DoubleThick.Value = 1) Then
                                        If norepeat = "" Then
                                            norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                            CheckStitch(StitchBit.LeftEdgeThick) = True
                                        End If
                                    End If
                                End If
                                
                                If ((DetailSetX >= 30 And DetailSetY < 30) Or (DetailSetX < 30 And DetailSetY >= 30)) And frmStudio.VScroll2.Value = -3 Then
        
                                    If (Not CheckStitch(StitchBit.ForwardSlashThin)) Then
                                        If norepeat = "" Then
                                            norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                            CheckStitch(StitchBit.ForwardSlashThin) = True
                                            If frmStudio.DoubleThick.Value = 1 Then
                                                CheckStitch(StitchBit.ForwardSlashThick) = True
                                            End If
                                        End If
                                    ElseIf (Not CheckStitch(StitchBit.ForwardSlashThick)) And (frmStudio.DoubleThick.Value = 1) Then
                                        If norepeat = "" Then
                                            norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                            CheckStitch(StitchBit.ForwardSlashThick) = True
                                        End If
                                    End If
                                End If
        
                                If ((DetailSetX < 30 And DetailSetY < 30) Or (DetailSetX >= 30 And DetailSetY >= 30)) And frmStudio.VScroll2.Value = -4 Then
        
                                    If (Not CheckStitch(StitchBit.BackSlashThin)) Then
                                        If norepeat = "" Then
                                            norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                            CheckStitch(StitchBit.BackSlashThin) = True
                                            If frmStudio.DoubleThick.Value = 1 Then
                                                CheckStitch(StitchBit.BackSlashThick) = True
                                            End If
                                        End If
                                    ElseIf (Not CheckStitch(StitchBit.BackSlashThick)) And (frmStudio.DoubleThick.Value = 1) Then
                                        If norepeat = "" Then
                                            norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                            CheckStitch(StitchBit.BackSlashThick) = True
                                        End If
                                    End If
                                End If
                      
                      
                                If frmStudio.VScroll2.Value = -5 Then
        
                                    If (Not CheckStitch(StitchBit.BackSlashThin)) And (Not CheckStitch(StitchBit.ForwardSlashThin)) Then
                                        If norepeat = "" Then
                                            norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                            CheckStitch(StitchBit.BackSlashThin) = True
                                            CheckStitch(StitchBit.ForwardSlashThin) = True
                                            If frmStudio.DoubleThick.Value = 1 Then
                                                CheckStitch(StitchBit.BackSlashThick) = True
                                                CheckStitch(StitchBit.ForwardSlashThick) = True
                                            End If
                                        End If
                                    ElseIf (Not CheckStitch(StitchBit.BackSlashThick)) And (Not CheckStitch(StitchBit.ForwardSlashThick)) And (frmStudio.DoubleThick.Value = 1) Then
                                        If norepeat = "" Then
                                            norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                            CheckStitch(StitchBit.BackSlashThick) = True
                                            CheckStitch(StitchBit.ForwardSlashThick) = True
                                        End If
                                    End If
                                End If
    
                            Else
                                If frmStudio.VScroll2.Value = -2 Then
                                    If (Not CheckStitch(StitchBit.TopEdgeThin)) Then
                                        If norepeat = "" Then
                                            norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                            CheckStitch(StitchBit.TopEdgeThin) = True
                                        End If
                                    End If
                                ElseIf frmStudio.VScroll2.Value = -1 Then
                                    If (Not CheckStitch(StitchBit.RightEdgeThin)) Then
                                        If norepeat = "" Then
                                            norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                            CheckStitch(StitchBit.RightEdgeThin) = True
                                        End If
                                    End If
                                ElseIf frmStudio.VScroll2.Value = -3 Then
        
                                    If (Not CheckStitch(StitchBit.ForwardSlashThin)) Then
                                        If norepeat = "" Then
                                            norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                            CheckStitch(StitchBit.ForwardSlashThin) = True
                                        End If
                                    End If
                                ElseIf frmStudio.VScroll2.Value = -4 Then
        
                                    If (Not CheckStitch(StitchBit.BackSlashThin)) Then
                                        If norepeat = "" Then
                                            norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                            CheckStitch(StitchBit.BackSlashThin) = True
                                        End If
                                    End If
                                ElseIf frmStudio.VScroll2.Value = -5 Then
        
                                    If (Not CheckStitch(StitchBit.BackSlashThin)) And (Not CheckStitch(StitchBit.ForwardSlashThin)) Then
                                        If norepeat = "" Then
                                            norepeat = BlockSetX & "," & BlockSetY & "|" & DetailSetX & "," & DetailSetY
                                            CheckStitch(StitchBit.BackSlashThin) = True
                                            CheckStitch(StitchBit.ForwardSlashThin) = True
                                        End If
                                    End If
                                End If
    
                            End If
                        End If
        
                    End If
                Else
                    norepeat = ""
                End If
            End If
        Else
            GetWindowRect frmStudio.Gallery1.hwnd, rec
    
            If mX > rec.Left And mX < rec.Right And mY > rec.Top And mY < rec.Bottom Then
                If mZ > 0 Then
                    frmStudio.Gallery1.ScrollUp
                    frmStudio.Gallery1.ScrollUp
                ElseIf mZ < 0 Then
                    frmStudio.Gallery1.ScrollDown
                    frmStudio.Gallery1.ScrollDown
                End If
            End If
    
        End If
    
        lX = mX
        lY = mX
        lZ = mZ
    End If

    DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
    DDevice.SetVertexShader FVF_RENDER

    D3DXMatrixIdentity matWorld
    DDevice.SetTransform D3DTS_WORLD, matWorld

    Dim fogSTate As Boolean
    fogSTate = DDevice.GetRenderState(D3DRS_FOGENABLE)
    If fogSTate Then DDevice.SetRenderState D3DRS_FOGENABLE, False
    DDevice.SetRenderState D3DRS_LIGHTING, 1
    DDevice.SetRenderState D3DRS_ZENABLE, False

    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, False

    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT

    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_POINT
    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_POINT

    Dim matProj As D3DMATRIX
    Dim matView As D3DMATRIX, matViewSave As D3DMATRIX
    DDevice.GetTransform D3DTS_VIEW, matViewSave
    matView = matViewSave
    matView.m41 = 0: matView.m42 = 0: matView.m43 = 0

    DDevice.SetTransform D3DTS_VIEW, matView

    D3DXMatrixPerspectiveFovLH matProj, FOV, ASPECT, NEAR, FAR
    DDevice.SetTransform D3DTS_PROJECTION, matProj

    DDevice.SetTransform D3DTS_WORLD, matWorld

    DDevice.SetStreamSource 0, SkyVBuf, Len(SkyPlaq(0))

    DDevice.SetTexture 1, Nothing

    DDevice.SetTexture 0, SkySkin(0)
    DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 30, 2
    DDevice.SetTexture 0, SkySkin(1)
    DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
    DDevice.SetTexture 0, SkySkin(2)
    DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 6, 2
    DDevice.SetTexture 0, SkySkin(3)
    DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 12, 2
    DDevice.SetTexture 0, SkySkin(4)
    DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 24, 2
    DDevice.SetTexture 0, SkySkin(5)
    DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 18, 2

    D3DXMatrixPerspectiveFovLH matProj, FOVY, ASPECT, NEAR, FAR
    DDevice.SetTransform D3DTS_PROJECTION, matProj

    DDevice.SetTransform D3DTS_VIEW, matViewSave

    DDevice.SetRenderState D3DRS_ZENABLE, 1
    DDevice.SetRenderState D3DRS_LIGHTING, 1
    If fogSTate Then DDevice.SetRenderState D3DRS_FOGENABLE, 1

    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC

    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC

    DDevice.SetRenderState D3DRS_FOGCOLOR, DDevice.GetRenderState(D3DRS_AMBIENT)

    DDevice.SetRenderState D3DRS_AMBIENT, D3DColorARGB(255, 255, 255, 255)


    Exit Sub
pausing:
    Err.Clear
    DoNotFocused
End Sub


Public Property Get CheckStitch(ByVal sBit As StitchBit, Optional ByVal bSetX As Long, Optional ByVal bSetY As Long, Optional ByVal Color As Long = -1) As Boolean
    Dim i As Byte
    If bSetX = 0 Then bSetX = BlockSetX
    If bSetY = 0 Then bSetY = BlockSetY
    If ProjGrid(bSetX, bSetY).count > 0 Then
        For i = 1 To ProjGrid(bSetX, bSetY).count
            If BitLong(ProjGrid(bSetX, bSetY).Details(i).Stitch, sBit) Then
                CheckStitch = True
                Exit Property
            End If
        Next
    End If
End Property

Public Property Let CheckStitch(ByVal sBit As StitchBit, Optional ByVal bSetX As Long, Optional ByVal bSetY As Long, Optional ByVal Color As Long = -1, ByVal RHS As Boolean)
    Dim i As Byte
    Dim N As Byte
    Dim isUndo As Boolean
    
    If bSetX = 0 And bSetY = 0 And Color = -1 Then
        bSetX = BlockSetX
        bSetY = BlockSetY
        frmStudio.UndoAction = frmStudio.UndoAction & "C" & frmStudio.Gallery1.BackgroundColors(frmStudio.Gallery1.ListIndex) & "," & CLng(sBit) & "," & bSetX & "," & bSetY & "," & RHS & vbCrLf
        frmStudio.UpdateTimer.Enabled = False
        Color = frmStudio.Gallery1.BackgroundColors(frmStudio.Gallery1.ListIndex)
    Else
        isUndo = True
    End If
    
    If ProjGrid(bSetX, bSetY).count > 0 And (Not RHS) Then
        For i = 1 To ProjGrid(bSetX, bSetY).count
            If BitLong(ProjGrid(bSetX, bSetY).Details(i).Stitch, sBit) Then
                BitLong(ProjGrid(bSetX, bSetY).Details(i).Stitch, sBit) = False
                If ProjGrid(bSetX, bSetY).Details(i).Stitch = 0 Then
                    If i < ProjGrid(bSetX, bSetY).count - 2 Then
                        For N = i To ProjGrid(bSetX, bSetY).count - 2
                            ProjGrid(bSetX, bSetY).Details(N) = ProjGrid(bSetX, bSetY).Details(N + 1)
                        Next
                    End If
                    If ProjGrid(bSetX, bSetY).count > 1 Then
                        ReDim Preserve ProjGrid(bSetX, bSetY).Details(1 To ProjGrid(bSetX, bSetY).count - 1) As ItemDetail
                    Else
                        Erase ProjGrid(bSetX, bSetY).Details
                    End If
                    ProjGrid(bSetX, bSetY).count = ProjGrid(bSetX, bSetY).count - 1
                End If
                Exit For
            End If
        Next

    End If
    If frmStudio.CreateColor(Color, True) Then frmStudio.UpdateGallery
    
    If RHS Then
        N = 0
        If ProjGrid(bSetX, bSetY).count > 0 Then
            For i = 1 To ProjGrid(bSetX, bSetY).count
                If ProjGrid(bSetX, bSetY).Details(i).Color = Color Then
                    N = i
                    Exit For
                End If
            Next
        End If
        If N = 0 Then
            ReDim Preserve ProjGrid(bSetX, bSetY).Details(1 To ProjGrid(bSetX, bSetY).count + 1) As ItemDetail
            ProjGrid(bSetX, bSetY).count = ProjGrid(bSetX, bSetY).count + 1
            N = ProjGrid(bSetX, bSetY).count
        End If
        ProjGrid(bSetX, bSetY).Details(N).Color = Color
        BitLong(ProjGrid(bSetX, bSetY).Details(N).Stitch, sBit) = True
        modProj.Dirty = True
    End If
    
    If Not isUndo Then
        frmStudio.UpdateTimer.Enabled = True
        frmStudio.UndoEnables
    End If
End Property

Private Function OneBlockAway(ByVal bx1 As Long, ByVal by1 As Long, ByVal bx2 As Long, ByVal by2 As Long) As Boolean
    OneBlockAway = (((bx1 - bx2 <= 1) Or (bx1 - bx2 >= -1)) And ((by1 - by2 <= 1) Or (by1 - by2 >= -1)))
End Function
Private Function TwoBlockAway(ByVal bx1 As Long, ByVal by1 As Long, ByVal bx2 As Long, ByVal by2 As Long) As Boolean
    TwoBlockAway = (((bx1 - bx2 <= 2) Or (bx1 - bx2 >= -2)) And ((by1 - by2 <= 2) Or (by1 - by2 >= -2)))
End Function
Private Function BlockDifference(ByVal bx1 As Long, ByVal by1 As Long, ByVal bx2 As Long, ByVal by2 As Long) As String
    BlockDifference = Trim((bx1 - bx2)) & "," & Trim((by1 - by2))
End Function
Private Sub MouseLook(ByVal X As Integer, ByVal Y As Integer, ByVal Z As Integer)

    Dim cnt As Long
    If Z < 0 Then
        For cnt = (Z * MouseSensitivity) To 0
            Player.CameraZoom = Player.CameraZoom + 2.15
        Next
    ElseIf Z > 0 Then
        For cnt = 0 To (Z * MouseSensitivity)
            Player.CameraZoom = Player.CameraZoom - 2.15
        Next
    End If
    
    If Player.CameraZoom > MouseZoomOutMax Then Player.CameraZoom = MouseZoomOutMax
    If Player.CameraZoom < MouseZoomInMax Then Player.CameraZoom = MouseZoomInMax

    If X < 0 Then
        For cnt = (X * MouseSensitivity) To 0
            Player.CameraAngle = Player.CameraAngle - -0.0015
        Next
    ElseIf X > 0 Then
        For cnt = 0 To (X * MouseSensitivity)
            Player.CameraAngle = Player.CameraAngle - 0.0015
        Next
    End If
    
    If Player.CameraAngle > (PI * 2) Then Player.CameraAngle = Player.CameraAngle - (PI * 2)
    If Player.CameraAngle < -(PI * 2) Then Player.CameraAngle = Player.CameraAngle + (PI * 2)
    
    If Y < 0 Then
        For cnt = (Y * MouseSensitivity) To 0
            Player.CameraPitch = Player.CameraPitch - -0.0015
        Next
    ElseIf Y > 0 Then
        For cnt = 0 To (Y * MouseSensitivity)
            Player.CameraPitch = Player.CameraPitch - 0.0015
        Next
    End If
    
    If Player.CameraPitch < -1.5 Then Player.CameraPitch = -1.5
    If Player.CameraPitch > 1.5 Then Player.CameraPitch = 1.5
    
End Sub


Public Sub ConsoleInput(ByRef kState As DIKEYBOARDSTATE)

    Dim cnt As Integer
    Dim char As String
    
    For cnt = 0 To 255
        
        If (Not (KeyState(cnt).VKState = kState.Key(cnt))) Then
            KeyState(cnt).VKState = kState.Key(cnt)
            If (KeyState(cnt).VKLatency = 0) Then
                KeyState(cnt).VKPressed = Not KeyState(cnt).VKPressed
                KeyState(cnt).VKLatency = Timer
            Else
                KeyState(cnt).VKPressed = False
            End If
            If KeyState(cnt).VKPressed = False Then
                KeyState(cnt).VKToggle = Not KeyState(cnt).VKToggle
            ElseIf KeyState(cnt).VKToggle Then
                KeyState(cnt).VKToggle = KeyState(cnt).VKToggle Xor (Not KeyState(cnt).VKPressed)
            End If
            If Not KeyState(cnt).VKPressed Then KeyState(cnt).VKLatency = 0
        Else
            If (KeyState(cnt).VKLatency <> 0) Then
                If ((Timer - -KeyState(cnt).VKLatency) > 0.1) And (KeyState(cnt).VKLatency < 0) Then
                    KeyState(cnt).VKPressed = True
                    KeyState(cnt).VKLatency = KeyState(cnt).VKLatency + -0.1
                    IdleInput = Timer
                ElseIf ((Timer - KeyState(cnt).VKLatency) > 0.6) And (KeyState(cnt).VKLatency > 0) Then
                    KeyState(cnt).VKPressed = True
                    IdleInput = Timer
                    KeyState(cnt).VKLatency = KeyState(cnt).VKLatency - 0.6
                Else
                    KeyState(cnt).VKPressed = False
                End If
            Else
                KeyState(cnt).VKPressed = False
                KeyState(cnt).VKLatency = 0
            End If
        End If

    Next

End Sub

Private Function InitKeys()

    KeyChars(2) = "1"
    KeyChars(3) = "2"
    KeyChars(4) = "3"
    KeyChars(5) = "4"
    KeyChars(6) = "5"
    KeyChars(7) = "6"
    KeyChars(8) = "7"
    KeyChars(9) = "8"
    KeyChars(10) = "9"
    KeyChars(11) = "0"
    KeyChars(12) = "-"
    KeyChars(13) = "="
    KeyChars(16) = "q"
    KeyChars(17) = "w"
    KeyChars(18) = "e"
    KeyChars(19) = "r"
    KeyChars(20) = "t"
    KeyChars(21) = "y"
    KeyChars(22) = "u"
    KeyChars(23) = "i"
    KeyChars(24) = "o"
    KeyChars(25) = "p"
    KeyChars(26) = "["
    KeyChars(27) = "]"
    
    KeyChars(30) = "a"
    KeyChars(31) = "s"
    KeyChars(32) = "d"
    KeyChars(33) = "f"
    KeyChars(34) = "g"
    KeyChars(35) = "h"
    KeyChars(36) = "j"
    KeyChars(37) = "k"
    KeyChars(38) = "l"
    KeyChars(39) = ";"
    KeyChars(40) = "'"
    KeyChars(43) = "\"
    KeyChars(44) = "z"
    KeyChars(45) = "x"
    KeyChars(46) = "c"
    KeyChars(47) = "v"
    KeyChars(48) = "b"
    KeyChars(49) = "n"
    KeyChars(50) = "m"
    KeyChars(51) = ","
    KeyChars(52) = "."
    KeyChars(53) = "/"
    KeyChars(55) = "*"
    KeyChars(57) = " "

    KeyChars(71) = "7"
    KeyChars(72) = "8"
    KeyChars(73) = "9"
    KeyChars(75) = "4"
    KeyChars(76) = "5"
    KeyChars(77) = "6"
    KeyChars(79) = "1"
    KeyChars(80) = "2"
    KeyChars(81) = "3"
    KeyChars(82) = "0"
    KeyChars(74) = "-"
    KeyChars(78) = "+"
    KeyChars(83) = "."
    KeyChars(181) = "/"

End Function

Public Sub CreateCmds()
    StitchInfo(StitchNum.BackSlashThick) = StitchBit.BackSlashThick
    StitchInfo(StitchNum.BackSlashThin) = StitchBit.BackSlashThin
    StitchInfo(StitchNum.BottomEdgeThick) = StitchBit.BottomEdgeThick
    StitchInfo(StitchNum.BottomEdgeThin) = StitchBit.BottomEdgeThin
    StitchInfo(StitchNum.ForwardSlashThick) = StitchBit.ForwardSlashThick
    StitchInfo(StitchNum.ForwardSlashThin) = StitchBit.ForwardSlashThin
    StitchInfo(StitchNum.LeftEdgeThick) = StitchBit.LeftEdgeThick
    StitchInfo(StitchNum.LeftEdgeThin) = StitchBit.LeftEdgeThin
    StitchInfo(StitchNum.TopEdgeThick) = StitchBit.TopEdgeThick
    StitchInfo(StitchNum.TopEdgeThin) = StitchBit.TopEdgeThin
    
    Set SkySkin(0) = LoadTexture(AppPath & "Base\top.bmp")
    Set SkySkin(1) = LoadTexture(AppPath & "Base\back.bmp")
    Set SkySkin(2) = LoadTexture(AppPath & "Base\left.bmp")
    Set SkySkin(3) = LoadTexture(AppPath & "Base\front.bmp")
    Set SkySkin(4) = LoadTexture(AppPath & "Base\right.bmp")
    Set SkySkin(5) = LoadTexture(AppPath & "Base\bottom.bmp")
                                
    CreateSquare SkyPlaq, 0, MakeVector(-5, -5, 5), _
                            MakeVector(-5, -5, -5), _
                            MakeVector(-5, 5, -5), _
                            MakeVector(-5, 5, 5), 1, 1
    CreateSquare SkyPlaq, 6, MakeVector(-5, -5, -5), _
                            MakeVector(5, -5, -5), _
                            MakeVector(5, 5, -5), _
                            MakeVector(-5, 5, -5), 1, 1
    CreateSquare SkyPlaq, 12, MakeVector(5, -5, -5), _
                            MakeVector(5, -5, 5), _
                            MakeVector(5, 5, 5), _
                            MakeVector(5, 5, -5), 1, 1
    CreateSquare SkyPlaq, 18, MakeVector(5, -5, -5), _
                            MakeVector(-5, -5, -5), _
                            MakeVector(-5, -5, 5), _
                            MakeVector(5, -5, 5), 1, 1
    CreateSquare SkyPlaq, 24, MakeVector(5, -5, 5), _
                            MakeVector(-5, -5, 5), _
                            MakeVector(-5, 5, 5), _
                            MakeVector(5, 5, 5), 1, 1
    CreateSquare SkyPlaq, 30, MakeVector(5, 5, 5), _
                            MakeVector(-5, 5, 5), _
                            MakeVector(-5, 5, -5), _
                            MakeVector(5, 5, -5), 1, 1

    Set SkyVBuf = DDevice.CreateVertexBuffer(Len(SkyPlaq(0)) * (UBound(SkyPlaq) + 1), 0, FVF_RENDER, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData SkyVBuf, 0, Len(SkyPlaq(0)) * (UBound(SkyPlaq) + 1), 0, SkyPlaq(0)
                        
'    Bottom = 0
'
'    Vertex(0).x = 0
'    Vertex(0).Z = -1
'    Vertex(0).RHW = 1
'    Vertex(0).Color = D3DColorARGB(255, 255, 255, 255)
'    Vertex(0).tu = 0
'    Vertex(0).tv = 0
'
'    Vertex(1).x = (frmMain.Width / Screen.TwipsPerPixelX)
'    Vertex(1).Z = -1
'    Vertex(1).RHW = 1
'    Vertex(1).Color = D3DColorARGB(255, 255, 255, 255)
'    Vertex(1).tu = 1
'    Vertex(1).tv = 0
'
'    Vertex(2).x = 0
'    Vertex(2).Z = -1
'    Vertex(2).RHW = 1
'    Vertex(2).Color = D3DColorARGB(255, 255, 255, 255)
'    Vertex(2).tu = 0
'    Vertex(2).tv = 1
'
'    Vertex(3).x = (frmMain.Width / Screen.TwipsPerPixelX)
'    Vertex(3).Z = -1
'    Vertex(3).RHW = 1
'    Vertex(3).Color = D3DColorARGB(255, 255, 255, 255)
'    Vertex(3).tu = 1
'    Vertex(3).tv = 1
'
'    ConsoleWidth = (frmMain.Width / Screen.TwipsPerPixelX)
'    ConsoleHeight = (MaxConsoleMsgs * (frmMain.TextHeight("A") / Screen.TwipsPerPixelY)) + (TextSpace * MaxConsoleMsgs) + TextSpace
'    If ConsoleHeight > ((frmMain.Height / Screen.TwipsPerPixelY) \ 2) Then ConsoleHeight = ((frmMain.Height / Screen.TwipsPerPixelY) \ 2)
'
    InitKeys
    
'    Set Backdrop = LoadTexture(AppPath & "Base\drop.bmp")
    
    Set DInput = dx.DirectInputCreate()
        
    Set DIKeyBoardDevice = DInput.CreateDevice("GUID_SysKeyboard")
    DIKeyBoardDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
    DIKeyBoardDevice.SetCooperativeLevel frmMain.hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    DIKeyBoardDevice.Acquire
    
    Set DIMouseDevice = DInput.CreateDevice("GUID_SysMouse")
    DIMouseDevice.SetCommonDataFormat DIFORMAT_MOUSE
    DIMouseDevice.SetCooperativeLevel frmMain.hwnd, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND
    DIMouseDevice.Acquire

End Sub

Public Sub CleanupCmds()

    DIKeyBoardDevice.Unacquire
    Set DIKeyBoardDevice = Nothing
    
    DIMouseDevice.Unacquire
    Set DIMouseDevice = Nothing
    
    Set DInput = Nothing
    
    Set SkySkin(0) = Nothing
    Set SkySkin(1) = Nothing
    Set SkySkin(2) = Nothing
    Set SkySkin(3) = Nothing
    Set SkySkin(4) = Nothing
    Set SkySkin(5) = Nothing
    
    Set SkyVBuf = Nothing
End Sub





