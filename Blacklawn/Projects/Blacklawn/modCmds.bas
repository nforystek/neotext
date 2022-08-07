Attribute VB_Name = "modCmds"
#Const modCmds = -1
Option Explicit
'TOP DOWN
Option Compare Binary


Option Private Module
Private DInput As DirectInput8
Private DIKeyBoardDevice As DirectInputDevice8
Private DIKEYBOARDSTATE As DIKEYBOARDSTATE

Private DIMouseDevice As DirectInputDevice8
Private DIMOUSESTATE As DIMOUSESTATE

Private ToggleSound1 As Boolean
Private ToggleSound2 As Boolean
Private ToggleSound3 As Boolean
Private TogglePress1 As Long
Private TogglePress2 As Long

Private TogglePress(0 To 2) As Long

Private Const MaxConsoleMsgs = 20
Private Const MaxHistoryMsgs = 10

Private ConsoleMsgs As Collection
Private HistoryMsgs As Collection
Private HistoryPoint As Integer
Private CommandLine As String

Private Type KeyState
    VKState As Integer
    VKToggle As Boolean
    VKPressed As Boolean
    VKLatency As Single
 End Type

Private CapsLock As Boolean
Private KeyState(255) As KeyState
Private KeyChars(255) As String

Private CursorPos As Long
Private CursorX As Long
Private CursorWidth As Long
Private ConsoleWidth As Single
Private ConsoleHeight As Single

Private Backdrop As Direct3DTexture8
Private Vertex(0 To 4) As MyScreen

Private State As Integer
Private Shift As Integer
Private Bottom As Long

Private EditFileName As String
Private EditFileData As String


Public Property Get TextHeight() As Single
    TextHeight = frmMain.TextHeight("A")
End Property

Public Property Get TextSpace() As Single
    TextSpace = 2
End Property

Public Sub InputScene()
        
    On Error GoTo pausing
    
    If ((frmMain.Recording And frmMain.IsPlayback) Or frmMain.IsPlayback) And (((Not PauseGame) And TrapMouse) And (Not ConsoleVisible)) Then
        frmMain.NextControl
    End If
    
    DIKeyBoardDevice.GetDeviceStateKeyboard DIKEYBOARDSTATE
    
    If (GetActiveWindow = frmMain.hwnd) Then
        If FullScreen And Not TrapMouse Then TrapMouse = True
        
        ConsoleInput DIKEYBOARDSTATE
        
        If DIKEYBOARDSTATE.Key(DIK_ESCAPE) Then
            If (Not TogglePress1 = DIK_ESCAPE) Then
                TogglePress1 = DIK_ESCAPE
                If ((Not FullScreen) And TrapMouse) Or ConsoleVisible Then
                    TrapMouse = False
                ElseIf FullScreen Or (Not FullScreen And Not TrapMouse) Then
                    If frmMain.IsPlayback Or frmMain.Recording Then
                        StopFilm
                    ElseIf frmMain.Multiplayer Then
                        frmMain.Disconnect
                    End If
                    StopGame = True
                End If
            End If
        ElseIf DIKEYBOARDSTATE.Key(DIK_F1) Then
            If (Not TogglePress1 = DIK_F1) Then
                TogglePress1 = DIK_F1
                HelpToggle = Not HelpToggle
            End If
        ElseIf DIKEYBOARDSTATE.Key(DIK_GRAVE) Then
            If (Not TogglePress1 = DIK_GRAVE) Then
                TogglePress1 = DIK_GRAVE
                ConsoleToggle
                PlayWave SOUND_TOGGLE
            End If
        ElseIf DIKEYBOARDSTATE.Key(DIK_LALT) Or DIKEYBOARDSTATE.Key(DIK_RALT) Then
            If DIKEYBOARDSTATE.Key(DIK_TAB) Then
                If (Not TogglePress1 = DIK_TAB) Then
                    TogglePress1 = DIK_TAB
                    TrapMouse = False
                    frmMain.WindowState = 1
                End If
            End If
        ElseIf Not TogglePress1 = 0 Then
            TogglePress1 = 0
        End If
        
    End If
       
    If (((Not PauseGame) And TrapMouse) And (Not ConsoleVisible)) Or frmMain.IsPlayback Then
    
        If (frmMain.Recording And Not frmMain.IsPlayback) Or frmMain.Multiplayer Then frmMain.Toggler(1) = DIKEYBOARDSTATE.Key(DIK_1)
        If (DIKEYBOARDSTATE.Key(DIK_1) And ((Not frmMain.Recording) Or (frmMain.Recording And Not frmMain.IsPlayback))) Or (frmMain.Toggler(1) And (frmMain.Recording And frmMain.IsPlayback)) Then
            If (Not TogglePress(1) = DIK_1) Then
                TogglePress(1) = DIK_1
                SetGainMode1
                
                If frmMain.Recording And Not frmMain.IsPlayback Then
                    WarpData = WarpData & Player.Object.Origin.X & "," & Player.Object.Origin.Y & "," & Player.Object.Origin.z & ","
                ElseIf frmMain.Recording And frmMain.IsPlayback Then
                    Player.Object.Origin.X = RemoveNextArg(WarpData, ",")
                    Player.Object.Origin.Y = RemoveNextArg(WarpData, ",")
                    Player.Object.Origin.z = RemoveNextArg(WarpData, ",")
                    If Clocker.FollowingMode Then
                        Partner.Object.Origin.X = Player.Object.Origin.X
                        Partner.Object.Origin.Y = Player.Object.Origin.Y
                        Partner.Object.Origin.z = Player.Object.Origin.z
                    End If
                End If
                
                XthTime = ""
            ElseIf TogglePress(1) = DIK_1 Then
                 TogglePress(1) = 0
            End If
            
        End If
        
        If (frmMain.Recording And Not frmMain.IsPlayback) Or frmMain.Multiplayer Then frmMain.Toggler(2) = DIKEYBOARDSTATE.Key(DIK_2)
        If (DIKEYBOARDSTATE.Key(DIK_2) And ((Not frmMain.Recording) Or (frmMain.Recording And Not frmMain.IsPlayback))) Or (frmMain.Toggler(2) And (frmMain.Recording And frmMain.IsPlayback)) Then
            If (Not TogglePress(2) = DIK_2) Then
                TogglePress(2) = DIK_2
                SetGainMode2
                
                If frmMain.Recording And Not frmMain.IsPlayback Then
                    WarpData = WarpData & Player.Object.Origin.X & "," & Player.Object.Origin.Y & "," & Player.Object.Origin.z & ","
                ElseIf frmMain.Recording And frmMain.IsPlayback Then
                    Player.Object.Origin.X = RemoveNextArg(WarpData, ",")
                    Player.Object.Origin.Y = RemoveNextArg(WarpData, ",")
                    Player.Object.Origin.z = RemoveNextArg(WarpData, ",")
                    If Clocker.FollowingMode Then
                        Partner.Object.Origin.X = Player.Object.Origin.X
                        Partner.Object.Origin.Y = Player.Object.Origin.Y
                        Partner.Object.Origin.z = Player.Object.Origin.z
                    End If

                End If
                
                XthTime = ""
            ElseIf TogglePress(2) = DIK_2 Then
                 TogglePress(2) = 0
            End If
        End If
        
        If (frmMain.Recording And Not frmMain.IsPlayback) Or frmMain.Multiplayer Then frmMain.Toggler(3) = DIKEYBOARDSTATE.Key(DIK_0)
        If (DIKEYBOARDSTATE.Key(DIK_0) And ((Not frmMain.Recording) Or (frmMain.Recording And Not frmMain.IsPlayback))) Or (frmMain.Toggler(3) And (frmMain.Recording And frmMain.IsPlayback)) Then
            If (Not TogglePress(0) = DIK_0) Then
                TogglePress(0) = DIK_0
                SetGainMode0
                
                If frmMain.Recording And Not frmMain.IsPlayback Then
                    WarpData = WarpData & Player.Object.Origin.X & "," & Player.Object.Origin.Y & "," & Player.Object.Origin.z & ","
                ElseIf frmMain.Recording And frmMain.IsPlayback Then
                    Player.Object.Origin.X = RemoveNextArg(WarpData, ",")
                    Player.Object.Origin.Y = RemoveNextArg(WarpData, ",")
                    Player.Object.Origin.z = RemoveNextArg(WarpData, ",")
                    If Clocker.FollowingMode Then
                        Partner.Object.Origin.X = Player.Object.Origin.X
                        Partner.Object.Origin.Y = Player.Object.Origin.Y
                        Partner.Object.Origin.z = Player.Object.Origin.z
                    End If

                End If
                
                XthTime = ""
            ElseIf TogglePress(0) = DIK_0 Then
                 TogglePress(0) = 0
            End If
        End If
        
        If (frmMain.Recording And Not frmMain.IsPlayback) Or frmMain.Multiplayer Then frmMain.Toggler(4) = DIKEYBOARDSTATE.Key(DIK_S)
        If (DIKEYBOARDSTATE.Key(DIK_S) And ((Not frmMain.Recording) Or (frmMain.Recording And Not frmMain.IsPlayback))) Or (frmMain.Toggler(4) And (frmMain.Recording And frmMain.IsPlayback)) Then
            Player.Rotation = Player.Rotation + 0.05
            If Not ToggleSound1 Then
                ToggleSound1 = True
                PlayWave SOUND_JUICE
                PartnersDancing
            End If
            XthTime = ""
        End If
        
        If (frmMain.Recording And Not frmMain.IsPlayback) Or frmMain.Multiplayer Then frmMain.Toggler(5) = DIKEYBOARDSTATE.Key(DIK_F)
        If (DIKEYBOARDSTATE.Key(DIK_F) And ((Not frmMain.Recording) Or (frmMain.Recording And Not frmMain.IsPlayback))) Or (frmMain.Toggler(5) And (frmMain.Recording And frmMain.IsPlayback)) Then
            Player.Rotation = Player.Rotation - 0.05
            If Not ToggleSound1 Then
                ToggleSound1 = True
                PlayWave SOUND_JUICE
                PartnersDancing
            End If
            XthTime = ""
        End If

        If Player.Rotation > (PI * 2) Then Player.Rotation = Player.Rotation - (PI * 2)
        If Player.Rotation < -(PI * 2) Then Player.Rotation = Player.Rotation + (PI * 2)


        If (frmMain.Recording And Not frmMain.IsPlayback) Or frmMain.Multiplayer Then frmMain.Toggler(15) = DIKEYBOARDSTATE.Key(DIK_Q)
        If (DIKEYBOARDSTATE.Key(DIK_Q) And ((Not frmMain.Recording) Or (frmMain.Recording And Not frmMain.IsPlayback))) Or (frmMain.Toggler(15) And (frmMain.Recording And frmMain.IsPlayback)) Then
            Player.MoveSpeed = Player.MoveSpeed + MoveSpeedInc
            If Player.MoveSpeed > MoveSpeedMax Then Player.MoveSpeed = MoveSpeedMax
            XthTime = ""
        End If
        
        If (frmMain.Recording And Not frmMain.IsPlayback) Or frmMain.Multiplayer Then frmMain.Toggler(16) = DIKEYBOARDSTATE.Key(DIK_A)
        If (DIKEYBOARDSTATE.Key(DIK_A) And ((Not frmMain.Recording) Or (frmMain.Recording And Not frmMain.IsPlayback))) Or (frmMain.Toggler(16) And (frmMain.Recording And frmMain.IsPlayback)) Then
            Player.MoveSpeed = Player.MoveSpeed - MoveSpeedInc
            If Player.MoveSpeed < MoveSpeedMin Then Player.MoveSpeed = MoveSpeedMin
            XthTime = ""
        End If
        
        If (frmMain.Recording And Not frmMain.IsPlayback) Or frmMain.Multiplayer Then frmMain.Toggler(6) = DIKEYBOARDSTATE.Key(DIK_T)
        If (DIKEYBOARDSTATE.Key(DIK_T) And ((Not frmMain.Recording) Or (frmMain.Recording And Not frmMain.IsPlayback))) Or (frmMain.Toggler(6) And (frmMain.Recording And frmMain.IsPlayback)) Then
            If (Not TogglePress2 = DIK_T) Then
                TogglePress2 = DIK_T
                If Not ConsoleVisible Then
                    Player.AutoMove = Not Player.AutoMove
                End If
            End If
            XthTime = ""
        ElseIf Not (TogglePress2 = 0) Then
            TogglePress2 = 0
        End If

        Dim vecDirect As D3DVECTOR

        If Not Player.AutoMove Then
        
            If (frmMain.Recording And Not frmMain.IsPlayback) Or frmMain.Multiplayer Then frmMain.Toggler(7) = DIKEYBOARDSTATE.Key(DIK_E)
            If (DIKEYBOARDSTATE.Key(DIK_E) And ((Not frmMain.Recording) Or (frmMain.Recording And Not frmMain.IsPlayback))) Or (frmMain.Toggler(7) And (frmMain.Recording And frmMain.IsPlayback)) Then
                If Not Player.Stalled Then
                    vecDirect.X = Sin(D720 - Player.CameraAngle)
                    vecDirect.z = Cos(D720 - Player.CameraAngle)
                    AddActivity Player.Object, vecDirect, Player.MoveSpeed, GroundFriction
                End If
                XthTime = ""
            End If
            
            If (frmMain.Recording And Not frmMain.IsPlayback) Or frmMain.Multiplayer Then frmMain.Toggler(8) = DIKEYBOARDSTATE.Key(DIK_D)
            If (DIKEYBOARDSTATE.Key(DIK_D) And ((Not frmMain.Recording) Or (frmMain.Recording And Not frmMain.IsPlayback))) Or (frmMain.Toggler(8) And (frmMain.Recording And frmMain.IsPlayback)) Then
                If Not Player.Stalled Then
                    vecDirect.X = -Sin(D720 - Player.CameraAngle)
                    vecDirect.z = -Cos(D720 - Player.CameraAngle)
                    AddActivity Player.Object, vecDirect, Player.MoveSpeed, GroundFriction
                End If
                XthTime = ""
            End If
            
            If (frmMain.Recording And Not frmMain.IsPlayback) Or frmMain.Multiplayer Then frmMain.Toggler(9) = DIKEYBOARDSTATE.Key(DIK_W)
            If (DIKEYBOARDSTATE.Key(DIK_W) And ((Not frmMain.Recording) Or (frmMain.Recording And Not frmMain.IsPlayback))) Or (frmMain.Toggler(9) And (frmMain.Recording And frmMain.IsPlayback)) Then
                If Not Player.Stalled Then
                    vecDirect.X = Sin((D720 - Player.CameraAngle) - D180)
                    vecDirect.z = Cos((D720 - Player.CameraAngle) - D180)
                    AddActivity Player.Object, vecDirect, Player.MoveSpeed, GroundFriction
                End If
                XthTime = ""
            End If
            
            If (frmMain.Recording And Not frmMain.IsPlayback) Or frmMain.Multiplayer Then frmMain.Toggler(10) = DIKEYBOARDSTATE.Key(DIK_R)
            If (DIKEYBOARDSTATE.Key(DIK_R) And ((Not frmMain.Recording) Or (frmMain.Recording And Not frmMain.IsPlayback))) Or (frmMain.Toggler(10) And (frmMain.Recording And frmMain.IsPlayback)) Then
                If Not Player.Stalled Then
                    vecDirect.X = Sin((D720 - Player.CameraAngle) + D180)
                    vecDirect.z = Cos((D720 - Player.CameraAngle) + D180)
                    AddActivity Player.Object, vecDirect, Player.MoveSpeed, GroundFriction
                End If
                XthTime = ""
            End If
        Else
            vecDirect.X = Sin(D720 - Player.CameraAngle)
            vecDirect.z = Cos(D720 - Player.CameraAngle)
            AddActivity Player.Object, vecDirect, Player.MoveSpeed, GroundFriction
            XthTime = ""
        End If

        If (frmMain.Recording And Not frmMain.IsPlayback) Or frmMain.Multiplayer Then frmMain.Toggler(11) = DIKEYBOARDSTATE.Key(DIK_SPACE)
        If (DIKEYBOARDSTATE.Key(DIK_SPACE) And ((Not frmMain.Recording) Or (frmMain.Recording And Not frmMain.IsPlayback))) Or (frmMain.Toggler(11) And (frmMain.Recording And frmMain.IsPlayback)) Then
            Player.Object.Origin.Y = Player.Object.Origin.Y + (GravityVelocity * ((Player.MoveSpeed - MoveSpeedMin) / MoveSpeedMax))
            Player.Gravity = -1
            XthTime = ""
        End If


        Player.Gravity = Player.Gravity + 1
        If Player.Gravity > MaxGravity Then Player.Gravity = MaxGravity
            
        If Not (Player.Object.Origin.Y = 0) Then

            If (Player.Gravity > 0) Then
                If ((Player.Object.Origin.Y - Player.Gravity) >= 0) Or (Player.Object.Origin.Y < 0) Then
                    Player.Object.Origin.Y = Player.Object.Origin.Y - Player.Gravity
                Else
                    Player.Object.Origin.Y = 0
                    Player.Gravity = 0
                    PlayWave SOUND_BOOM
                End If
                
            Else
                Player.Object.Origin.Y = Player.Object.Origin.Y + GravityVelocity
            End If
        ElseIf (Player.Object.Origin.Y = 0) Then
            If (Player.Gravity < 0) Then
                Player.Object.Origin.Y = Player.Object.Origin.Y + GravityVelocity
            End If
        End If
    
    End If
    
    DIMouseDevice.GetDeviceStateMouse DIMOUSESTATE
    
    Dim mX As Integer
    Dim mY As Integer
    Dim mZ As Integer
    
    mX = DIMOUSESTATE.lX
    mY = DIMOUSESTATE.lY
    mZ = DIMOUSESTATE.lZ
    
    If ((GetActiveWindow = frmMain.hwnd) And TrapMouse) Then
    
        If (((Not PauseGame) And TrapMouse) And (Not ConsoleVisible)) Or frmMain.IsPlayback Then
            mX = mX + IIf(DIKEYBOARDSTATE.Key(DIK_RIGHT), MouseSensitivity, IIf(DIKEYBOARDSTATE.Key(DIK_LEFT), -MouseSensitivity, 0))
            mY = mY + IIf(DIKEYBOARDSTATE.Key(DIK_UP), -MouseSensitivity, IIf(DIKEYBOARDSTATE.Key(DIK_DOWN), MouseSensitivity, 0))
            mZ = mZ + IIf(DIKEYBOARDSTATE.Key(DIK_PGUP), MouseSensitivity, IIf(DIKEYBOARDSTATE.Key(DIK_PGDN), -MouseSensitivity, 0))
               
            If (frmMain.Recording And frmMain.IsPlayback) Then
                mX = frmMain.MouseX
                mY = frmMain.MouseY
                mZ = frmMain.MouseZ
            ElseIf (frmMain.Recording And Not frmMain.IsPlayback) Or frmMain.Multiplayer Then
                frmMain.MouseX = mX
                frmMain.MouseY = mY
                frmMain.MouseZ = mZ
            End If
        End If
    
        If Not (frmMain.MousePointer = 99) Then
            frmMain.MousePointer = 99
            frmMain.MouseIcon = LoadPicture(AppPath & "Base\mouse.cur")
        End If
        
        MouseLook mX, mY, mZ
        
        If (((Not PauseGame) And TrapMouse) And (Not ConsoleVisible)) Or frmMain.IsPlayback Then

            If (frmMain.Recording And Not frmMain.IsPlayback) Or frmMain.Multiplayer Then frmMain.Toggler(12) = (DIMOUSESTATE.Buttons(2) Or DIKEYBOARDSTATE.Key(DIK_Y))
            If ((DIMOUSESTATE.Buttons(2) Or DIKEYBOARDSTATE.Key(DIK_Y)) And ((Not frmMain.Recording) Or (frmMain.Recording And Not frmMain.IsPlayback))) Or (frmMain.Toggler(12) And (frmMain.Recording And frmMain.IsPlayback)) Then 'mid
                Player.FlapLock = True
                If Not ToggleSound2 Then
                    ToggleSound2 = True
                    PlayWave SOUND_SMACK
                    PartnersDancing
                End If
                XthTime = ""
            Else
                Player.FlapLock = False
                
                If (frmMain.Recording And Not frmMain.IsPlayback) Or frmMain.Multiplayer Then frmMain.Toggler(13) = (DIMOUSESTATE.Buttons(0) Or DIKEYBOARDSTATE.Key(DIK_G))
                If ((DIMOUSESTATE.Buttons(0) Or DIKEYBOARDSTATE.Key(DIK_G)) And ((Not frmMain.Recording) Or (frmMain.Recording And Not frmMain.IsPlayback))) Or (frmMain.Toggler(13) And (frmMain.Recording And frmMain.IsPlayback)) Then 'left
                    Player.LeftFlap = True
                    If Not ToggleSound3 Then
                        ToggleSound3 = True
                        PlayWave SOUND_BANG
                        PartnersDancing
                    End If
                    XthTime = ""
                Else
                     Player.LeftFlap = False
                End If
                
                If (frmMain.Recording And Not frmMain.IsPlayback) Or frmMain.Multiplayer Then frmMain.Toggler(14) = (DIMOUSESTATE.Buttons(1) Or DIKEYBOARDSTATE.Key(DIK_H))
                If ((DIMOUSESTATE.Buttons(1) Or DIKEYBOARDSTATE.Key(DIK_H)) And ((Not frmMain.Recording) Or (frmMain.Recording And Not frmMain.IsPlayback))) Or (frmMain.Toggler(14) And (frmMain.Recording And frmMain.IsPlayback)) Then 'right
                    Player.RightFlap = True
                    If Not ToggleSound3 Then
                        ToggleSound3 = True
                        PlayWave SOUND_BANG
                        PartnersDancing
                    End If
                    XthTime = ""
                Else
                    Player.RightFlap = False
                End If
                
            End If
            
        End If
        
        Dim rec As RECT
        GetWindowRect frmMain.hwnd, rec
        SetCursorPos rec.Right + ((rec.Left - rec.Right) / 2), rec.Top + ((rec.Bottom - rec.Top) / 2)
    Else
        If Not (frmMain.MousePointer = 1) Then frmMain.MousePointer = 1
    End If

    If (Not (DIKEYBOARDSTATE.Key(DIK_S) = 128)) And (Not (DIKEYBOARDSTATE.Key(DIK_F) = 128)) Then
        ToggleSound1 = False
    End If
    If (Not Player.FlapLock) Then
        ToggleSound2 = False
    End If
    If (Not Player.LeftFlap) And (Not Player.RightFlap) Then
        ToggleSound3 = False
    End If
    
    If (Player.Object.Origin.X > BlackBoundary) Or (Player.Object.Origin.X < -BlackBoundary) Then Player.Object.Origin.X = -Player.Object.Origin.X
    If (Player.Object.Origin.z > BlackBoundary) Or (Player.Object.Origin.z < -BlackBoundary) Then Player.Object.Origin.z = -Player.Object.Origin.z
    
    If frmMain.Multiplayer Then
        frmMain.StreamControls
    ElseIf (frmMain.Recording And (Not frmMain.IsPlayback)) Then
        frmMain.AddControls CStr(Partner.Object.Origin.X) & "," & CStr(Partner.Object.Origin.Y) & "," & CStr(Partner.Object.Origin.z) & "," & CStr(Partner.Rotation) & "," & IIf(frmMain.MouseX = 0, "", frmMain.MouseX) & "," & IIf(frmMain.MouseY = 0, "", frmMain.MouseY) & "," & IIf(frmMain.MouseZ = 0, "", frmMain.MouseZ) & "," & frmMain.ToggleString
    End If
    
    Exit Sub
pausing:
    Err.Clear
    DoPauseGame
End Sub

Private Sub MouseLook(ByVal X As Integer, ByVal Y As Integer, ByVal z As Integer)

    Dim cnt As Long
    If z < 0 Then
        For cnt = (z * MouseSensitivity) To 0
            Player.CameraZoom = Player.CameraZoom + 0.15
        Next
        XthTime = ""
    ElseIf z > 0 Then
        For cnt = 0 To (z * MouseSensitivity)
            Player.CameraZoom = Player.CameraZoom - 0.15
        Next
        XthTime = ""
    End If
    
    If Player.CameraZoom > 3000 Then Player.CameraZoom = 3000
    If Player.CameraZoom < 200 Then Player.CameraZoom = 200

    If X < 0 Then
        For cnt = (X * MouseSensitivity) To 0
            Player.CameraAngle = Player.CameraAngle - -0.0015
        Next
        XthTime = ""
    ElseIf X > 0 Then
        For cnt = 0 To (X * MouseSensitivity)
            Player.CameraAngle = Player.CameraAngle - 0.0015
        Next
        XthTime = ""
    End If
    
    If Player.CameraAngle > (PI * 2) Then Player.CameraAngle = Player.CameraAngle - (PI * 2)
    If Player.CameraAngle < -(PI * 2) Then Player.CameraAngle = Player.CameraAngle + (PI * 2)
    
    If Y < 0 Then
        For cnt = (Y * MouseSensitivity) To 0
            Player.CameraPitch = Player.CameraPitch - -0.0015
        Next
        XthTime = ""
    ElseIf Y > 0 Then
        For cnt = 0 To (Y * MouseSensitivity)
            Player.CameraPitch = Player.CameraPitch - 0.0015
        Next
        XthTime = ""
    End If
    
    If Player.CameraPitch < -1.5 Then Player.CameraPitch = -1.5
    If Player.CameraPitch > 1.5 Then Player.CameraPitch = 1.5
    
End Sub

Public Property Get ConsoleVisible() As Boolean
    ConsoleVisible = Not (State = 0)
End Property

Public Sub ConsoleToggle()
    If State = 1 Then
        Shift = -Shift
    ElseIf State = 0 Then
        Shift = 18
        State = 1
    ElseIf State = 2 Then
        Shift = -18
        State = 1
    End If
End Sub

Public Sub RenderCmds()

    Select Case State
        Case 0 'up
            Bottom = 0
            Shift = 0
        Case 2 'down
            Bottom = ConsoleHeight
            Shift = 0
        Case 1 'moving
            If (Bottom + Shift) > ConsoleHeight Then
                State = 2
            ElseIf (Bottom + Shift) < 0 Then
                State = 0
            Else
                Bottom = Bottom + Shift
            End If
    End Select

    Vertex(0).Y = Bottom - ConsoleHeight
    Vertex(1).Y = Bottom - ConsoleHeight
    Vertex(2).Y = Bottom
    Vertex(3).Y = Bottom
    
    If ConsoleVisible Then
    
        DDevice.SetVertexShader FVF_SCREEN
        DDevice.SetRenderState D3DRS_ZENABLE, False
        DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

        DDevice.SetRenderState D3DRS_LIGHTING, False
        DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    
        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, False
    
        DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
        DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
    
        DDevice.SetTexture 0, Backdrop
        DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0), LenB(Vertex(0))
    
        If Len(CommandLine) > 0 Then
            DrawText Replace(CommandLine, vbTab, "     "), TextSpace + 2, Bottom - (TextHeight / Screen.TwipsPerPixelY) - TextSpace + 2
        End If
        
        Static DrawCursor As Double
        Static DrawBlink As Boolean
        If DrawCursor = 0 Or CDbl(((GetTimer * 1000) - DrawCursor)) >= 2000 Then
            DrawCursor = GetTimer
            DrawBlink = Not DrawBlink
            
            If DrawBlink Then DrawText String(CursorPos + (CountWord(Left(CommandLine, CursorPos), vbTab) * 4), " ") & "_", TextSpace + 2, Bottom - (TextHeight / Screen.TwipsPerPixelY) - TextSpace + 2

        End If
        
        If ConsoleMsgs.Count > 0 Then
            Dim ConsoleMsgY As Long
            ConsoleMsgY = (Bottom - (((TextHeight / Screen.TwipsPerPixelY) + TextSpace) * 2))
            Dim cnt As Integer
            For cnt = ConsoleMsgs.Count To 1 Step -1
                DrawText ConsoleMsgs(cnt), TextSpace + 2, ConsoleMsgY - ((ConsoleMsgs.Count - cnt) * (((TextHeight / Screen.TwipsPerPixelY) + TextSpace)))
            Next
        End If
        
        DDevice.SetRenderState D3DRS_ZENABLE, 1
        
        DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, RENDER_MAGFILTER
        DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, RENDER_MINFILTER
    End If
End Sub

Public Function AddMessage(ByVal Message As String)
    If Not (ConsoleMsgs Is Nothing) Then
    
        If ConsoleMsgs.Count > 0 Then
            If (ConsoleMsgs.Item(ConsoleMsgs.Count) = Message) Then
                Exit Function
            End If
        End If
            
        If ConsoleMsgs.Count > MaxConsoleMsgs Then
            ConsoleMsgs.Remove 1
        End If
        ConsoleMsgs.Add Message
    
    End If
End Function


Private Function Toggled(ByVal vkCode As Long) As Boolean
    Toggled = KeyState(vkCode).VKToggle
End Function
Private Function Pressed(ByVal vkCode As Long) As Boolean
    Pressed = KeyState(vkCode).VKPressed
End Function


Public Sub ConsoleInput(ByRef kState As DIKEYBOARDSTATE)

    If ConsoleVisible And (Not (Shift < 0)) And (GetActiveWindow = frmMain.hwnd) Then
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
                    ElseIf ((Timer - KeyState(cnt).VKLatency) > 0.6) And (KeyState(cnt).VKLatency > 0) Then
                        KeyState(cnt).VKPressed = True
                        KeyState(cnt).VKLatency = KeyState(cnt).VKLatency - 0.6
                    Else
                        KeyState(cnt).VKPressed = False
                    End If
                Else
                    KeyState(cnt).VKPressed = False
                    KeyState(cnt).VKLatency = 0
                End If
            End If

            If Pressed(cnt) Then
            
                If cnt = DIK_CAPITAL Then
                    CapsLock = Not CapsLock
                ElseIf cnt = DIK_LSHIFT Or cnt = DIK_RSHIFT Then
                
                ElseIf cnt = DIK_RETURN Then
                    HistoryMsgs.Add CommandLine
                    If HistoryMsgs.Count > MaxHistoryMsgs Then
                        HistoryMsgs.Remove 1
                    End If
                    HistoryPoint = HistoryMsgs.Count + 1
                    
                    Process CommandLine
                    
                    CommandLine = ""
                    CursorPos = 0
                ElseIf cnt = DIK_BACK Then
                    If CursorPos > 0 Then
                        CommandLine = Left(CommandLine, CursorPos - 1) & Mid(CommandLine, CursorPos + 1)
                        CursorPos = CursorPos - 1
                    End If
                
                ElseIf cnt = DIK_DELETE Then
                    
                    CommandLine = Left(CommandLine, CursorPos) & Mid(CommandLine, CursorPos + 2)
                
                ElseIf cnt = DIK_LEFT Then
                    If CursorPos > 0 Then
                        CursorPos = CursorPos - 1
                    End If
                ElseIf cnt = DIK_HOME Then
                    CursorPos = 0
                ElseIf cnt = DIK_END Then
                    CursorPos = Len(CommandLine)
                ElseIf cnt = DIK_RIGHT Then
                    If CursorPos < Len(CommandLine) Then
                        CursorPos = CursorPos + 1
                    End If
                ElseIf cnt = DIK_UP Then
                    If HistoryMsgs.Count > 0 Then
                        If HistoryPoint > 1 Then
                            HistoryPoint = HistoryPoint - 1
                            CommandLine = HistoryMsgs(HistoryPoint)
                            CursorPos = Len(CommandLine)
                        End If
                    End If
                ElseIf cnt = DIK_DOWN Then
                    If HistoryPoint = HistoryMsgs.Count Then
                        HistoryPoint = HistoryPoint + 1
                        CommandLine = ""
                        CursorPos = 0
                    ElseIf HistoryPoint <= HistoryMsgs.Count Then
                    
                        HistoryPoint = HistoryPoint + 1
                        CommandLine = HistoryMsgs(HistoryPoint)
                        CursorPos = Len(CommandLine)
                    End If
                ElseIf cnt = DIK_TAB Then
                    If Len(CommandLine) <= 40 Then
                    
                        CommandLine = Left(CommandLine, CursorPos) & vbTab & Mid(CommandLine, CursorPos + 1)
                        CursorPos = CursorPos + 1
                        
                    End If
                Else
                
                    char = KeyChars(cnt)
                    If Not char = "" Then
                        If CapsLock Or kState.Key(DIK_LSHIFT) Or kState.Key(DIK_RSHIFT) Then
                            Select Case char
                                Case "1"
                                    char = "!"
                                Case "2"
                                    char = "@"
                                Case "3"
                                    char = "#"
                                Case "4"
                                    char = "$"
                                Case "5"
                                    char = "%"
                                Case "6"
                                    char = "^"
                                Case "7"
                                    char = "&"
                                Case "8"
                                    char = "*"
                                Case "9"
                                    char = "("
                                Case "0"
                                    char = ")"
                                Case "-"
                                    char = "_"
                                Case "="
                                    char = "+"
                                Case "["
                                    char = "{"
                                Case "]"
                                    char = "}"
                                Case "\"
                                    char = "|"
                                Case ";"
                                    char = ":"
                                Case "'"
                                    char = """"
                                Case ","
                                    char = "<"
                                Case "."
                                    char = ">"
                                Case "/"
                                    char = "?"
                                Case Else
                                    char = UCase(char)
                            End Select
                        End If
                        
                        If Len(CommandLine) <= 40 Then
                        
                            CommandLine = Left(CommandLine, CursorPos) & char & Mid(CommandLine, CursorPos + 1)
                            CursorPos = CursorPos + 1
                            
                        End If
                    
                    End If
                                            
                End If
                
                CursorX = (frmMain.TextWidth(Left(CommandLine, CursorPos)) / Screen.TwipsPerPixelX)
                CursorWidth = (frmMain.TextWidth(Mid(CommandLine, CursorPos + 1, 1)) / Screen.TwipsPerPixelX)
                If CursorWidth = 0 Then CursorWidth = 8
                
            End If
        Next
    End If
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
    
    Set ConsoleMsgs = New Collection
    Set HistoryMsgs = New Collection
    
    Bottom = 0
    
    Vertex(0).X = 0
    Vertex(0).z = -1
    Vertex(0).RHW = 1
    Vertex(0).Color = D3DColorARGB(255, 255, 255, 255)
    Vertex(0).tu = 0
    Vertex(0).tv = 0
    
    Vertex(1).X = (frmMain.width / Screen.TwipsPerPixelX)
    Vertex(1).z = -1
    Vertex(1).RHW = 1
    Vertex(1).Color = D3DColorARGB(255, 255, 255, 255)
    Vertex(1).tu = 1
    Vertex(1).tv = 0
    
    Vertex(2).X = 0
    Vertex(2).z = -1
    Vertex(2).RHW = 1
    Vertex(2).Color = D3DColorARGB(255, 255, 255, 255)
    Vertex(2).tu = 0
    Vertex(2).tv = 1
    
    Vertex(3).X = (frmMain.width / Screen.TwipsPerPixelX)
    Vertex(3).z = -1
    Vertex(3).RHW = 1
    Vertex(3).Color = D3DColorARGB(255, 255, 255, 255)
    Vertex(3).tu = 1
    Vertex(3).tv = 1
    
    ConsoleWidth = (frmMain.width / Screen.TwipsPerPixelX)
    ConsoleHeight = (MaxConsoleMsgs * (frmMain.TextHeight("A") / Screen.TwipsPerPixelY)) + (TextSpace * MaxConsoleMsgs) + TextSpace
    If ConsoleHeight > ((frmMain.height / Screen.TwipsPerPixelY) \ 2) Then ConsoleHeight = ((frmMain.height / Screen.TwipsPerPixelY) \ 2)
    
    InitKeys
    
    Set Backdrop = LoadTexture(AppPath & "Base\drop.bmp")
    
    Set DInput = dx.DirectInputCreate()
        
    Set DIKeyBoardDevice = DInput.CreateDevice("GUID_SysKeyboard")
    DIKeyBoardDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
    If FullScreen Then
        DIKeyBoardDevice.SetCooperativeLevel frmMain.hwnd, DISCL_EXCLUSIVE Or DISCL_FOREGROUND
    Else
        DIKeyBoardDevice.SetCooperativeLevel frmMain.hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    End If
    DIKeyBoardDevice.Acquire
    
    Set DIMouseDevice = DInput.CreateDevice("GUID_SysMouse")
    DIMouseDevice.SetCommonDataFormat DIFORMAT_MOUSE
    If FullScreen Then
        DIMouseDevice.SetCooperativeLevel frmMain.hwnd, DISCL_EXCLUSIVE Or DISCL_FOREGROUND
    Else
        DIMouseDevice.SetCooperativeLevel frmMain.hwnd, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND
    End If
    DIMouseDevice.Acquire
    
End Sub

Public Sub CleanupCmds()
    
    Do While ConsoleMsgs.Count > 0
        ConsoleMsgs.Remove 1
    Loop
    
    Do While HistoryMsgs.Count > 0
        HistoryMsgs.Remove 1
    Loop
    
    Set ConsoleMsgs = Nothing
    Set HistoryMsgs = Nothing
    
    Set Backdrop = Nothing
    
    DIKeyBoardDevice.Unacquire
    Set DIKeyBoardDevice = Nothing
    
    DIMouseDevice.Unacquire
    Set DIMouseDevice = Nothing
    
    Set DInput = Nothing
End Sub

Public Sub InitialCommands()

    Dim CommandsINI As String

    If PathExists(AppPath & "commands.ini") Then CommandsINI = ReadFile(AppPath & "commands.ini")
    
    Dim inLine As String
    Do Until CommandsINI = ""
        inLine = RemoveNextArg(CommandsINI, vbCrLf)
        Process inLine
    Loop
    
End Sub

Private Sub Process(ByVal inArg As String)
    On Error GoTo processerr
    
    Dim o As Long
    Dim l As Long
    Dim cnt As Long
    Dim inNew As String
    Dim inTmp As String
    Dim inCmd As String
    
    inCmd = RemoveNextArg(inArg, " ")
    If Left(inCmd, 1) = "/" Then inCmd = Mid(inCmd, 2)
    
    Select Case LCase(inCmd)
        Case "stat"
            If (Not frmMain.Recording) And (Not frmMain.IsPlayback) And (Not frmMain.Multiplayer) Then

                AddMessage "Distance: " & CLng(Distance(Player.Object.Origin.X, Player.Object.Origin.Y, Player.Object.Origin.z, 0, 0, 0))
                AddMessage "Location: " & CLng(Player.Object.Origin.X) & ", " & CLng(Player.Object.Origin.Y) & ", " & CLng(Player.Object.Origin.z)
                AddMessage "Camera Angle: " & Player.CameraAngle
                AddMessage "Camera Pitch: " & Player.CameraPitch
                AddMessage "Camera Zoom: " & Player.CameraZoom
            Else
                AddMessage "Stat command is currently disabled..."
            End If
        Case "refresh"

            TermGameData
            InitGameData
            AddMessage "Game data refresh finished."

        Case "load"
            If inArg = "" Then
                If EditFileName = "" Then
                    AddMessage "No file is loaded, use ""LOAD <name>"" to load one."
                Else
                    AddMessage "File loaded [" & AppPath & "Base\Lawn\" & EditFileName & "]"
                End If
            Else
                If PathExists(AppPath & "Base\Lawn\" & inArg, True) Then
                    EditFileName = inArg
                    EditFileData = ReadFile(AppPath & "Base\Lawn\" & inArg)
                    AddMessage "File loaded [" & AppPath & "Base\Lawn\" & inArg & "]"
                Else
                    AddMessage "File not found [" & AppPath & "Base\Lawn\" & inArg & "]"
                End If
            End If
            
        Case "view"
                If PathExists(AppPath & "Base\Lawn\" & EditFileName, True) Then
                    
                    If (IsNumeric(NextArg(NextArg(inArg, " "), "-")) And IsNumeric(RemoveArg(NextArg(inArg, " "), "-"))) Or IsNumeric(NextArg(inArg, " ")) Then
                        
                        If Not IsNumeric(NextArg(inArg, " ")) Then
                            l = NextArg(NextArg(inArg, " "), "-")
                            o = RemoveArg(NextArg(inArg, " "), "-")
                        Else
                            l = NextArg(inArg, " ")
                            o = l
                        End If
                        If l <= o Then
                            AddMessage "Begin View"
                            inTmp = EditFileData
                            cnt = 1
                            Do Until inTmp = ""
                                If cnt >= l And cnt <= o Then
                                    AddMessage String(3 - Len(Trim(CStr(cnt))), "0") & Trim(CStr(cnt)) & ": " & Replace(RemoveNextArgNoTrim(inTmp, vbCrLf), vbTab, "     ")
                                Else
                                    RemoveNextArg inTmp, vbCrLf
                                End If
                                cnt = cnt + 1
                            Loop
                            AddMessage "End View"
                        Else
                            AddMessage "Invalid line number(s) specified."
                        End If
                    Else
                        AddMessage "Invalid line number(s) specified."
                    End If
                ElseIf EditFileName = "" Then
                    AddMessage "File not loaded."
                Else
                    AddMessage "File not found [" & AppPath & "Base\Lawn\" & EditFileName & "]"
                End If
        Case "lines"
                If PathExists(AppPath & "Base\Lawn\" & EditFileName, True) Then
                    If (IsNumeric(NextArg(NextArg(inArg, " "), "-")) And IsNumeric(RemoveArg(NextArg(inArg, " "), "-"))) Or IsNumeric(NextArg(inArg, " ")) Then
                        
                        If Not IsNumeric(NextArg(inArg, " ")) Then
                            l = NextArg(NextArg(inArg, " "), "-")
                            o = RemoveArg(NextArg(inArg, " "), "-")
                        Else
                            l = NextArg(inArg, " ")
                            o = l
                        End If
                        If l <= o Then
                            inTmp = EditFileData
                            cnt = 1
                            Do Until inTmp = ""
                                If cnt >= l And cnt <= o Then
                                    inNew = inNew & vbCrLf
                                Else
                                    inNew = inNew & RemoveNextArgNoTrim(inTmp, vbCrLf) & vbCrLf
                                End If
                                cnt = cnt + 1
                            Loop
                            EditFileData = inNew
                            AddMessage "Blank line" & IIf(l = o, " ", "s ") & inArg & " added."
                        Else
                            AddMessage "Invalid line number(s) specified."
                        End If
                    Else
                        AddMessage "Invalid line number(s) specified."
                    End If
                ElseIf EditFileName = "" Then
                    AddMessage "File not loaded."
                Else
                    AddMessage "File not found [" & AppPath & "Base\Lawn\" & EditFileName & "]"
                End If
        Case "edit"
                If PathExists(AppPath & "Base\Lawn\" & EditFileName, True) Then
                    If IsNumeric(NextArg(inArg, " ")) Then
                        l = RemoveNextArgNoTrim(inArg, " ")
                        inTmp = EditFileData
                        cnt = 1
                        Do Until inTmp = ""
                            If cnt = l Then
                                inNew = inNew & inArg & vbCrLf
                                RemoveNextArg inTmp, vbCrLf
                            Else
                                inNew = inNew & RemoveNextArgNoTrim(inTmp, vbCrLf) & vbCrLf
                            End If
                            cnt = cnt + 1
                        Loop
                        EditFileData = inNew
                        AddMessage "Edited " & l & ": " & Replace(inArg, vbTab, "     ")
                    Else
                        AddMessage "Invalid line number(s) specified."
                    End If
                ElseIf EditFileName = "" Then
                    AddMessage "File not loaded."
                Else
                    AddMessage "File not found [" & AppPath & "Base\Lawn\" & EditFileName & "]"
                End If
        Case "save"
                If PathExists(AppPath & "Base\Lawn\" & EditFileName, True) Then
                    WriteFile AppPath & "Base\Lawn\" & EditFileName, EditFileData
                    AddMessage "Saved data file [" & AppPath & "Base\Lawn\" & EditFileName & "]"
                ElseIf EditFileName = "" Then
                    AddMessage "File not loaded."
                Else
                    AddMessage "File not found [" & AppPath & "Base\Lawn\" & EditFileName & "]"
                End If
                
                
        Case "nickels"
            If (Not frmMain.Recording) And (Not frmMain.IsPlayback) And (Not frmMain.Multiplayer) Then
                AddMessage "Gathered: " & ScoreNth & ", Total Generated: " & NthTime
                If BeaconCount > 0 Then
                    For o = 1 To BeaconCount
                        If (Left(Beacons(o).Identity, 1) = "n" And Len(Beacons(o).Identity) = 2) And (Not (Beacons(o).Identity = "n0")) Then
                            If Beacons(o).OriginCount > 0 Then
                                For l = 1 To Beacons(o).OriginCount
                                    AddMessage "Special N Location: " & Beacons(o).Origins(l).X & ", " & Beacons(o).Origins(l).Y & ", " & Beacons(o).Origins(l).z
                                Next
                            End If
                        End If
                    Next
                End If
            Else
                AddMessage "Nickels command is currently disabled..."
            End If
        Case "idols"
            If (Not frmMain.Recording) And (Not frmMain.IsPlayback) And (Not frmMain.Multiplayer) Then
                If ObjectCount > 0 Then
                    For o = 1 To ObjectCount
                        If Objects(o).IsIdol And Objects(o).MeshIndex > 0 Then
    
                            AddMessage "Idol: " & Meshes(Objects(o).MeshIndex).FileName & ", Location: " & Objects(o).Origin.X & ", " & Objects(o).Origin.Y & ", " & Objects(o).Origin.z
    
                        End If
                    Next
                End If
            Else
                AddMessage "Idols command is currently disabled..."
            End If
        Case "warp"
        
            If (Not frmMain.Recording) And (Not frmMain.IsPlayback) And (Not frmMain.Multiplayer) Then
                AddMessage "Warping to " & NextArg(inArg, " ") & ", 0, " & RemoveArg(inArg, " ") & "..."
                Player.Object.Origin.X = NextArg(inArg, " ")
                Player.Object.Origin.z = RemoveArg(inArg, " ")
            Else
                AddMessage "Warp command is currently disabled..."
            End If
            
        Case "god"
            If (Not frmMain.Recording) And (Not frmMain.IsPlayback) And (Not frmMain.Multiplayer) Then
                GodMode = Not GodMode
                AddMessage "GodMode = " & GodMode
            Else
                AddMessage "God currently disabled..."
            End If
        Case "quit"
            If frmMain.IsPlayback Or frmMain.Recording Then
                StopFilm
            ElseIf frmMain.Multiplayer Then
                frmMain.Disconnect
            End If
            StopGame = True
        Case "sound"
            If Not DisableSound Then
            
                Select Case LCase(Trim(inArg))
                    Case "-1", "1", "true", "on"
                        PlaySound = True
                    Case "0", "false", "off"
                        PlaySound = False
                End Select
                db.dbQuery "UPDATE Settings SET SoundEnabled = " & IIf(PlaySound, "Yes", "No") & " WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"
                AddMessage "SoundEnabled = " & PlaySound
            Else
                AddMessage "Sound system disabled."
            End If
        Case "music"
            If Not DisableSound Then

                Select Case LCase(Trim(inArg))
                    Case "-1", "1", "true", "on"
                        PlayMusic = True
                        Track1.PlaySound
                        Track2.PlaySound
                    Case "0", "false", "off"
                        PlayMusic = False
                        Track1.StopSound
                        Track2.StopSound
                End Select
                db.dbQuery "UPDATE Settings SET MusicEnabled = " & IIf(PlayMusic, "Yes", "No") & " WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"
                AddMessage "MusicEnabled = " & PlayMusic
            Else
                AddMessage "Sound system disabled."
            End If
        Case "wire"
            Select Case LCase(Trim(inArg))
                Case "-1", "1", "true", "on"
                    WireFrame = True
                Case "0", "false", "off"
                    WireFrame = False
            End Select
            DDevice.SetRenderState D3DRS_FILLMODE, IIf(WireFrame, D3DFILL_WIREFRAME, D3DFILL_SOLID)
            db.dbQuery "UPDATE Settings SET WireFrame = " & IIf(WireFrame, "Yes", "No") & " WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"
            AddMessage "WireFrame = " & WireFrame
            
        Case "gamereset"
            StopFilm
            ClearUserData
            CleanupLawn
            CreateLawn
        Case "echo"
            AddMessage inArg
        Case "films"
            If Not PathExists(AppPath & "Films") Then MkDir AppPath & "Films"
            
            Dim fso As Object
            Set fso = CreateObject("Scripting.FileSystemObject")
            Dim f As Object
            Dim sf As Object
            Set f = fso.getfolder(AppPath & "Films")
            Dim films As String
            
            If f.Files.Count > 0 Then
                For Each sf In f.Files
                    films = films & GetFileTitle(sf.name) & ", "
                Next
                films = Left(films, Len(films) - 2)
                AddMessage "Film Titles: " & films
            Else
                AddMessage "No Films Found In [" & AppPath & "Films]"
            End If
            Set f = Nothing
            Set fso = Nothing
            
        Case "record"
            GodMode = False
            RecordFilm inArg
        Case "stop"
            StopFilm
        Case "play"
            GodMode = False
            PlayFilm inArg
        Case "speak"
            If frmMain.Multiplayer Then
                frmMain.Speak inArg
            Else
                AddMessage "Not Connected."
            End If
        Case "scores"
            If frmMain.Multiplayer Then
                frmMain.Scores
            Else
                AddMessage "Not Connected."
            End If
        Case "webconnect"
            If (Not frmMain.Recording) Then

                GodMode = False
                SaveUserData
                If frmMain.Multiplayer Then frmMain.Disconnect
                AddMessage "Web connecting..."
                frmMain.Connect "www.sosouix.net", inArg
            Else
                AddMessage "Unable to use multiplayer commands during record or playback."
            End If
        Case "connect"
            If (Not frmMain.Recording) Then

                GodMode = False
                SaveUserData
                If frmMain.Multiplayer Then frmMain.Disconnect
                frmMain.Connect NextArg(inArg, " "), RemoveArg(inArg, " ")
            Else
                AddMessage "Unable to use multiplayer commands during record or playback."
            End If
        Case "disconnect"
            If (Not frmMain.Recording) Then
                frmMain.Disconnect
            Else
                AddMessage "Unable to use multiplayer commands during record or playback."
            End If
        Case "database"
            Select Case LCase(inArg)
                Case "reset"
                    ResetDB
                    AddMessage "Database reset complete."
                Case "backup"
                    BackupDB
                    AddMessage "Database backed up to [" & AppPath & "backup.ini" & "]"
                Case "restore"
                    RestoreDB False
                    AddMessage "Database restored from [" & AppPath & "backup.ini" & "]"
                Case "compact"
                    CompactDB
                    AddMessage "Database compact complete."
            End Select
            
        Case "partner"
            If (Not frmMain.Recording) And (Not frmMain.IsPlayback) And (Not frmMain.Multiplayer) Then
                AddMessage "Partner location " & Partner.Object.Origin.X & ", " & Partner.Object.Origin.Y & ", " & Partner.Object.Origin.z
                AddMessage "Following " & Clocker.FollowingMode & ", Claps " & Clocker.WantsToDance & ", Timer " & Clocker.NooneIsDancing
            Else
                AddMessage "Partner command is currently disabled..."
            End If
            
        Case "help", "cmdlist", "?", "--?"
            Select Case LCase(inArg)
                Case "commands"
                    AddMessage ""
                    AddMessage "Console Commands:"
                    AddMessage "   QUIT (completely close out of the game)"
                    AddMessage "   MUSIC [<ON/OFF>] (get or set whether music is enabled)"
                    AddMessage "   SOUND [<ON/OFF>] (get or set whether sound is enabled)"
                    AddMessage "   WIRE [<ON/OFF>] (get or set whether the lawn is rendered in wireframe)"
                    AddMessage "   GAMERESET (resets your game play to the beginning clearing all scores)"
                    AddMessage "   DATABASE <[BACKUP/RESTORE/RESET/COMPACT]> (data maintenance commands)"
                    AddMessage "   FILMS (lists the titles of the saved recordings existing in the folder)"
                    AddMessage "   PLAY [<file>] (saves to or plays from recorded data in a file by title)"
                    AddMessage "   RECORD [<file>] (begins game recording everything the player is doing)"
                    AddMessage "   STOP (Stops the game from recording or stops the game from playback)"
                    AddMessage "   CONNECT <host> <name> (connect as player <name> to a multiplayer <host>)"
                    AddMessage "   DISCONNECT (disconnects from the active server and return to solo play)"
                    AddMessage "   WEBCONNECT <name> [<port>] (Connect to neotext.org's online game server)"
                    AddMessage "   SCORES (lists all the players on the server and their score time lines)"
                    AddMessage "   SPEAK (sends a message to everybody that is connected to the server)"
                Case "cheats"
                    AddMessage ""
                    AddMessage "Cheat Commands:"
                    AddMessage "   GOD (Enables god mod which displays your distance, x, y, z and fps)"
                    AddMessage "   WARP <x> <z> (Warps to a position anywhere by a X and Z coordinate)"
                    AddMessage "   NICKELS (Displays how many nickels generated and dispose locations)"
                    AddMessage "   IDOLS (Displays all objective idols their names and their locations)"
                    AddMessage "   STAT (Displays distance, location, and camera values in the console)"
                    AddMessage "   PARTNER (Displays information on a small non playing character bot)"
                Case "editing"
                    AddMessage ""
                    AddMessage "Editing Commands:"
                    AddMessage "   REFRESH (This command will cause the engine to reload the game files)"
                    AddMessage "   LOAD <file> (This command loads the specified file to be edited live)"
                    AddMessage "   VIEW <#[-#]> (This command displays the specified line numbers text)"
                    AddMessage "   LINES <#[-#]> (This command adds lines # to the current loaded file)"
                    AddMessage "   EDIT <#> <text> (This command changes the specified line numbers text)"
                    AddMessage "   SAVE (This command saves any edited changes made with a loaded file)"
                    AddMessage "   (For information on object and identifier editing see documentation)"
                Case Else
                    AddMessage ""
                    AddMessage "For detail help please type one of the following help commands:"
                    AddMessage "   HELP COMMANDS (Displays the help of basic console commands)"
                    AddMessage "   HELP EDITING (Displays the help of editing files in console)"
                    AddMessage "   HELP CHEATS (Displays the help of cheat commands to the game)"
            End Select

        Case Else
            AddMessage "Unknown command."
    End Select
    
    Exit Sub
processerr:
    AddMessage "Error in command [" & inCmd & "]"
    Err.Clear
End Sub

