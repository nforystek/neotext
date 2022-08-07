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
Private TogglePress3 As Long

Private ToggleMouse1 As Long
Private ToggleMouse2 As Long

Private IdleInput As Single

Private Enum ToggleIdents
    MoveAuto = 0
    MoveJump = 1
    MoveForward = 2
    MoveBackward = 3
    MoveStepLeft = 4
    MoveStepRight = 5
End Enum

Private ToggleMotion(0 To 5) As Long

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

Private EditFileName As String
Private EditFileData As String


Private State As Integer
Private Shift As Integer
Private Bottom As Long

Public DrawCount As Long
Public Draws() As Variant


Public JumpGUID As String

Public Bindings(0 To 255) As String

'Public Enum SurfaceControl
'    Forward = 1
'    Backward = 2
'    LeftStep = 4
'    RightStep = 8
'    MouseUp = 16
'    MouseDown = 32
'    MouseLeft = 64
'    MouseRight = 128
'    Jump = 256
'End Enum


Public Function GetBindingIndex(ByVal BindText As String) As Integer
    Select Case Trim(UCase(BindText))
        Case "0"
            GetBindingIndex = DIK_0
        Case "1"
            GetBindingIndex = DIK_1
        Case "2"
            GetBindingIndex = DIK_2
        Case "3"
            GetBindingIndex = DIK_3
        Case "4"
            GetBindingIndex = DIK_4
        Case "5"
            GetBindingIndex = DIK_5
        Case "6"
            GetBindingIndex = DIK_6
        Case "7"
            GetBindingIndex = DIK_7
        Case "8"
            GetBindingIndex = DIK_8
        Case "9"
            GetBindingIndex = DIK_9
        Case "A"
            GetBindingIndex = DIK_A
        Case "ABNT_C1"
            GetBindingIndex = DIK_ABNT_C1
        Case "ABNT_C2"
            GetBindingIndex = DIK_ABNT_C2
        Case "ADD"
            GetBindingIndex = DIK_ADD
        Case "APOSTROPHE"
            GetBindingIndex = DIK_APOSTROPHE
        Case "APPS"
            GetBindingIndex = DIK_APPS
        Case "AT"
            GetBindingIndex = DIK_AT
        Case "AX"
            GetBindingIndex = DIK_AX
        Case "B"
            GetBindingIndex = DIK_B
        Case "BACK"
            GetBindingIndex = DIK_BACK
        Case "BACKSLASH"
            GetBindingIndex = DIK_BACKSLASH
        Case "BACKSPACE"
            GetBindingIndex = DIK_BACKSPACE
        Case "C"
            GetBindingIndex = DIK_C
        Case "CALCULATOR"
            GetBindingIndex = DIK_CALCULATOR
        Case "CAPITAL"
            GetBindingIndex = DIK_CAPITAL
        Case "CAPSLOCK"
            GetBindingIndex = DIK_CAPSLOCK
        Case "CIRCUMFLEX"
            GetBindingIndex = DIK_CIRCUMFLEX
        Case "COLON"
            GetBindingIndex = DIK_COLON
        Case "COMMA"
            GetBindingIndex = DIK_COMMA
        Case "CONVERT"
            GetBindingIndex = DIK_CONVERT
        Case "D"
            GetBindingIndex = DIK_D
        Case "DECIMAL"
            GetBindingIndex = DIK_DECIMAL
        Case "DELETE"
            GetBindingIndex = DIK_DELETE
        Case "DIVIDE"
            GetBindingIndex = DIK_DIVIDE
        Case "DOWN"
            GetBindingIndex = DIK_DOWN
        Case "DOWNARROW"
            GetBindingIndex = DIK_DOWNARROW
        Case "E"
            GetBindingIndex = DIK_E
        Case "END"
            GetBindingIndex = DIK_END
        Case "EQUALS"
            GetBindingIndex = DIK_EQUALS
        Case "ESCAPE"
            GetBindingIndex = DIK_ESCAPE
        Case "F"
            GetBindingIndex = DIK_F
        Case "F1"
            GetBindingIndex = DIK_F1
        Case "F10"
            GetBindingIndex = DIK_F10
        Case "F11"
            GetBindingIndex = DIK_F11
        Case "F12"
            GetBindingIndex = DIK_F12
        Case "F13"
            GetBindingIndex = DIK_F13
        Case "F14"
            GetBindingIndex = DIK_F14
        Case "F15"
            GetBindingIndex = DIK_F15
        Case "F2"
            GetBindingIndex = DIK_F2
        Case "F3"
            GetBindingIndex = DIK_F3
        Case "F4"
            GetBindingIndex = DIK_F4
        Case "F5"
            GetBindingIndex = DIK_F5
        Case "F6"
            GetBindingIndex = DIK_F6
        Case "F7"
            GetBindingIndex = DIK_F7
        Case "F8"
            GetBindingIndex = DIK_F8
        Case "F9"
            GetBindingIndex = DIK_F9
        Case "G"
            GetBindingIndex = DIK_G
        Case "GRAVE"
            GetBindingIndex = DIK_GRAVE
        Case "H"
            GetBindingIndex = DIK_H
        Case "HOME"
            GetBindingIndex = DIK_HOME
        Case "I"
            GetBindingIndex = DIK_I
        Case "INSERT"
            GetBindingIndex = DIK_INSERT
        Case "J"
            GetBindingIndex = DIK_J
        Case "K"
            GetBindingIndex = DIK_K
        Case "KANA"
            GetBindingIndex = DIK_KANA
        Case "KANJI"
            GetBindingIndex = DIK_KANJI
        Case "L"
            GetBindingIndex = DIK_L
        Case "LALT"
            GetBindingIndex = DIK_LALT
        Case "LBRACKET"
            GetBindingIndex = DIK_LBRACKET
        Case "LCONTROL"
            GetBindingIndex = DIK_LCONTROL
        Case "LEFT"
            GetBindingIndex = DIK_LEFT
        Case "LEFTARROW"
            GetBindingIndex = DIK_LEFTARROW
        Case "LMENU"
            GetBindingIndex = DIK_LMENU
        Case "LSHIFT"
            GetBindingIndex = DIK_LSHIFT
        Case "LWIN"
            GetBindingIndex = DIK_LWIN
        Case "M"
            GetBindingIndex = DIK_M
        Case "MAIL"
            GetBindingIndex = DIK_MAIL
        Case "MEDIASELECT"
            GetBindingIndex = DIK_MEDIASELECT
        Case "MEDIASTOP"
            GetBindingIndex = DIK_MEDIASTOP
        Case "MINUS"
            GetBindingIndex = DIK_MINUS
        Case "MULTIPLY"
            GetBindingIndex = DIK_MULTIPLY
        Case "MUTE"
            GetBindingIndex = DIK_MUTE
        Case "MYCOMPUTER"
            GetBindingIndex = DIK_MYCOMPUTER
        Case "N"
            GetBindingIndex = DIK_N
        Case "NEXT"
            GetBindingIndex = DIK_NEXT
        Case "NEXTTRACK"
            GetBindingIndex = DIK_NEXTTRACK
        Case "NOCONVERT"
            GetBindingIndex = DIK_NOCONVERT
        Case "NUMLOCK"
            GetBindingIndex = DIK_NUMLOCK
        Case "NUMPAD0"
            GetBindingIndex = DIK_NUMPAD0
        Case "NUMPAD1"
            GetBindingIndex = DIK_NUMPAD1
        Case "NUMPAD2"
            GetBindingIndex = DIK_NUMPAD2
        Case "NUMPAD3"
            GetBindingIndex = DIK_NUMPAD3
        Case "NUMPAD4"
            GetBindingIndex = DIK_NUMPAD4
        Case "NUMPAD5"
            GetBindingIndex = DIK_NUMPAD5
        Case "NUMPAD6"
            GetBindingIndex = DIK_NUMPAD6
        Case "NUMPAD7"
            GetBindingIndex = DIK_NUMPAD7
        Case "NUMPAD8"
            GetBindingIndex = DIK_NUMPAD8
        Case "NUMPAD9"
            GetBindingIndex = DIK_NUMPAD9
        Case "NUMPADCOMMA"
            GetBindingIndex = DIK_NUMPADCOMMA
        Case "NUMPADENTER"
            GetBindingIndex = DIK_NUMPADENTER
        Case "NUMPADEQUALS"
            GetBindingIndex = DIK_NUMPADEQUALS
        Case "NUMPADMINUS"
            GetBindingIndex = DIK_NUMPADMINUS
        Case "NUMPADPERIOD"
            GetBindingIndex = DIK_NUMPADPERIOD
        Case "NUMPADPLUS"
            GetBindingIndex = DIK_NUMPADPLUS
        Case "NUMPADSLASH"
            GetBindingIndex = DIK_NUMPADSLASH
        Case "NUMPADSTAR"
            GetBindingIndex = DIK_NUMPADSTAR
        Case "O"
            GetBindingIndex = DIK_O
        Case "OEM_102"
            GetBindingIndex = DIK_OEM_102
        Case "P"
            GetBindingIndex = DIK_P
        Case "PAUSE"
            GetBindingIndex = DIK_PAUSE
        Case "PERIOD"
            GetBindingIndex = DIK_PERIOD
        Case "PGDN"
            GetBindingIndex = DIK_PGDN
        Case "PGUP"
            GetBindingIndex = DIK_PGUP
        Case "PLAYPAUSE"
            GetBindingIndex = DIK_PLAYPAUSE
        Case "POWER"
            GetBindingIndex = DIK_POWER
        Case "PREVTRACK"
            GetBindingIndex = DIK_PREVTRACK
        Case "PRIOR"
            GetBindingIndex = DIK_PRIOR
        Case "Q"
            GetBindingIndex = DIK_Q
        Case "R"
            GetBindingIndex = DIK_R
        Case "RALT"
            GetBindingIndex = DIK_RALT
        Case "RBRACKET"
            GetBindingIndex = DIK_RBRACKET
        Case "RCONTROL"
            GetBindingIndex = DIK_RCONTROL
        Case "RETURN"
            GetBindingIndex = DIK_RETURN
        Case "RIGHT"
            GetBindingIndex = DIK_RIGHT
        Case "RIGHTARROW"
            GetBindingIndex = DIK_RIGHTARROW
        Case "RMENU"
            GetBindingIndex = DIK_RMENU
        Case "RSHIFT"
            GetBindingIndex = DIK_RSHIFT
        Case "RWIN"
            GetBindingIndex = DIK_RWIN
        Case "S"
            GetBindingIndex = DIK_S
        Case "SCROLL"
            GetBindingIndex = DIK_SCROLL
        Case "SEMICOLON"
            GetBindingIndex = DIK_SEMICOLON
        Case "SLASH"
            GetBindingIndex = DIK_SLASH
        Case "SLEEP"
            GetBindingIndex = DIK_SLEEP
        Case "STOP"
            GetBindingIndex = DIK_STOP
        Case "SUBTRACT"
            GetBindingIndex = DIK_SUBTRACT
        Case "SYSRQ"
            GetBindingIndex = DIK_SYSRQ
        Case "T"
            GetBindingIndex = DIK_T
        Case "TAB"
            GetBindingIndex = DIK_TAB
        Case "U"
            GetBindingIndex = DIK_U
        Case "UNDERLINE"
            GetBindingIndex = DIK_UNDERLINE
        Case "UNLABELED"
            GetBindingIndex = DIK_UNLABELED
        Case "UP"
            GetBindingIndex = DIK_UP
        Case "UPARROW"
            GetBindingIndex = DIK_UPARROW
        Case "V"
            GetBindingIndex = DIK_V
        Case "VOLUMEDOWN"
            GetBindingIndex = DIK_VOLUMEDOWN
        Case "VOLUMEUP"
            GetBindingIndex = DIK_VOLUMEUP
        Case "W"
            GetBindingIndex = DIK_W
        Case "WAKE"
            GetBindingIndex = DIK_WAKE
        Case "WEBBACK"
            GetBindingIndex = DIK_WEBBACK
        Case "WEBFAVORITES"
            GetBindingIndex = DIK_WEBFAVORITES
        Case "WEBFORWARD"
            GetBindingIndex = DIK_WEBFORWARD
        Case "WEBHOME"
            GetBindingIndex = DIK_WEBHOME
        Case "WEBREFRESH"
            GetBindingIndex = DIK_WEBREFRESH
        Case "WEBSEARCH"
            GetBindingIndex = DIK_WEBSEARCH
        Case "WEBSTOP"
            GetBindingIndex = DIK_WEBSTOP
        Case "X"
            GetBindingIndex = DIK_X
        Case "Y"
            GetBindingIndex = DIK_Y
        Case "YEN"
            GetBindingIndex = DIK_YEN
        Case "Z"
            GetBindingIndex = DIK_Z
        Case Else
            GetBindingIndex = -1
    End Select
End Function
Public Sub ResetIdle()
    IdleInput = Timer
End Sub
Public Function CheckIdle(ByVal Seconds As Single) As Boolean
    CheckIdle = ((Timer - IdleInput) >= Seconds)
End Function

Public Sub PrintText(ByVal inTxt As String, ByVal inX As Single, ByVal inY As Single)
  
    DrawCount = DrawCount + 1
    ReDim Preserve Draws(1 To 3, 1 To DrawCount) As Variant
    Draws(1, DrawCount) = inTxt
    Draws(2, DrawCount) = inX
    Draws(3, DrawCount) = inY
End Sub

Public Sub ClearText()
    If DrawCount > 0 Then
        DrawCount = 0
        Erase Draws
    End If
End Sub

Public Property Get TextHeight() As Single
    TextHeight = frmMain.TextHeight("A")
End Property

Public Property Get TextSpace() As Single
    TextSpace = 2
End Property

'Public Function SurfaceCatch(ByVal X As Single, ByVal Y As Single) As Boolean
'    Static ToggleJump As Boolean
'    Dim vecDirect As D3DVECTOR
'    Dim moveFriction As Single
'    moveFriction = 0.05
'    SurfaceCatch = True
'    Select Case SurfaceHit(X / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY)
'        Case MouseUp
'            MouseLook 0, -20, 0
'            ResetIdle
'            ToggleJump = False
'        Case MouseDown
'            MouseLook 0, 20, 0
'            ResetIdle
'            ToggleJump = False
'        Case MouseLeft
'            MouseLook -20, 0, 0
'            ResetIdle
'            ToggleJump = False
'        Case MouseRight
'            MouseLook 20, 0, 0
'            ResetIdle
'            ToggleJump = False
'        Case Forward
'            If Player.Object.Direct.Y = 0 Then
'                vecDirect.X = Sin(D720 - Player.CameraAngle)
'                vecDirect.z = Cos(D720 - Player.CameraAngle)
'                If ((Perspective = Spectator) Or DebugMode) Or Player.Object.States.InLiquid Then
'                    vecDirect.Y = -(Tan(D720 - Player.CameraPitch))
'                End If
'                D3DXVec3Normalize vecDirect, vecDirect
'                If Player.Object.States.InLiquid And (Not ((Perspective = Spectator) Or DebugMode)) Then
'                    AddActivity Player.Object, Actions.Directing, Replace(modGuid.GUID, "-", ""), vecDirect, (Player.MoveSpeed / 2), moveFriction
'                Else
'                    AddActivity Player.Object, Actions.Directing, Replace(modGuid.GUID, "-", ""), vecDirect, Player.MoveSpeed, moveFriction
'                End If
'            End If
'            ResetIdle
'            ToggleJump = False
'        Case Backward
'
'            If Player.Object.Direct.Y = 0 Then
'                vecDirect.X = -Sin(D720 - Player.CameraAngle)
'                vecDirect.z = -Cos(D720 - Player.CameraAngle)
'                If (Perspective = Spectator) Or DebugMode Then
'                    vecDirect.Y = Tan(D720 - Player.CameraPitch)
'                End If
'                D3DXVec3Normalize vecDirect, vecDirect
'                If Player.Object.States.InLiquid And (Not ((Perspective = Spectator) Or DebugMode)) Then
'                    AddActivity Player.Object, Actions.Directing, Replace(modGuid.GUID, "-", ""), vecDirect, (Player.MoveSpeed / 2), moveFriction
'                Else
'                    AddActivity Player.Object, Actions.Directing, Replace(modGuid.GUID, "-", ""), vecDirect, Player.MoveSpeed, moveFriction
'                End If
'            End If
'            ResetIdle
'            ToggleJump = False
'        Case LeftStep
'            If Player.Object.Direct.Y = 0 Then
'                vecDirect.X = Sin((D720 - Player.CameraAngle) - D180)
'                vecDirect.z = Cos((D720 - Player.CameraAngle) - D180)
'                D3DXVec3Normalize vecDirect, vecDirect
'                If Player.Object.States.InLiquid And (Not ((Perspective = Spectator) Or DebugMode)) Then
'                    AddActivity Player.Object, Actions.Directing, Replace(modGuid.GUID, "-", ""), vecDirect, (Player.MoveSpeed / 2), moveFriction
'                Else
'                    AddActivity Player.Object, Actions.Directing, Replace(modGuid.GUID, "-", ""), vecDirect, Player.MoveSpeed, moveFriction
'                End If
'            End If
'            ResetIdle
'            ToggleJump = False
'        Case RightStep
'            If Player.Object.Direct.Y = 0 Then
'                vecDirect.X = Sin((D720 - Player.CameraAngle) + D180)
'                vecDirect.z = Cos((D720 - Player.CameraAngle) + D180)
'                D3DXVec3Normalize vecDirect, vecDirect
'                If Player.Object.States.InLiquid And (Not ((Perspective = Spectator) Or DebugMode)) Then
'                    AddActivity Player.Object, Actions.Directing, Replace(modGuid.GUID, "-", ""), vecDirect, (Player.MoveSpeed / 2), moveFriction
'                Else
'                    AddActivity Player.Object, Actions.Directing, Replace(modGuid.GUID, "-", ""), vecDirect, Player.MoveSpeed, moveFriction
'                End If
'            End If
'            ResetIdle
'            ToggleJump = False
'        Case Jump
'
'            If Not ToggleJump Then
'                ToggleJump = True
'
'                If (Perspective = Spectator) Or DebugMode Then
'                    vecDirect.Y = vecDirect.Y + IIf(Player.MoveSpeed < 1, 1, Player.MoveSpeed)
'                    AddActivity Player.Object, Actions.Directing, Replace(modGuid.GUID, "-", ""), vecDirect, Player.MoveSpeed, moveFriction
'                Else
'
'                    If ActivityExists(Player.Object, JumpGUID) Then
'                        If (Not ((Player.Object.States.IsMoving And Moving.Flying) = Moving.Flying) Or _
'                                ((Player.Object.States.IsMoving And Moving.Falling) = Moving.Falling)) Then
'                            'If ActivityExists(Player.Object, JumpGUID) Then
'                                Do Until Not ActivityExists(Player.Object, JumpGUID)
'                                    DeleteActivity Player.Object, JumpGUID
'                                Loop
'                            'End If
'                        End If
'                    End If
'                    If Not ActivityExists(Player.Object, JumpGUID) Then
'                        vecDirect.Y = IIf(Player.Object.States.InLiquid, 5, 9)
'                        JumpGUID = AddActivity(Player.Object, Actions.Directing, JumpGUID, vecDirect, (Player.MoveSpeed * 4), moveFriction)
'                    End If
'                End If
'            Else
'                ToggleJump = False
'            End If
'            ResetIdle
'
'        Case Else
'            ToggleJump = False
'            SurfaceCatch = False
'    End Select
'End Function

Private Function Toggled(ByVal vkCode As Long) As Boolean
    Toggled = KeyState(vkCode).VKToggle
End Function
Private Function Pressed(ByVal vkCode As Long) As Boolean
    Pressed = KeyState(vkCode).VKPressed
End Function

Public Sub InputScene()
        
    Dim cnt As Long
    Dim cnt2 As Long
    Dim hit As Long
    
    On Error GoTo pausing
   
    DIKeyBoardDevice.GetDeviceStateKeyboard DIKEYBOARDSTATE

    If (GetActiveWindow = frmMain.hwnd) Then
        If FullScreen And Not TrapMouse Then TrapMouse = True
        
        ConsoleInput DIKEYBOARDSTATE
        
        Dim uses(0 To 255) As Boolean
        For cnt = 0 To 255
            If DIKEYBOARDSTATE.Key(cnt) Then
                If Not Bindings(cnt) = "" Then
                    uses(cnt) = True
    
                    ParseLand 0, Bindings(cnt)
    
                End If
            End If
        Next
        
        If DIKEYBOARDSTATE.Key(DIK_F2) And (Not uses(DIK_F2)) Then
            If (Not TogglePress1 = DIK_F2) Then
                TogglePress1 = DIK_F2
                
                '############### SHOW StATS ################
                ResetIdle
                ShowStat = (Not ShowStat) Or (ShowStat And Not ShowHelp)
                If ShowStat Then ShowHelp = True
                
            End If
        ElseIf DIKEYBOARDSTATE.Key(DIK_F1) And (Not uses(DIK_F1)) Then
            If (Not TogglePress1 = DIK_F1) Then
                TogglePress1 = DIK_F1
                
                '############### SHOW HELP ################
                ResetIdle
                ShowHelp = Not ShowHelp
                
            End If
        ElseIf DIKEYBOARDSTATE.Key(DIK_F3) And (Not uses(DIK_F3)) Then
            If (Not TogglePress1 = DIK_F3) Then
                TogglePress1 = DIK_F3
                
                '############### SHOW CrEDIts ################
                ResetIdle
                ShowCredits = Not ShowCredits
                
            End If
    
    '    ElseIf DIKEYBOARDSTATE.Key(DIK_F5) And (Not uses(DIK_F5)) Then
    '        If DebugMode Then
    '            If (Not TogglePress1 = DIK_F5) Then
    '                TogglePress1 = DIK_F5
    '
    '                '###############  ################
    '                ResetIdle
    '                CullingObject.Position = Player.Object.Origin
    '                CullingObject.Direction = VectorNormalize(VectorSubtract(MakeVector(Player.Object.Origin.X + (Sin(D720 - Player.CameraAngle) * 1), _
    '                                                                    Player.Object.Origin.Y - (Tan(D720 - Player.CameraPitch) * 1), _
    '                                                                Player.Object.Origin.z + (Cos(D720 - Player.CameraAngle) * 1)), Player.Object.Origin))
    '                CullingObject.UpVector = VectorNormalize(VectorSubtract(MakeVector(Player.Object.Origin.X + (Sin(D720 - Player.CameraAngle) * 1), _
    '                                                                    Player.Object.Origin.Y - (Tan(D720 - Player.CameraPitch) * 1), _
    '                                                                Player.Object.Origin.z + (Cos(D720 - Player.CameraAngle) * 1)), Player.Object.Origin))
    '                CullingSetup = 1
    '
    '            End If
    '        End If
    '    ElseIf DIKEYBOARDSTATE.Key(DIK_F6) And (Not uses(DIK_F6)) Then
    '        If DebugMode Then
    '            If (Not TogglePress1 = DIK_F6) Then
    '                TogglePress1 = DIK_F6
    '
    '                '###############  ################
    '                ResetIdle
    '                CullingSetup = 2
    '
    '            End If
    '        End If
    '    ElseIf DIKEYBOARDSTATE.Key(DIK_F7) And (Not uses(DIK_F7)) Then
    '        If DebugMode Then
    '            If (Not TogglePress1 = DIK_F7) Then
    '                TogglePress1 = DIK_F7
    '
    '                '###############  ################
    '                ResetIdle
    '                CullingCount = CullingCount + 1
    '                ReDim Preserve Cullings(1 To CullingCount) As MyCulling
    '                Cullings(CullingCount) = CullingObject
    '                CullingSetup = 0
    '
    '            End If
    '        End If
    '    ElseIf DIKEYBOARDSTATE.Key(DIK_F8) And (Not uses(DIK_F8)) Then
    '        If DebugMode Then
    '            If (Not TogglePress1 = DIK_F8) Then
    '                TogglePress1 = DIK_F8
    '
    '                '###############  ################
    '                ResetIdle
    '                If CullingCount > 0 Then
    '                    CullingCount = 0
    '                    Erase Cullings
    '                End If
    '
    '            End If
    '        End If
        ElseIf DIKEYBOARDSTATE.Key(DIK_ESCAPE) And (Not uses(DIK_ESCAPE)) Then
            If (Not TogglePress1 = DIK_ESCAPE) Then
                TogglePress1 = DIK_ESCAPE
                
                '###############  ################
                ResetIdle
                If ((Not FullScreen) And TrapMouse) Or ConsoleVisible Then
                    TrapMouse = False
                ElseIf FullScreen Or (Not FullScreen And Not TrapMouse) Then
                    StopGame = True
                End If
                
            End If
        ElseIf DIKEYBOARDSTATE.Key(DIK_GRAVE) And (Not uses(DIK_GRAVE)) Then
            If (Not TogglePress1 = DIK_GRAVE) Then
                TogglePress1 = DIK_GRAVE
                
                '###############  ################
                ResetIdle
                ConsoleToggle
                
            End If
    '    ElseIf DIKEYBOARDSTATE.Key(DIK_LALT) And (Not uses(DIK_LALT)) Then
    '        If (Not TogglePress1 = DIK_LALT) Then
    '            TogglePress1 = DIK_LALT
    '
    '            '###############  ################
    '            ResetIdle
    '            If DIKEYBOARDSTATE.Key(DIK_TAB) Then
    '
    '
    '                DoPauseGame
    '                frmMain.WindowState = 1
    '            End If
    '
    '        End If
    '    ElseIf DIKEYBOARDSTATE.Key(DIK_RALT) And (Not uses(DIK_RALT)) Then
    '        If (Not TogglePress1 = DIK_RALT) Then
    '            TogglePress1 = DIK_RALT
    '
    '            '###############  ################
    '            ResetIdle
    '            If DIKEYBOARDSTATE.Key(DIK_TAB) Then
    '
    '                DoPauseGame
    '                frmMain.WindowState = 1
    '
    '            End If
    '
    '        End If
            ElseIf DIKEYBOARDSTATE.Key(DIK_LALT) Or DIKEYBOARDSTATE.Key(DIK_RALT) Then
                If DIKEYBOARDSTATE.Key(DIK_TAB) Then
                    If (Not TogglePress1 = DIK_TAB) Then
                        TogglePress1 = DIK_TAB
                        TrapMouse = False
                        frmMain.WindowState = 1
                    End If
                End If
                
    '    ElseIf ((DIKEYBOARDSTATE.Key(DIK_RALT) Or DIKEYBOARDSTATE.Key(DIK_LALT)) And DIKEYBOARDSTATE.Key(DIK_TAB)) Then
    '
    '        If (Not TogglePress1 = DIK_LALT + DIK_RALT + DIK_TAB) Then
    '            TogglePress1 = DIK_LALT + DIK_RALT + DIK_TAB
    '
    '            '###############  ################
    '            ResetIdle
    '
    '            If (GetActiveWindow = frmMain.hwnd) And TrapMouse Then
    '                TrapMouse = False
    '                frmMain.WindowState = 1
    '            End If
    '
    '        End If
            
        ElseIf Not (TogglePress1 = 0) Then
            TogglePress1 = 0
        End If
    
    End If
    
       
    If (Not ConsoleVisible) And TrapMouse Then
        
        If (DIKEYBOARDSTATE.Key(DIK_TAB)) And (Not (DIKEYBOARDSTATE.Key(DIK_LALT) Or DIKEYBOARDSTATE.Key(DIK_RALT))) Then
            If (Not TogglePress3 = DIK_TAB) Then
                TogglePress3 = DIK_TAB
                
                '###############  ################
                ResetIdle
                If Perspective = Playmode.ThirdPerson Then
                    Perspective = Playmode.FirstPerson
                ElseIf Perspective = Playmode.FirstPerson Then
                    Perspective = IIf((CameraCount > 0), Playmode.CameraMode, Playmode.ThirdPerson)
                ElseIf Perspective = Playmode.CameraMode Then
                    Perspective = Playmode.ThirdPerson
                End If
                
            End If
        ElseIf Not (TogglePress3 = 0) Then
            TogglePress3 = 0
        End If
        
        If Player.MoveSpeed > MaxDisplacement Then Player.MoveSpeed = MaxDisplacement
        If Player.MoveSpeed < 0.01 Then Player.MoveSpeed = 0.01
        
        Dim vecDirect As D3DVECTOR
        Dim moveFriction As Single
        moveFriction = 0.05
        
        If ((Perspective = Spectator) Or DebugMode) Or _
            ((((Not ((Player.Object.States.IsMoving And Moving.Falling) = Moving.Falling))) _
            Or Player.Object.States.InLiquid) And Player.Object.Visible) Then
                     
            If (DIKEYBOARDSTATE.Key(DIK_E)) And (Not uses(DIK_E)) Then
            
                '###############  ################
                ResetIdle
                If Player.Object.Direct.Y = 0 Then
                    vecDirect.X = Sin(D720 - Player.CameraAngle)
                    vecDirect.z = Cos(D720 - Player.CameraAngle)
                    If ((Perspective = Spectator) Or DebugMode) Or Player.Object.States.InLiquid Then
                        vecDirect.Y = -(Tan(D720 - Player.CameraPitch))
                    End If
                    D3DXVec3Normalize vecDirect, vecDirect
                    If Player.Object.States.InLiquid And (Not ((Perspective = Spectator) Or DebugMode)) Then
                        AddActivity Player.Object, Actions.Directing, Replace(modGuid.GUID, "-", ""), vecDirect, (Player.MoveSpeed / 2), moveFriction
                    Else
                        AddActivity Player.Object, Actions.Directing, Replace(modGuid.GUID, "-", ""), vecDirect, Player.MoveSpeed, moveFriction
                    End If
                End If
                
            End If
            If (DIKEYBOARDSTATE.Key(DIK_D)) And (Not uses(DIK_D)) Then
            
                '###############  ################
                ResetIdle
                If Player.Object.Direct.Y = 0 Then
                    vecDirect.X = -Sin(D720 - Player.CameraAngle)
                    vecDirect.z = -Cos(D720 - Player.CameraAngle)
                    If (Perspective = Spectator) Or DebugMode Then
                        vecDirect.Y = Tan(D720 - Player.CameraPitch)
                    End If
                    D3DXVec3Normalize vecDirect, vecDirect
                    If Player.Object.States.InLiquid And (Not ((Perspective = Spectator) Or DebugMode)) Then
                        AddActivity Player.Object, Actions.Directing, Replace(modGuid.GUID, "-", ""), vecDirect, (Player.MoveSpeed / 2), moveFriction
                    Else
                        AddActivity Player.Object, Actions.Directing, Replace(modGuid.GUID, "-", ""), vecDirect, Player.MoveSpeed, moveFriction
                    End If
                End If
                
            End If
            If (DIKEYBOARDSTATE.Key(DIK_W)) And (Not uses(DIK_W)) Then
            
                '###############  ################
                ResetIdle
                If Player.Object.Direct.Y = 0 Then
                    vecDirect.X = Sin((D720 - Player.CameraAngle) - D180)
                    vecDirect.z = Cos((D720 - Player.CameraAngle) - D180)
                    D3DXVec3Normalize vecDirect, vecDirect
                    If Player.Object.States.InLiquid And (Not ((Perspective = Spectator) Or DebugMode)) Then
                        AddActivity Player.Object, Actions.Directing, Replace(modGuid.GUID, "-", ""), vecDirect, (Player.MoveSpeed / 2), moveFriction
                    Else
                        AddActivity Player.Object, Actions.Directing, Replace(modGuid.GUID, "-", ""), vecDirect, Player.MoveSpeed, moveFriction
                    End If
                End If
                
            End If
            If (DIKEYBOARDSTATE.Key(DIK_R)) And (Not uses(DIK_R)) Then
            
                '###############  ################
                ResetIdle
                If Player.Object.Direct.Y = 0 Then
                    vecDirect.X = Sin((D720 - Player.CameraAngle) + D180)
                    vecDirect.z = Cos((D720 - Player.CameraAngle) + D180)
                    D3DXVec3Normalize vecDirect, vecDirect
                    If Player.Object.States.InLiquid And (Not ((Perspective = Spectator) Or DebugMode)) Then
                        AddActivity Player.Object, Actions.Directing, Replace(modGuid.GUID, "-", ""), vecDirect, (Player.MoveSpeed / 2), moveFriction
                    Else
                        AddActivity Player.Object, Actions.Directing, Replace(modGuid.GUID, "-", ""), vecDirect, Player.MoveSpeed, moveFriction
                    End If
                End If
                
            End If
        End If
        If ((Perspective = Spectator) Or DebugMode) Or (((Not ((Player.Object.States.IsMoving And Moving.Flying) = Moving.Flying))) And _
                                        (Not ((Player.Object.States.IsMoving And Moving.Falling) = Moving.Falling))) Then
            
            If (DIKEYBOARDSTATE.Key(DIK_SPACE)) And (Not uses(DIK_SPACE)) Then
                If (ToggleMotion(ToggleIdents.MoveJump) <> DIK_SPACE) Then
                    ToggleMotion(ToggleIdents.MoveJump) = DIK_SPACE
                    
                    '###############  ################
                    ResetIdle
                    If (Perspective = Spectator) Or DebugMode Then
                        vecDirect.Y = vecDirect.Y + IIf(Player.MoveSpeed < 1, 1, Player.MoveSpeed)
                        AddActivity Player.Object, Actions.Directing, Replace(modGuid.GUID, "-", ""), vecDirect, Player.MoveSpeed, moveFriction
                    Else
                    
                        If ActivityExists(Player.Object, JumpGUID) Then
                            If (Not ((Player.Object.States.IsMoving And Moving.Flying) = Moving.Flying) Or _
                                    ((Player.Object.States.IsMoving And Moving.Falling) = Moving.Falling)) Then
                                Do Until Not ActivityExists(Player.Object, JumpGUID)
                                    DeleteActivity Player.Object, JumpGUID
                                Loop
                            End If
                        End If
                        If Not ActivityExists(Player.Object, JumpGUID) Then
                            vecDirect.Y = IIf(Player.Object.States.InLiquid, 5, 9)
                            JumpGUID = AddActivity(Player.Object, Actions.Directing, JumpGUID, vecDirect, (Player.MoveSpeed * 4), moveFriction)
                        End If
                    End If
                    
                End If
            ElseIf (ToggleMotion(ToggleIdents.MoveJump) = DIK_SPACE) Then
                ToggleMotion(ToggleIdents.MoveJump) = 0
            End If
        End If
    End If

    DIMouseDevice.GetDeviceStateMouse DIMOUSESTATE

    Dim mX As Single
    Dim mY As Single
    Dim mZ As Single

    mX = DIMOUSESTATE.lX
    mY = DIMOUSESTATE.lY
    mZ = DIMOUSESTATE.lZ
    
    If (GetActiveWindow = frmMain.hwnd) And TrapMouse Then

        If Not (frmMain.MousePointer = 99) Then
            frmMain.MousePointer = 99
            frmMain.MouseIcon = LoadPicture(AppPath & "mouse.cur")
        End If
        
        If Player.Object.Visible Then
            MouseLook mX, mY, mZ
     
            If (((Not PauseGame) And TrapMouse) And (Not ConsoleVisible)) Then
            
                If (DIKEYBOARDSTATE.Key(DIK_RIGHTARROW)) And (Not uses(DIK_RIGHTARROW)) Then
                
        '        Case MouseRight
                    '###############  ################
                    MouseLook 20, 0, 0
                    ResetIdle

                ElseIf (DIKEYBOARDSTATE.Key(DIK_LEFTARROW)) And (Not uses(DIK_LEFTARROW)) Then
                
        '        Case MouseLeft
                    '###############  ################
                    MouseLook -20, 0, 0
                    ResetIdle

                ElseIf (DIKEYBOARDSTATE.Key(DIK_DOWNARROW)) And (Not uses(DIK_DOWNARROW)) Then
                
        '        Case MouseDown
                    '###############  ################
                    MouseLook 0, 20, 0
                    ResetIdle

                ElseIf (DIKEYBOARDSTATE.Key(DIK_UPARROW)) And (Not uses(DIK_UPARROW)) Then
                
        '        Case MouseUp
                    '###############  ################
                    MouseLook 0, -20, 0
                    ResetIdle

                End If
    
                If DIMOUSESTATE.Buttons(0) Then 'left
                    If Not (ToggleMouse1 = DIMOUSESTATE.Buttons(0)) Then
                        ToggleMouse1 = DIMOUSESTATE.Buttons(0)
                        
                        '###############  ################
                        ResetIdle
                        If (Perspective = Playmode.Spectator) Then
                            Player.CameraIndex = Player.CameraIndex + 1
                            If Player.CameraIndex > CameraCount Then
                                Player.CameraIndex = 0
                            End If

                            If Player.CameraIndex = 0 Then
                                FadeMessage "Spectator View"
                            Else
                                FadeMessage "Camera View " & Player.CameraIndex
                            End If
                        End If
                        
                        
                    End If
                ElseIf Not (ToggleMouse1 = 0) Then
                    ToggleMouse1 = 0
                End If
                
                If DIMOUSESTATE.Buttons(1) Then 'right
                    If Not (ToggleMouse2 = DIMOUSESTATE.Buttons(1)) Then
                        ToggleMouse2 = DIMOUSESTATE.Buttons(1)
                        
                        '###############  ################
                        ResetIdle
                       
                    End If
                ElseIf Not (ToggleMouse2 = 0) Then
                    ToggleMouse2 = 0
                End If
                                

            End If
        End If
    
        Dim rec As RECT
        GetWindowRect frmMain.hwnd, rec
        SetCursorPos rec.Right + ((rec.Left - rec.Right) / 2), rec.Top + ((rec.Bottom - rec.Top) / 2)

        
    'ElseIf (GetActiveWindow = frmMain.hwnd) And Not TrapMouse Then
        'TrapMouse = True
    Else
        If Not (frmMain.MousePointer = 1) Then frmMain.MousePointer = 1
    End If

    Exit Sub
pausing:
    Err.Clear
    DoPauseGame
End Sub

Public Sub MouseLook(ByVal X As Integer, ByVal Y As Integer, ByVal z As Integer)

    Dim cnt As Long

    If Perspective = ThirdPerson Then
    
        If z < 0 Then
            Player.CameraZoom = Player.CameraZoom + 0.25
        ElseIf z > 0 Then
            Player.CameraZoom = Player.CameraZoom - 0.25
        End If
        
        If Player.CameraZoom > MaxCameraZoom Then Player.CameraZoom = MaxCameraZoom
        If Player.CameraZoom < MinCameraZoom Then Player.CameraZoom = MinCameraZoom
    
    End If
    
    If Perspective = CameraMode Then
        If Player.CameraIndex > 0 Then
            Player.CameraAngle = Cameras(Player.CameraIndex).Angle
        End If
    Else
        If X < 0 Then
            For cnt = (X * MouseSensitivity) To 0
                Player.CameraAngle = Player.CameraAngle - -0.0015
            Next
        ElseIf X > 0 Then
            For cnt = 0 To (X * MouseSensitivity)
                Player.CameraAngle = Player.CameraAngle - 0.0015
            Next
        End If
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
    
'    If Surface Then
'
'        DDevice.SetPixelShader PixelShaderDefault
'
'        DDevice.SetVertexShader FVF_SCREEN
'        DDevice.SetRenderState D3DRS_ZENABLE, False
'        DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
'
'        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
'        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
'
'        DDevice.SetMaterial GenericMaterial
'
'        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
'        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
'
'        DDevice.SetTexture 0, CtrlsLeftText
'        DDevice.SetTexture 1, Nothing
'        DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, CtrlsLeftVertex(0), LenB(CtrlsLeftVertex(0))
'
'        DDevice.SetTexture 0, CtrlsRightText
'        DDevice.SetTexture 1, Nothing
'        DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, CtrlsRightVertex(0), LenB(CtrlsRightVertex(0))
'
'    End If
    
    If ConsoleVisible Then
    
        Dim lastColor As Long
        lastColor = TextColor
        TextColor = D3DColorARGB(255, 255, 255, 255)
    
        DDevice.SetVertexShader FVF_SCREEN
        DDevice.SetRenderState D3DRS_ZENABLE, False
        DDevice.SetRenderState D3DRS_LIGHTING, False
        DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE

        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, False

        DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
        DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
        DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
        DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_LINEAR
        
        DDevice.SetTexture 0, Backdrop
        DDevice.SetTexture 1, Nothing
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
        
        TextColor = lastColor

        DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
        DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
        DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
        DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
        
        DDevice.SetRenderState D3DRS_ZENABLE, 1
        DDevice.SetRenderState D3DRS_LIGHTING, 1
    End If
End Sub

Public Function AddMessage(ByVal Message As String)
'    If ConsoleMsgs.Count > 0 Then
'        If (ConsoleMsgs.Item(ConsoleMsgs.Count) = Message) Then
'            Exit Function
'        End If
'    End If
        
    If ConsoleMsgs.Count > MaxConsoleMsgs Then
        ConsoleMsgs.Remove 1
    End If
    ConsoleMsgs.Add Message
    
End Function

Public Sub ConsoleInput(ByRef kState As DIKEYBOARDSTATE)

    If ConsoleVisible And (Not (Shift < 0)) And GetActiveWindow = frmMain.hwnd Then
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
                        ResetIdle
                    ElseIf ((Timer - KeyState(cnt).VKLatency) > 0.6) And (KeyState(cnt).VKLatency > 0) Then
                        KeyState(cnt).VKPressed = True
                        ResetIdle
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
                
                    ResetIdle
                    HistoryMsgs.Add CommandLine
                    If HistoryMsgs.Count > MaxHistoryMsgs Then
                        HistoryMsgs.Remove 1
                    End If
                    HistoryPoint = HistoryMsgs.Count + 1
                    
                    Process CommandLine
                    
                    CommandLine = ""
                    CursorPos = 0
                    
                ElseIf cnt = DIK_BACK Then
                
                    ResetIdle
                    If CursorPos > 0 Then
                        CommandLine = Left(CommandLine, CursorPos - 1) & Mid(CommandLine, CursorPos + 1)
                        CursorPos = CursorPos - 1
                    End If
                
                ElseIf cnt = DIK_DELETE Then
                
                    ResetIdle
                    CommandLine = Left(CommandLine, CursorPos) & Mid(CommandLine, CursorPos + 2)
                
                ElseIf cnt = DIK_LEFT Then
                
                    ResetIdle
                    If CursorPos > 0 Then
                        CursorPos = CursorPos - 1
                    End If
                    
                ElseIf cnt = DIK_HOME Then
                
                    ResetIdle
                    CursorPos = 0
                    
                ElseIf cnt = DIK_END Then
                
                    ResetIdle
                    CursorPos = Len(CommandLine)
                    
                ElseIf cnt = DIK_RIGHT Then
                
                    ResetIdle
                    If CursorPos < Len(CommandLine) Then
                        CursorPos = CursorPos + 1
                    End If
                    
                ElseIf cnt = DIK_UP Then
                
                    ResetIdle
                    If HistoryMsgs.Count > 0 Then
                        If HistoryPoint > 1 Then
                            HistoryPoint = HistoryPoint - 1
                            CommandLine = HistoryMsgs(HistoryPoint)
                            CursorPos = Len(CommandLine)
                        End If
                    End If
                    
                ElseIf cnt = DIK_DOWN Then
                
                    ResetIdle
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
                
                    ResetIdle
                    If Len(CommandLine) <= 40 Then
                        CommandLine = Left(CommandLine, CursorPos) & vbTab & Mid(CommandLine, CursorPos + 1)
                        CursorPos = CursorPos + 1
                    End If
                    
                Else
                    char = KeyChars(cnt)
                    If Not char = "" Then
                    
                        ResetIdle
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
                        
                        If Len(CommandLine) <= ColumnCount Then
                        
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
Public Function BindingIndex(ByVal KeyString As String) As Integer
    Select Case UCase(KeyString)
        Case "0"
            BindingIndex = DIK_0
        Case "1"
            BindingIndex = DIK_1
        Case "2"
            BindingIndex = DIK_2
        Case "3"
            BindingIndex = DIK_3
        Case "4"
            BindingIndex = DIK_4
        Case "5"
            BindingIndex = DIK_5
        Case "6"
            BindingIndex = DIK_6
        Case "7"
            BindingIndex = DIK_7
        Case "8"
            BindingIndex = DIK_8
        Case "9"
            BindingIndex = DIK_9
        Case "A"
            BindingIndex = DIK_A
        Case "ABNT_C1"
            BindingIndex = DIK_ABNT_C1
        Case "ABNT_C2"
            BindingIndex = DIK_ABNT_C2
        Case "ADD"
            BindingIndex = DIK_ADD
        Case "APOSTROPHE"
            BindingIndex = DIK_APOSTROPHE
        Case "APPS"
            BindingIndex = DIK_APPS
        Case "AT"
            BindingIndex = DIK_AT
        Case "AX"
            BindingIndex = DIK_AX
        Case "B"
            BindingIndex = DIK_B
        Case "BACK"
            BindingIndex = DIK_BACK
        Case "BACKSLASH"
            BindingIndex = DIK_BACKSLASH
        Case "BACKSPACE"
            BindingIndex = DIK_BACKSPACE
        Case "C"
            BindingIndex = DIK_C
        Case "CALCULATOR"
            BindingIndex = DIK_CALCULATOR
        Case "CAPITAL"
            BindingIndex = DIK_CAPITAL
        Case "CAPSLOCK"
            BindingIndex = DIK_CAPSLOCK
        Case "CIRCUMFLEX"
            BindingIndex = DIK_CIRCUMFLEX
        Case "COLON"
            BindingIndex = DIK_COLON
        Case "COMMA"
            BindingIndex = DIK_COMMA
        Case "CONVERT"
            BindingIndex = DIK_CONVERT
        Case "D"
            BindingIndex = DIK_D
        Case "DECIMAL"
            BindingIndex = DIK_DECIMAL
        Case "DELETE"
            BindingIndex = DIK_DELETE
        Case "DIVIDE"
            BindingIndex = DIK_DIVIDE
        Case "DOWN"
            BindingIndex = DIK_DOWN
        Case "DOWNARROW"
            BindingIndex = DIK_DOWNARROW
        Case "E"
            BindingIndex = DIK_E
        Case "END"
            BindingIndex = DIK_END
        Case "EQUALS"
            BindingIndex = DIK_EQUALS
        Case "ESCAPE"
            BindingIndex = DIK_ESCAPE
        Case "F"
            BindingIndex = DIK_F
        Case "F1"
            BindingIndex = DIK_F1
        Case "F10"
            BindingIndex = DIK_F10
        Case "F11"
            BindingIndex = DIK_F11
        Case "F12"
            BindingIndex = DIK_F12
        Case "F13"
            BindingIndex = DIK_F13
        Case "F14"
            BindingIndex = DIK_F14
        Case "F15"
            BindingIndex = DIK_F15
        Case "F2"
            BindingIndex = DIK_F2
        Case "F3"
            BindingIndex = DIK_F3
        Case "F4"
            BindingIndex = DIK_F4
        Case "F5"
            BindingIndex = DIK_F5
        Case "F6"
            BindingIndex = DIK_F6
        Case "F7"
            BindingIndex = DIK_F7
        Case "F8"
            BindingIndex = DIK_F8
        Case "F9"
            BindingIndex = DIK_F9
        Case "G"
            BindingIndex = DIK_G
        Case "GRAVE"
            BindingIndex = DIK_GRAVE
        Case "H"
            BindingIndex = DIK_H
        Case "HOME"
            BindingIndex = DIK_HOME
        Case "I"
            BindingIndex = DIK_I
        Case "INSERT"
            BindingIndex = DIK_INSERT
        Case "J"
            BindingIndex = DIK_J
        Case "K"
            BindingIndex = DIK_K
        Case "KANA"
            BindingIndex = DIK_KANA
        Case "KANJI"
            BindingIndex = DIK_KANJI
        Case "L"
            BindingIndex = DIK_L
        Case "LALT"
            BindingIndex = DIK_LALT
        Case "LBRACKET"
            BindingIndex = DIK_LBRACKET
        Case "LCONTROL"
            BindingIndex = DIK_LCONTROL
        Case "LEFT"
            BindingIndex = DIK_LEFT
        Case "LEFTARROW"
            BindingIndex = DIK_LEFTARROW
        Case "LMENU"
            BindingIndex = DIK_LMENU
        Case "LSHIFT"
            BindingIndex = DIK_LSHIFT
        Case "LWIN"
            BindingIndex = DIK_LWIN
        Case "M"
            BindingIndex = DIK_M
        Case "MAIL"
            BindingIndex = DIK_MAIL
        Case "MEDIASELECT"
            BindingIndex = DIK_MEDIASELECT
        Case "MEDIASTOP"
            BindingIndex = DIK_MEDIASTOP
        Case "MINUS"
            BindingIndex = DIK_MINUS
        Case "MULTIPLY"
            BindingIndex = DIK_MULTIPLY
        Case "MUTE"
            BindingIndex = DIK_MUTE
        Case "MYCOMPUTER"
            BindingIndex = DIK_MYCOMPUTER
        Case "N"
            BindingIndex = DIK_N
        Case "NEXT"
            BindingIndex = DIK_NEXT
        Case "NEXTTRACK"
            BindingIndex = DIK_NEXTTRACK
        Case "NOCONVERT"
            BindingIndex = DIK_NOCONVERT
        Case "NUMLOCK"
            BindingIndex = DIK_NUMLOCK
        Case "NUMPAD0"
            BindingIndex = DIK_NUMPAD0
        Case "NUMPAD1"
            BindingIndex = DIK_NUMPAD1
        Case "NUMPAD2"
            BindingIndex = DIK_NUMPAD2
        Case "NUMPAD3"
            BindingIndex = DIK_NUMPAD3
        Case "NUMPAD4"
            BindingIndex = DIK_NUMPAD4
        Case "NUMPAD5"
            BindingIndex = DIK_NUMPAD5
        Case "NUMPAD6"
            BindingIndex = DIK_NUMPAD6
        Case "NUMPAD7"
            BindingIndex = DIK_NUMPAD7
        Case "NUMPAD8"
            BindingIndex = DIK_NUMPAD8
        Case "NUMPAD9"
            BindingIndex = DIK_NUMPAD9
        Case "NUMPADCOMMA"
            BindingIndex = DIK_NUMPADCOMMA
        Case "NUMPADENTER"
            BindingIndex = DIK_NUMPADENTER
        Case "NUMPADEQUALS"
            BindingIndex = DIK_NUMPADEQUALS
        Case "NUMPADMINUS"
            BindingIndex = DIK_NUMPADMINUS
        Case "NUMPADPERIOD"
            BindingIndex = DIK_NUMPADPERIOD
        Case "NUMPADPLUS"
            BindingIndex = DIK_NUMPADPLUS
        Case "NUMPADSLASH"
            BindingIndex = DIK_NUMPADSLASH
        Case "NUMPADSTAR"
            BindingIndex = DIK_NUMPADSTAR
        Case "O"
            BindingIndex = DIK_O
        Case "OEM_102"
            BindingIndex = DIK_OEM_102
        Case "P"
            BindingIndex = DIK_P
        Case "PAUSE"
            BindingIndex = DIK_PAUSE
        Case "PERIOD"
            BindingIndex = DIK_PERIOD
        Case "PGDN"
            BindingIndex = DIK_PGDN
        Case "PGUP"
            BindingIndex = DIK_PGUP
        Case "PLAYPAUSE"
            BindingIndex = DIK_PLAYPAUSE
        Case "POWER"
            BindingIndex = DIK_POWER
        Case "PREVTRACK"
            BindingIndex = DIK_PREVTRACK
        Case "PRIOR"
            BindingIndex = DIK_PRIOR
        Case "Q"
            BindingIndex = DIK_Q
        Case "R"
            BindingIndex = DIK_R
        Case "RALT"
            BindingIndex = DIK_RALT
        Case "RBRACKET"
            BindingIndex = DIK_RBRACKET
        Case "RCONTROL"
            BindingIndex = DIK_RCONTROL
        Case "RETURN"
            BindingIndex = DIK_RETURN
        Case "RIGHT"
            BindingIndex = DIK_RIGHT
        Case "RIGHTARROW"
            BindingIndex = DIK_RIGHTARROW
        Case "RMENU"
            BindingIndex = DIK_RMENU
        Case "RSHIFT"
            BindingIndex = DIK_RSHIFT
        Case "RWIN"
            BindingIndex = DIK_RWIN
        Case "S"
            BindingIndex = DIK_S
        Case "SCROLL"
            BindingIndex = DIK_SCROLL
        Case "SEMICOLON"
            BindingIndex = DIK_SEMICOLON
        Case "SLASH"
            BindingIndex = DIK_SLASH
        Case "SLEEP"
            BindingIndex = DIK_SLEEP
        Case "STOP"
            BindingIndex = DIK_STOP
        Case "SUBTRACT"
            BindingIndex = DIK_SUBTRACT
        Case "SYSRQ"
            BindingIndex = DIK_SYSRQ
        Case "T"
            BindingIndex = DIK_T
        Case "TAB"
            BindingIndex = DIK_TAB
        Case "U"
            BindingIndex = DIK_U
        Case "UNDERLINE"
            BindingIndex = DIK_UNDERLINE
        Case "UNLABELED"
            BindingIndex = DIK_UNLABELED
        Case "UP"
            BindingIndex = DIK_UP
        Case "UPARROW"
            BindingIndex = DIK_UPARROW
        Case "V"
            BindingIndex = DIK_V
        Case "VOLUMEDOWN"
            BindingIndex = DIK_VOLUMEDOWN
        Case "VOLUMEUP"
            BindingIndex = DIK_VOLUMEUP
        Case "W"
            BindingIndex = DIK_W
        Case "WAKE"
            BindingIndex = DIK_WAKE
        Case "WEBBACK"
            BindingIndex = DIK_WEBBACK
        Case "WEBFAVORITES"
            BindingIndex = DIK_WEBFAVORITES
        Case "WEBFORWARD"
            BindingIndex = DIK_WEBFORWARD
        Case "WEBHOME"
            BindingIndex = DIK_WEBHOME
        Case "WEBREFRESH"
            BindingIndex = DIK_WEBREFRESH
        Case "WEBSEARCH"
            BindingIndex = DIK_WEBSEARCH
        Case "WEBSTOP"
            BindingIndex = DIK_WEBSTOP
        Case "X"
            BindingIndex = DIK_X
        Case "Y"
            BindingIndex = DIK_Y
        Case "YEN"
            BindingIndex = DIK_YEN
        Case "Z"
            BindingIndex = DIK_Z
        Case Else
            BindingIndex = -1
    End Select
End Function

'Public Function SurfaceHit(ByVal X As Single, ByVal Y As Single) As SurfaceControl
'    Dim idx As Single
'    Dim ysize As Single
'    Dim xsize As Single
'
'    If (X >= CtrlsLeftPos.Left And X <= CtrlsLeftPos.Right) And (Y >= CtrlsLeftPos.Top And Y <= CtrlsLeftPos.Bottom) Then
'        idx = Y - CtrlsLeftPos.Top
'        ysize = ((CtrlsLeftPos.Bottom - CtrlsLeftPos.Top) / 4)
'
'        idx = (idx \ ysize) + 1
'
'        If idx = 2 Then
'            xsize = ((CtrlsLeftPos.Right - CtrlsLeftPos.Left) / 2)
'            If (X - CtrlsLeftPos.Left) / xsize <= 1 And (X - CtrlsLeftPos.Left) >= 0 Then
'                SurfaceHit = SurfaceHit + LeftStep
'            ElseIf (X - CtrlsLeftPos.Left) / xsize <= 2 And (X - CtrlsLeftPos.Left) >= xsize Then
'                SurfaceHit = SurfaceHit + RightStep
'            End If
'        Else
'            xsize = ((CtrlsLeftPos.Right - CtrlsLeftPos.Left) / 4)
'            If (((X - CtrlsLeftPos.Left) - xsize) >= 0) And (((X - CtrlsLeftPos.Left) - xsize) <= (xsize * 2)) Then
'                Select Case idx
'                    Case 1
'                        SurfaceHit = SurfaceHit + Forward
'                    Case 3
'                        SurfaceHit = SurfaceHit + Backward
'                    Case 4
'
'                End Select
'            End If
'
'        End If
'
'    End If
'
'    If (X >= CtrlsRightPos.Left And X <= CtrlsRightPos.Right) And (Y >= CtrlsRightPos.Top And Y <= CtrlsRightPos.Bottom) Then
'        idx = Y - CtrlsRightPos.Top
'        ysize = ((CtrlsRightPos.Bottom - CtrlsRightPos.Top) / 4)
'
'        idx = (idx \ ysize) + 1
'
'        If idx = 2 Then
'            xsize = ((CtrlsRightPos.Right - CtrlsRightPos.Left) / 2)
'            If (X - CtrlsRightPos.Left) / xsize <= 1 And (X - CtrlsRightPos.Left) >= 0 Then
'                SurfaceHit = SurfaceHit + MouseLeft
'            ElseIf (X - CtrlsRightPos.Left) / xsize <= 2 And (X - CtrlsRightPos.Left) >= xsize Then
'                SurfaceHit = SurfaceHit + MouseRight
'            End If
'        Else
'            xsize = ((CtrlsRightPos.Right - CtrlsRightPos.Left) / 4)
'            If (((X - CtrlsRightPos.Left) - xsize) >= 0) And (((X - CtrlsRightPos.Left) - xsize) <= (xsize * 2)) Then
'                Select Case idx
'                    Case 1
'                        SurfaceHit = SurfaceHit + MouseUp
'                    Case 3
'                        SurfaceHit = SurfaceHit + MouseDown
'                    Case 4
'                        SurfaceHit = SurfaceHit + Jump
'                End Select
'            End If
'
'        End If
'    End If
'End Function

Public Sub CreateCmds()
    
    IdleInput = Timer - 4
'    If Surface Then
'
'        DPI = GetMonitorDPI
'
'        CtrlsLeftPos.Left = 0
'        CtrlsLeftPos.Right = 1.5 * DPI.width
'        CtrlsLeftPos.Top = ((frmMain.ScaleHeight / Screen.TwipsPerPixelY) / 2) - ((3 * DPI.height) / 2)
'        CtrlsLeftPos.Bottom = ((frmMain.ScaleHeight / Screen.TwipsPerPixelY) / 2) + ((3 * DPI.height) / 2)
'
'        Set CtrlsLeftText = LoadTexture(AppPath & "ctrls.bmp")
'        CtrlsLeftVertex(0) = MakeScreen(CtrlsLeftPos.Left, CtrlsLeftPos.Top, -1, 0, 0)
'        CtrlsLeftVertex(1) = MakeScreen(CtrlsLeftPos.Right, CtrlsLeftPos.Top, -1, 1, 0)
'        CtrlsLeftVertex(2) = MakeScreen(CtrlsLeftPos.Left, CtrlsLeftPos.Bottom, -1, 0, 1)
'        CtrlsLeftVertex(3) = MakeScreen(CtrlsLeftPos.Right, CtrlsLeftPos.Bottom, -1, 1, 1)
'
'        CtrlsRightPos.Left = (frmMain.ScaleWidth / Screen.TwipsPerPixelX) - (1.5 * DPI.width)
'        CtrlsRightPos.Right = (frmMain.ScaleWidth / Screen.TwipsPerPixelX)
'        CtrlsRightPos.Top = ((frmMain.ScaleHeight / Screen.TwipsPerPixelY) / 2) - ((3 * DPI.height) / 2)
'        CtrlsRightPos.Bottom = ((frmMain.ScaleHeight / Screen.TwipsPerPixelY) / 2) + ((3 * DPI.height) / 2)
'
'        Set CtrlsRightText = LoadTexture(AppPath & "ctrls.bmp")
'        CtrlsRightVertex(0) = MakeScreen(CtrlsRightPos.Left, CtrlsRightPos.Top, -1, 0, 0)
'        CtrlsRightVertex(1) = MakeScreen(CtrlsRightPos.Right, CtrlsRightPos.Top, -1, 1, 0)
'        CtrlsRightVertex(2) = MakeScreen(CtrlsRightPos.Left, CtrlsRightPos.Bottom, -1, 0, 1)
'        CtrlsRightVertex(3) = MakeScreen(CtrlsRightPos.Right, CtrlsRightPos.Bottom, -1, 1, 1)
'
'    End If
    
    
    Set ConsoleMsgs = New Collection
    Set HistoryMsgs = New Collection
    
    Bottom = 0
    
    Vertex(0) = MakeScreen(0, 0, -1, 0, 0)
    Vertex(1) = MakeScreen((frmMain.width / Screen.TwipsPerPixelX), 0, -1, 1, 0)
    Vertex(2) = MakeScreen(0, 0, -1, 0, 1)
    Vertex(3) = MakeScreen((frmMain.width / Screen.TwipsPerPixelX), 0, -1, 1, 1)
    
    ConsoleWidth = (frmMain.width / Screen.TwipsPerPixelX)
    ConsoleHeight = (MaxConsoleMsgs * (frmMain.TextHeight("A") / Screen.TwipsPerPixelY)) + (TextSpace * MaxConsoleMsgs) + TextSpace
    If ConsoleHeight > ((frmMain.height / Screen.TwipsPerPixelY) \ 2) Then ConsoleHeight = ((frmMain.height / Screen.TwipsPerPixelY) \ 2)
    
    InitKeys
    
    Set Backdrop = LoadTexture(AppPath & "drop.bmp")
    
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
    
    If CurrentLoadedLevel = "" Then InitialCommands
    
End Sub

Public Sub CleanupCmds()
    ClearText
    
    Dim cnt As Integer
    For cnt = 0 To 255
        Bindings(cnt) = ""
    Next
    
    Erase Draws
    
    If Not (ConsoleMsgs Is Nothing) Then
        Do While ConsoleMsgs.Count > 0
            ConsoleMsgs.Remove 1
        Loop
    End If
    
    If Not (HistoryMsgs Is Nothing) Then
        Do While HistoryMsgs.Count > 0
            HistoryMsgs.Remove 1
        Loop
    End If
    
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

Public Sub Process(ByVal inArg As String)

    Dim o As Long
    Dim l As Long
    Dim cnt As Long
    Dim inNew As String
    Dim inTmp As String
    Dim inX As Single
    Dim inY As Single
    
    Dim inCmd As String
    
    inCmd = RemoveNextArg(inArg, " ")
    If Left(inCmd, 1) = "/" Then inCmd = Mid(inCmd, 2)
    
    Select Case Trim(LCase(inCmd))
        Case "debug"
            DebugMode = Not DebugMode
            If DebugMode Then
                AddMessage "Debug mode enabled."
            Else
                AddMessage "Debug mode disabled."
            End If
        Case "parse"
            If PathExists(inArg, True) Then
                inTmp = ReadFile(inArg)
                If inTmp <> "" Then
                    ParseLand 0, inTmp
                    AddMessage "Parse complete."
                Else
                    AddMessage "Nothing to parse."
                End If
            ElseIf inArg <> "" Then
                AddMessage "Parse complete."
            Else
                AddMessage "Nothing to parse."
            End If
        Case "exit", "quit", "close"
            StopGame = True
        Case "spectate"
            'If Not DebugMode Then
                If Not (Perspective = Spectator) Then
                    Perspective = Spectator
                    AddMessage "Changed to spectate mode."
                Else
                    AddMessage "Already in spectate mode."
                End If
            'Else
            '    AddMessage "Not available in debug mode."
            'End If
        Case "join"
            If Not DebugMode Then
                If (Perspective = Spectator) Then
                    Perspective = ThirdPerson
                    AddMessage "You've entered the game."
                Else
                    
                    AddMessage "Already joined the game."
                End If
            Else
                AddMessage "Not available in debug mode."
            End If
        Case "eval"
            Process ParseValues(inArg)
        Case "echo"
            AddMessage ParseSetGet(0, inArg)
        
        Case "fade"
            FadeMessage inArg
        Case "clear"
            ClearText
        Case "draw"
            inX = RemoveNextArg(inArg, " ")
            inY = RemoveNextArg(inArg, " ")
            PrintText inArg, inX, inY
        Case "print"
            inX = RemoveNextArg(inArg, " ")
            inY = RemoveNextArg(inArg, " ")
            PrintText inArg, (((frmMain.ScaleWidth / Screen.TwipsPerPixelX) - (TextSpace * 2)) / ColumnCount) * inX, Row(inY)
        Case "help", "cmdlist", "?", "--?"
            Select Case LCase(inArg)
                Case "commands"
                    
                    AddMessage ""
                    AddMessage "Console Commands:"
                    AddMessage "   EXIT (completely exit out of the game)"
                    AddMessage "   ECHO <Text> (Displays Text in the console)"
                    AddMessage "   EVAL <Text> (Evalutes Text as a console command)"
                    AddMessage "   FADE <Text> (Displays Text in the center of screen)"
                    AddMessage "   DRAW <X> <Y> <Text> (Draw Text on the screen at X,Y)"
                    AddMessage "   PRINT <Row> <Col> <Test> (Like DRAW but use row,col)"
                    AddMessage "   CLEAR (Clears all text being drawn by DRAW or PRINT)"
                    AddMessage "   STAT (Displays XYZ, Distance, Camera Angle and Pitch)"
                    AddMessage "   LEVEL [<Title>] (Load a Title PX in the Levels Folder)"
                    AddMessage "   RESET (Resets the game reloading everything like new)"
                    AddMessage "   SPECTATE (Change in game involvement to a spectator)"
                    AddMessage "   JOIN (Rejoins the game when yopur in spectator mode)"
                Case "editing"
                    AddMessage ""
                    AddMessage "Editing Commands:"
                    AddMessage "   REFRESH (This command will cause the engine to reload the levels PX)"
                    AddMessage "   LOAD <file> (This command loads the specified file to be edited live)"
                    AddMessage "   VIEW <#[-#]> (This command displays the specified line numbers text)"
                    AddMessage "   LINES <#[-#]> (This command adds lines # to the current loaded file)"
                    AddMessage "   EDIT <#> <text> (This command changes the specified line numbers text)"
                    AddMessage "   SAVE (This command saves any edited changes made with a loaded file)"
                Case Else
                    AddMessage ""
                    AddMessage "For detail help please type one of the following help commands:"
                    AddMessage "   HELP COMMANDS (Displays the help of basic console commands)"
                    AddMessage "   HELP EDITING (Displays the help of editing files in console)"
            End Select
        

        Case "stat"
            AddMessage ""
            AddMessage "Origin X: " & Round(CSng(Player.Object.Origin.X), 3)
            AddMessage "Origin Y: " & Round(CSng(Player.Object.Origin.Y), 3)
            AddMessage "Origin Z: " & Round(CSng(Player.Object.Origin.z), 3)
            AddMessage "Distance: " & Round(CSng(Distance(Player.Object.Origin, MakeVector(0, 0, 0))), 3)
            AddMessage "Angle: " & Round(CSng(Player.CameraAngle), 3)
            AddMessage "Pitch: " & Round(CSng(Player.CameraPitch), 3)
        Case "credits"
            ShowCredits = Not ShowCredits
        Case "showcredits"
            ShowCredits = True
        Case "hidecredits"
            ShowCredits = False
        Case "reset"
            AddMessage "Resetting Game."
            CurrentLoadedLevel = ""
            InitialCommands
            CleanupLand True
            CleanupMove
            CreateMove
            CreateLand
            
        Case "refresh"
            
            CleanupLand
            CleanupMove
            CreateMove
            CreateLand
            AddMessage "Level Refreshed."
        Case "level"
            If PathExists(AppPath & "Levels\" & inArg & ".px", False) Then
                
                If Not CurrentLoadedLevel = "" Then
                    CleanupLand
                    CleanupMove
                    CurrentLoadedLevel = inArg
                    CreateMove
                    CreateLand
                    AddMessage "Level Loaded."
                Else
                    CurrentLoadedLevel = inArg
                End If
                
            ElseIf PathExists(AppPath & "Levels\" & CurrentLoadedLevel & ".px", False) Then
                CleanupLand
                CleanupMove
                CreateMove
                CreateLand
                AddMessage "Level Reloaded."
            Else
                AddMessage "Invalid Level - [" & AppPath & "Levels\" & inArg & ".px" & "]"
            End If
            
            
        Case "load"
            If inArg = "" Then
                If EditFileName = "" Then
                    AddMessage "No file is loaded, use ""LOAD <name>"" to load one."
                Else
                    AddMessage "File loaded [" & AppPath & "Levels\" & EditFileName & ".px" & "]"
                End If
            Else
                If PathExists(AppPath & "Levels\" & inArg & ".px", True) Then
                    EditFileName = inArg
                    EditFileData = ReadFile(AppPath & "Levels\" & inArg & ".px")
                    AddMessage "File loaded [" & AppPath & "Levels\" & inArg & ".px" & "]"
                Else
                    AddMessage "File not found [" & AppPath & "Levels\" & inArg & ".px" & "]"
                End If
            End If
            
        Case "view"
                If PathExists(AppPath & "Levels\" & EditFileName & ".px", True) Then
                    
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
                    AddMessage "File not found [" & AppPath & "Levels\" & EditFileName & ".px" & "]"
                End If
        Case "lines"
                If PathExists(AppPath & "Levels\" & EditFileName & ".px", True) Then
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
                    AddMessage "File not found [" & AppPath & "Levels\" & EditFileName & ".px" & "]"
                End If
        Case "edit"
                If PathExists(AppPath & "Levels\" & EditFileName & ".px", True) Then
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
                    AddMessage "File not found [" & AppPath & "Levels\" & EditFileName & ".px" & "]"
                End If
        Case "save"
                If PathExists(AppPath & "Levels\" & EditFileName & ".px", True) Then
                    WriteFile AppPath & "Levels\" & EditFileName & ".px", EditFileData
                    AddMessage "Saved data file [" & AppPath & "Levels\" & EditFileName & ".px" & "]"
                ElseIf EditFileName = "" Then
                    AddMessage "File not loaded."
                Else
                    AddMessage "File not found [" & AppPath & "Levels\" & EditFileName & ".px" & "]"
                End If
        Case ""
        Case Else
            AddMessage "Unknown command."
    End Select
End Sub


