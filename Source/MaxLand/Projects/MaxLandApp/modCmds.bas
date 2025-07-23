Attribute VB_Name = "modCmds"
#Const modCmds = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module


Private Const MaxConsoleMsgs = 20
Private Const MaxHistoryMsgs = 10

Private ConsoleMsgs As VBA.Collection
Private HistoryMsgs As VBA.Collection
Private HistoryPoint As Integer
Private CommandLine As String

Private DInput As DirectInput8
Private DIKeyBoardDevice As DirectInputDevice8
Private DIKEYBOARDSTATE As DIKEYBOARDSTATE

Private DIMouseDevice As DirectInputDevice8
Private DIMOUSESTATE As DIMOUSESTATE

Private TogglePress1 As Long
Private TogglePress2 As Long
Private TogglePress3 As Long

Private Type KeyState
    VKState As Integer
    VKToggle As Boolean
    VKPressed As Boolean
    VKLatency As Single
 End Type

Private IdleInput As Single


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




Private ToggleSound1 As Boolean
Private ToggleSound2 As Boolean
Private ToggleSound3 As Boolean

Private ToggleMouse1 As Long
Private ToggleMouse2 As Long


Private Enum ToggleIdents
    MoveAuto = 0
    MoveJump = 1
    MoveForward = 2
    MoveBackward = 3
    MoveStepLeft = 4
    MoveStepRight = 5
End Enum

Private ToggleMotion(0 To 5) As Long

Public JumpGUID As String

Private lX As Integer
Private lY As Integer
Private lZ As Integer

'Public Bindings(0 To 255) As String

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
        Case "CONTROLLER"
            GetBindingIndex = -2
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
        Case "Next"
            GetBindingIndex = DIK_NEXT
        Case "NextTRACK"
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

Public Function GetBindingText(ByVal BindIndex As Integer) As String
    Select Case BindIndex
        Case -2
            GetBindingText = "CONTROLLER"
        Case DIK_0
            GetBindingText = "0"
        Case DIK_1
            GetBindingText = "1"
        Case DIK_2
            GetBindingText = "2"
        Case DIK_3
            GetBindingText = "3"
        Case DIK_4
            GetBindingText = "4"
        Case DIK_5
            GetBindingText = "5"
        Case DIK_6
            GetBindingText = "6"
        Case DIK_7
            GetBindingText = "7"
        Case DIK_8
            GetBindingText = "8"
        Case DIK_9
            GetBindingText = "9"
        Case DIK_A
            GetBindingText = "A"
        Case DIK_ABNT_C1
            GetBindingText = "ABNT_C1"
        Case DIK_ABNT_C2
            GetBindingText = "ABNT_C2"
        Case DIK_ADD
            GetBindingText = "ADD"
        Case DIK_APOSTROPHE
            GetBindingText = "APOSTROPHE"
        Case DIK_APPS
            GetBindingText = "APPS"
        Case DIK_AT
            GetBindingText = "AT"
        Case DIK_AX
            GetBindingText = "AX"
        Case DIK_B
            GetBindingText = "B"
        Case DIK_BACK
            GetBindingText = "BACK"
        Case DIK_BACKSLASH
            GetBindingText = "BACKSLASH"
        Case DIK_BACKSPACE
            GetBindingText = "BACKSPACE"
        Case DIK_C
            GetBindingText = "C"
        Case DIK_CALCULATOR
            GetBindingText = "CALCULATOR"
        Case DIK_CAPITAL
            GetBindingText = "CAPITAL"
        Case DIK_CAPSLOCK
            GetBindingText = "CAPSLOCK"
        Case DIK_CIRCUMFLEX
            GetBindingText = "CIRCUMFLEX"
        Case DIK_COLON
            GetBindingText = "COLON"
        Case DIK_COMMA
            GetBindingText = "COMMA"
        Case DIK_CONVERT
            GetBindingText = "CONVERT"
        Case DIK_D
            GetBindingText = "D"
        Case DIK_DECIMAL
            GetBindingText = "DECIMAL"
        Case DIK_DELETE
            GetBindingText = "DELETE"
        Case DIK_DIVIDE
            GetBindingText = "DIVIDE"
        Case DIK_DOWN
            GetBindingText = "DOWN"
        Case DIK_DOWNARROW
            GetBindingText = "DOWNARROW"
        Case DIK_E
            GetBindingText = "E"
        Case DIK_END
            GetBindingText = "END"
        Case DIK_EQUALS
            GetBindingText = "EQUALS"
        Case DIK_ESCAPE
            GetBindingText = "ESCAPE"
        Case DIK_F
            GetBindingText = "F"
        Case DIK_F1
            GetBindingText = "F1"
        Case DIK_F10
            GetBindingText = "F10"
        Case DIK_F11
            GetBindingText = "F11"
        Case DIK_F12
            GetBindingText = "F12"
        Case DIK_F13
            GetBindingText = "F13"
        Case DIK_F14
            GetBindingText = "F14"
        Case DIK_F15
            GetBindingText = "F15"
        Case DIK_F2
            GetBindingText = "F2"
        Case DIK_F3
            GetBindingText = "F3"
        Case DIK_F4
            GetBindingText = "F4"
        Case DIK_F5
            GetBindingText = "F5"
        Case DIK_F6
            GetBindingText = "F6"
        Case DIK_F7
            GetBindingText = "F7"
        Case DIK_F8
            GetBindingText = "F8"
        Case DIK_F9
            GetBindingText = "F9"
        Case DIK_G
            GetBindingText = "G"
        Case DIK_GRAVE
            GetBindingText = "GRAVE"
        Case DIK_H
            GetBindingText = "H"
        Case DIK_HOME
            GetBindingText = "HOME"
        Case DIK_I
            GetBindingText = "I"
        Case DIK_INSERT
            GetBindingText = "INSERT"
        Case DIK_J
            GetBindingText = "J"
        Case DIK_K
            GetBindingText = "K"
        Case DIK_KANA
            GetBindingText = "KANA"
        Case DIK_KANJI
            GetBindingText = "KANJI"
        Case DIK_L
            GetBindingText = "L"
        Case DIK_LALT
            GetBindingText = "LALT"
        Case DIK_LBRACKET
            GetBindingText = "LBRACKET"
        Case DIK_LCONTROL
            GetBindingText = "LCONTROL"
        Case DIK_LEFT
            GetBindingText = "LEFT"
        Case DIK_LEFTARROW
            GetBindingText = "LEFTARROW"
        Case DIK_LMENU
            GetBindingText = "LMENU"
        Case DIK_LSHIFT
            GetBindingText = "LSHIFT"
        Case DIK_LWIN
            GetBindingText = "LWIN"
        Case DIK_M
            GetBindingText = "M"
        Case DIK_MAIL
            GetBindingText = "MAIL"
        Case DIK_MEDIASELECT
            GetBindingText = "MEDIASELECT"
        Case DIK_MEDIASTOP
            GetBindingText = "MEDIASTOP"
        Case DIK_MINUS
            GetBindingText = "MINUS"
        Case DIK_MULTIPLY
            GetBindingText = "MULTIPLY"
        Case DIK_MUTE
            GetBindingText = "MUTE"
        Case DIK_MYCOMPUTER
            GetBindingText = "MYCOMPUTER"
        Case DIK_N
            GetBindingText = "N"
        Case DIK_NEXTTRACK
            GetBindingText = "NextTRACK"
        Case DIK_NOCONVERT
            GetBindingText = "NOCONVERT"
        Case DIK_NUMLOCK
            GetBindingText = "NUMLOCK"
        Case DIK_NUMPAD0
            GetBindingText = "NUMPAD0"
        Case DIK_NUMPAD1
            GetBindingText = "NUMPAD1"
        Case DIK_NUMPAD2
            GetBindingText = "NUMPAD2"
        Case DIK_NUMPAD3
            GetBindingText = "NUMPAD3"
        Case DIK_NUMPAD4
            GetBindingText = "NUMPAD4"
        Case DIK_NUMPAD5
            GetBindingText = "NUMPAD5"
        Case DIK_NUMPAD6
            GetBindingText = "NUMPAD6"
        Case DIK_NUMPAD7
            GetBindingText = "NUMPAD7"
        Case DIK_NUMPAD8
            GetBindingText = "NUMPAD8"
        Case DIK_NUMPAD9
            GetBindingText = "NUMPAD9"
        Case DIK_NUMPADCOMMA
            GetBindingText = "NUMPADCOMMA"
        Case DIK_NUMPADENTER
            GetBindingText = "NUMPADENTER"
        Case DIK_NUMPADEQUALS
            GetBindingText = "NUMPADEQUALS"
        Case DIK_NUMPADMINUS
            GetBindingText = "NUMPADMINUS"
        Case DIK_NUMPADPERIOD
            GetBindingText = "NUMPADPERIOD"
        Case DIK_NUMPADPLUS
            GetBindingText = "NUMPADPLUS"
        Case DIK_NUMPADSLASH
            GetBindingText = "NUMPADSLASH"
        Case DIK_NUMPADSTAR
            GetBindingText = "NUMPADSTAR"
        Case DIK_O
            GetBindingText = "O"
        Case DIK_OEM_102
            GetBindingText = "OEM_102"
        Case DIK_P
            GetBindingText = "P"
        Case DIK_PAUSE
            GetBindingText = "PAUSE"
        Case DIK_PERIOD
            GetBindingText = "PERIOD"
        Case DIK_PGDN
            GetBindingText = "PGDN"
        Case DIK_PGUP
            GetBindingText = "PGUP"
        Case DIK_PLAYPAUSE
            GetBindingText = "PLAYPAUSE"
        Case DIK_NEXT
            GetBindingText = "Next"
        Case DIK_POWER
            GetBindingText = "POWER"
        Case DIK_PREVTRACK
            GetBindingText = "PREVTRACK"
        Case DIK_PRIOR
            GetBindingText = "PRIOR"
        Case DIK_Q
            GetBindingText = "Q"
        Case DIK_R
            GetBindingText = "R"
        Case DIK_RALT
            GetBindingText = "RALT"
        Case DIK_RBRACKET
            GetBindingText = "RBRACKET"
        Case DIK_RCONTROL
            GetBindingText = "RCONTROL"
        Case DIK_RETURN
            GetBindingText = "RETURN"
        Case DIK_RIGHT
            GetBindingText = "RIGHT"
        Case DIK_RIGHTARROW
            GetBindingText = "RIGHTARROW"
        Case DIK_RMENU
            GetBindingText = "RMENU"
        Case DIK_RSHIFT
            GetBindingText = "RSHIFT"
        Case DIK_RWIN
            GetBindingText = "RWIN"
        Case DIK_S
            GetBindingText = "S"
        Case DIK_SCROLL
            GetBindingText = "SCROLL"
        Case DIK_SEMICOLON
            GetBindingText = "SEMICOLON"
        Case DIK_SLASH
            GetBindingText = "SLASH"
        Case DIK_SLEEP
            GetBindingText = "SLEEP"
        Case DIK_STOP
            GetBindingText = "STOP"
        Case DIK_SUBTRACT
            GetBindingText = "SUBTRACT"
        Case DIK_SYSRQ
            GetBindingText = "SYSRQ"
        Case DIK_T
            GetBindingText = "T"
        Case DIK_TAB
            GetBindingText = "TAB"
        Case DIK_U
            GetBindingText = "U"
        Case DIK_UNDERLINE
            GetBindingText = "UNDERLINE"
        Case DIK_UNLABELED
            GetBindingText = "UNLABELED"
        Case DIK_UP
            GetBindingText = "UP"
        Case DIK_UPARROW
            GetBindingText = "UPARROW"
        Case DIK_V
            GetBindingText = "V"
        Case DIK_VOLUMEDOWN
            GetBindingText = "VOLUMEDOWN"
        Case DIK_VOLUMEUP
            GetBindingText = "VOLUMEUP"
        Case DIK_W
            GetBindingText = "W"
        Case DIK_WAKE
            GetBindingText = "WAKE"
        Case DIK_WEBBACK
            GetBindingText = "WEBBACK"
        Case DIK_WEBFAVORITES
            GetBindingText = "WEBFAVORITES"
        Case DIK_WEBFORWARD
            GetBindingText = "WEBFORWARD"
        Case DIK_WEBHOME
            GetBindingText = "WEBHOME"
        Case DIK_WEBREFRESH
            GetBindingText = "WEBREFRESH"
        Case DIK_WEBSEARCH
            GetBindingText = "WEBSEARCH"
        Case DIK_WEBSTOP
            GetBindingText = "WEBSTOP"
        Case DIK_X
            GetBindingText = "X"
        Case DIK_Y
            GetBindingText = "Y"
        Case DIK_YEN
            GetBindingText = "YEN"
        Case DIK_Z
            GetBindingText = "Z"
        Case Else
            GetBindingText = -1
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
'    Dim Friction As Single
'    Friction = 0.05
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
'            If Player.Direct.Y = 0 Then
'                vecDirect.X = Sin(D720 - Player.Angle)
'                vecDirect.z = Cos(D720 - Player.Angle)
'                If ((Perspective = Spectator) Or DebugMode) Or Player.InLiquid Then
'                    vecDirect.Y = -(Tan(D720 - Player.Pitch))
'                End If
'                D3DXVec3Normalize vecDirect, vecDirect
'                If Player.InLiquid And (Not ((Perspective = Spectator) Or DebugMode)) Then
'                    AddMotion Player, Actions.Directing, Replace(modGuid.GUID, "-", "K"), vecDirect, (Player.Speed / 2), Friction
'                Else
'                    AddMotion Player, Actions.Directing, Replace(modGuid.GUID, "-", "K"), vecDirect, Player.Speed, Friction
'                End If
'            End If
'            ResetIdle
'            ToggleJump = False
'        Case Backward
'
'            If Player.Direct.Y = 0 Then
'                vecDirect.X = -Sin(D720 - Player.Angle)
'                vecDirect.z = -Cos(D720 - Player.Angle)
'                If (Perspective = Spectator) Or DebugMode Then
'                    vecDirect.Y = Tan(D720 - Player.Pitch)
'                End If
'                D3DXVec3Normalize vecDirect, vecDirect
'                If Player.InLiquid And (Not ((Perspective = Spectator) Or DebugMode)) Then
'                    AddMotion Player, Actions.Directing, Replace(modGuid.GUID, "-", "K"), vecDirect, (Player.Speed / 2), Friction
'                Else
'                    AddMotion Player, Actions.Directing, Replace(modGuid.GUID, "-", "K"), vecDirect, Player.Speed, Friction
'                End If
'            End If
'            ResetIdle
'            ToggleJump = False
'        Case LeftStep
'            If Player.Direct.Y = 0 Then
'                vecDirect.X = Sin((D720 - Player.Angle) - D180)
'                vecDirect.z = Cos((D720 - Player.Angle) - D180)
'                D3DXVec3Normalize vecDirect, vecDirect
'                If Player.InLiquid And (Not ((Perspective = Spectator) Or DebugMode)) Then
'                    AddMotion Player, Actions.Directing, Replace(modGuid.GUID, "-", "K"), vecDirect, (Player.Speed / 2), Friction
'                Else
'                    AddMotion Player, Actions.Directing, Replace(modGuid.GUID, "-", "K"), vecDirect, Player.Speed, Friction
'                End If
'            End If
'            ResetIdle
'            ToggleJump = False
'        Case RightStep
'            If Player.Direct.Y = 0 Then
'                vecDirect.X = Sin((D720 - Player.Angle) + D180)
'                vecDirect.z = Cos((D720 - Player.Angle) + D180)
'                D3DXVec3Normalize vecDirect, vecDirect
'                If Player.InLiquid And (Not ((Perspective = Spectator) Or DebugMode)) Then
'                    AddMotion Player, Actions.Directing, Replace(modGuid.GUID, "-", "K"), vecDirect, (Player.Speed / 2), Friction
'                Else
'                    AddMotion Player, Actions.Directing, Replace(modGuid.GUID, "-", "K"), vecDirect, Player.Speed, Friction
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
'                    vecDirect.Y = vecDirect.Y + IIf(Player.Speed < 1, 1, Player.Speed)
'                    AddMotion Player, Actions.Directing, Replace(modGuid.GUID, "-", "K"), vecDirect, Player.Speed, Friction
'                Else
'
'                    If Player.Activities.Exists(JumpGUID) Then
'                        If (Not ((Player.IsMoving And Moving.Flying) = Moving.Flying) Or _
'                                ((Player.IsMoving And Moving.Falling) = Moving.Falling)) Then
'                            'If Player.Activities.Exists(JumpGUID) Then
'                                Do Until Not MotionExists(JumpGUID)
'                                    DeleteMotion Player, JumpGUID
'                                Loop
'                            'End If
'                        End If
'                    End If
'                    If Not Player.Activities.Exists(JumpGUID) Then
'                        vecDirect.Y = IIf(Player.InLiquid, 5, 9)
'                        JumpGUID = AddMotion(Player, Actions.Directing, JumpGUID, vecDirect, (Player.Speed * 4), Friction)
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
    On Error GoTo pausing

    DIKeyBoardDevice.GetDeviceStateKeyboard DIKEYBOARDSTATE

    Dim cnt As Long
    Dim uses(0 To 255) As Boolean
    
    If ((GetActiveWindow = frmMain.hwnd)) Then
    
        If (FullScreen And Not TrapMouse) Then TrapMouse = True And (Bindings.MouseInput = Trapping)

        If DIKEYBOARDSTATE.Key(DIK_ESCAPE) Then
            If (Not TogglePress1 = DIK_ESCAPE) Then
                TogglePress1 = DIK_ESCAPE

                '############################################
                '############### UNTRAP/QUIT ################
                ResetIdle

'                If ((Not FullScreen) And TrapMouse) Then
'                    TrapMouse = False
'                ElseIf (FullScreen Or (Not FullScreen And Not TrapMouse)) And (Bindings.MouseInput = Trapping) Then
'                    StopGame = True
'                End If


                If ((Not FullScreen) Or ConsoleVisible) And TrapMouse Then
                    TrapMouse = False
                ElseIf FullScreen Or (Not FullScreen And Not TrapMouse) Then
                    StopGame = True
                End If
                '############################################
                '############################################

            End If
        ElseIf DIKEYBOARDSTATE.Key(DIK_F1) Then
            If (Not TogglePress1 = DIK_F1) Then
                TogglePress1 = DIK_F1

                '############################################
                '################ SHOW HELP #################
                ResetIdle
                ShowHelp = Not ShowHelp
                '############################################
                '############################################

            End If
        ElseIf DIKEYBOARDSTATE.Key(DIK_F2) Then
            If (Not TogglePress1 = DIK_F2) Then
                TogglePress1 = DIK_F2

                '############################################
                '################ SHOW StATS ################
                ResetIdle
                ShowStat = (Not ShowStat) Or (ShowStat And Not ShowHelp)
                If ShowStat Then ShowHelp = True
                '############################################
                '############################################

            End If
        ElseIf DIKEYBOARDSTATE.Key(DIK_F3) Then
            If (Not TogglePress1 = DIK_F3) Then
                TogglePress1 = DIK_F3

                '############################################
                '############### RESET GAME #################
                
                ResetIdle
                Process "reset"

                '############################################
                '############################################
            End If
            
        ElseIf DIKEYBOARDSTATE.Key(DIK_F4) Then
            If (Not TogglePress1 = DIK_F4) Then
                TogglePress1 = DIK_F4
            
                '############################################
                '############### SHOW CREDITS ###############
                ResetIdle
                ShowCredits = Not ShowCredits
                '############################################
                '############################################

            End If
        ElseIf DIKEYBOARDSTATE.Key(DIK_F6) Then
            If (Not TogglePress1 = DIK_F6) Then
                TogglePress1 = DIK_F6
            
                '############################################
                '############ SWITCH PERSPECTIVE ############
                ResetIdle
                If Perspective = Playmode.ThirdPerson Then
                    Perspective = Playmode.FirstPerson
                ElseIf Perspective = Playmode.FirstPerson Then
                    Perspective = IIf((Cameras.Count > 0), Playmode.CameraMode, Playmode.ThirdPerson)
                ElseIf Perspective = Playmode.CameraMode Then
                    Perspective = Playmode.ThirdPerson
                End If
                '############################################
                '############################################

            End If
        ElseIf DIKEYBOARDSTATE.Key(DIK_LALT) Or DIKEYBOARDSTATE.Key(DIK_RALT) Then
            If DIKEYBOARDSTATE.Key(DIK_TAB) Then
                If (Not TogglePress1 = DIK_TAB) Then
                    TogglePress1 = DIK_TAB
                    
                    '############################################
                    '############### ALT+TAB SWAP ###############
                    ResetIdle

                    TrapMouse = False

                    If FullScreen Then
                        frmMain.WindowState = 1
                        DoPauseGame
                    End If

'                    TrapMouse = False
'                    frmMain.WindowState = 1
    
                    '############################################
                    '############################################
                    
                    
                End If
            End If
        ElseIf DIKEYBOARDSTATE.Key(DIK_GRAVE) Then
            If (Not TogglePress1 = DIK_GRAVE) Then
                TogglePress1 = DIK_GRAVE


                '############################################
                '############## TOGGLE CONSOLE ##############
                ResetIdle
                ConsoleToggle

                '############################################
                '############################################
            End If
        ElseIf Not (TogglePress1 = 0) Then
            TogglePress1 = 0
        End If
    
        If (Not TogglePress1 = DIK_TAB) Then ConsoleInput DIKEYBOARDSTATE
    
        Dim rec  As RECT
        Dim mloc As modDecs.POINTAPI
        GetCursorPos mloc
        GetWindowRect frmMain.hwnd, rec
   
        Dim mX As Integer
        Dim mY As Integer
        Dim mZ As Integer
        DIMouseDevice.GetDeviceStateMouse DIMOUSESTATE
        mX = DIMOUSESTATE.lX
        mY = DIMOUSESTATE.lY
        mZ = DIMOUSESTATE.lZ
                       
        If ((TrapMouse Or FullScreen) And (Bindings.MouseInput = Trapping)) Or ((Bindings.MouseInput = Hidden) _
            And ((mloc.X >= rec.Left) And (mloc.X <= rec.Right) And (mloc.Y >= rec.Top) And (mloc.Y <= rec.Bottom))) Then
            If (Not (VB.Screen.MousePointer = 99)) Then
                VB.Screen.MousePointer = 99
                Set VB.Screen.MouseIcon = LoadPicture(AppPath & "mouse.cur")
            End If
            
        ElseIf Not TrapMouse Then
            If (Not (VB.Screen.MousePointer = 0)) Then VB.Screen.MousePointer = 0
        End If


        If Player.Speed > MaxDisplacement Then Player.Speed = MaxDisplacement
        If Player.Speed < 0.01 Then Player.Speed = 0.01


        If ((TrapMouse And (Not ConsoleVisible)) And (Bindings.MouseInput = Trapping)) Or _
            (((Not TrapMouse) And (Not ConsoleVisible)) And (Bindings.MouseInput = Hidden)) Or _
            ((Bindings.MouseInput = Visual) And ((mloc.X >= rec.Left) And (mloc.X <= rec.Right) _
            And (mloc.Y >= rec.Top) And (mloc.Y <= rec.Bottom))) Then
            
            MouseLook mX, mY, mZ
        
            For cnt = 0 To 255
                If DIKEYBOARDSTATE.Key(cnt) Then
                    If Not Bindings(cnt) = "" Then
                        uses(cnt) = True
                        frmMain.ExecuteStatement Bindings(cnt) & vbCrLf
                    End If
                End If
            Next

            If DIKEYBOARDSTATE.Key(DIK_E) And (Bindings(DIK_E) = "") And (Not uses(DIK_E)) Then
                uses(DIK_E) = True
                ResetIdle
                Player.MoveForward
            ElseIf DIKEYBOARDSTATE.Key(DIK_D) And (Bindings(DIK_D) = "") And (Not uses(DIK_D)) Then
                uses(DIK_D) = True
                ResetIdle
                Player.MoveBackwards
            End If

            If DIKEYBOARDSTATE.Key(DIK_W) And (Bindings(DIK_W) = "") And (Not uses(DIK_W)) Then
                uses(DIK_W) = True
                ResetIdle
                Player.SlideLeft
            ElseIf DIKEYBOARDSTATE.Key(DIK_R) And (Bindings(DIK_R) = "") And (Not uses(DIK_R)) Then
                uses(DIK_R) = True
                ResetIdle
                Player.SlideRight
            End If
            
            If DIKEYBOARDSTATE.Key(DIK_SPACE) And (Bindings(DIK_SPACE) = "") And (Not uses(DIK_SPACE)) Then
                uses(DIK_SPACE) = True

                If (ToggleMotion(ToggleIdents.MoveJump) <> DIK_SPACE) Then
                    ToggleMotion(ToggleIdents.MoveJump) = DIK_SPACE
                    
                    ResetIdle
                    Player.Jump
                    
                End If
            ElseIf (ToggleMotion(ToggleIdents.MoveJump) = DIK_SPACE) Then
                ToggleMotion(ToggleIdents.MoveJump) = 0
            End If
            

            
'            If ((mX < 0) And (Bindings(DIK_LEFT) <> "")) And (Not uses(DIK_LEFT)) Then
'                uses(DIK_LEFT) = True
'                For cnt = (mX * MouseSensitivity) To 0
'                    frmMain.RunEvent Bindings(DIK_LEFT)
'                Next
'            ElseIf ((mX > 0) And (Bindings(DIK_RIGHT) <> "")) And (Not uses(DIK_RIGHT)) Then
'                uses(DIK_RIGHT) = True
'                For cnt = 0 To (mX * MouseSensitivity)
'                    frmMain.RunEvent Bindings(DIK_RIGHT)
'                Next
'            End If
'
'            If ((mY < 0) And (Bindings(DIK_DOWN) <> "")) And (Not uses(DIK_DOWN)) Then
'                uses(DIK_DOWN) = True
'                For cnt = (mY * MouseSensitivity) To 0
'                    frmMain.RunEvent Bindings(DIK_DOWN)
'                Next
'            ElseIf ((mY > 0) And (Bindings(DIK_UP) <> "")) And (Not uses(DIK_UP)) Then
'                uses(DIK_UP) = True
'                For cnt = 0 To (mY * MouseSensitivity)
'                    frmMain.RunEvent Bindings(DIK_UP)
'                Next
'            End If

            If (DIMOUSESTATE.Buttons(0) And (Bindings(DIK_LCONTROL) <> "")) And (Not uses(DIK_LCONTROL)) Then
                frmMain.ExecuteStatement Bindings(DIK_LCONTROL)
            End If
            If (DIMOUSESTATE.Buttons(1) And (Bindings(DIK_LALT) <> "")) And (Not uses(DIK_LALT)) Then
                frmMain.ExecuteStatement Bindings(DIK_LALT)
            End If

            If (DIMOUSESTATE.Buttons(2) And (Bindings(DIK_RCONTROL) <> "")) And (Not uses(DIK_RCONTROL)) Then
                frmMain.ExecuteStatement Bindings(DIK_RCONTROL)
            End If
            If (DIMOUSESTATE.Buttons(3) And (Bindings(DIK_RALT) <> "")) And (Not uses(DIK_RALT)) Then
                frmMain.ExecuteStatement Bindings(DIK_RALT)
            End If

            If (Bindings.MouseInput = Trapping) Then
                SetCursorPos (frmMain.Left / VB.Screen.TwipsPerPixelX) + (frmMain.Width / VB.Screen.TwipsPerPixelX / 2), (frmMain.Top / VB.Screen.TwipsPerPixelY) + (frmMain.Height / VB.Screen.TwipsPerPixelY / 2)
            End If
            
        End If
    
        If ((mloc.X > rec.Left) And (mloc.X < rec.Right)) And ((mloc.Y > rec.Top) And (mloc.Y < rec.Bottom)) Then
            If DIMOUSESTATE.Buttons(0) Then
                TrapMouse = True And (Bindings.MouseInput = Trapping)
            End If
        End If
                
        lX = mX
        lY = mX
        lZ = mZ
        
    End If

    Exit Sub
pausing:
    Err.Clear




    Err.Clear
    DoPauseGame
End Sub

Public Sub MouseLook(ByVal X As Integer, ByVal Y As Integer, ByVal Z As Integer)

    Dim cnt As Long
    
    If Z < 0 Then
        If Bindings(DIK_PGDN) = "" Then
            ResetIdle
            For cnt = 0 To -Z
                Player.ZoomOut
            Next
        End If
    ElseIf Z > 0 Then
        If Bindings(DIK_PGUP) = "" Then
            ResetIdle
            For cnt = 0 To Z
                Player.ZoomIn
            Next
        End If
    End If
    
    
    If Perspective = CameraMode Then
        If Player.CameraIndex > 0 Then
            Player.angle = Cameras(Player.CameraIndex).angle
        End If
    Else
        If X < 0 Then
            If Bindings(DIK_LEFT) = "" Then
                ResetIdle
                For cnt = 0 To (-X * MouseSensitivity)
                    Player.LookLeft
                Next
            End If
        ElseIf X > 0 Then
            If Bindings(DIK_RIGHT) = "" Then
                ResetIdle
                For cnt = 0 To (X * MouseSensitivity)
                    Player.LookRight
                Next
            End If
        End If
    End If

    If Y < 0 Then
        If Bindings(DIK_DOWN) = "" Then
            ResetIdle
            For cnt = 0 To (-Y * MouseSensitivity)
                Player.LookDown
            Next
        End If
    ElseIf Y > 0 Then
        If Bindings(DIK_UP) = "" Then
            ResetIdle
            For cnt = 0 To (Y * MouseSensitivity)
                Player.LookUp
            Next
        End If
    End If
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
    
    Set Bindings = New Bindings
    
    Set ConsoleMsgs = New VBA.Collection
    Set HistoryMsgs = New VBA.Collection
    
    Bottom = 0
    
    Vertex(0) = MakeScreen(0, 0, -1, 0, 0)
    Vertex(1) = MakeScreen((frmMain.Width / Screen.TwipsPerPixelX), 0, -1, 1, 0)
    Vertex(2) = MakeScreen(0, 0, -1, 0, 1)
    Vertex(3) = MakeScreen((frmMain.Width / Screen.TwipsPerPixelX), 0, -1, 1, 1)
    
    ConsoleWidth = (frmMain.Width / Screen.TwipsPerPixelX)
    ConsoleHeight = (MaxConsoleMsgs * (frmMain.TextHeight("A") / Screen.TwipsPerPixelY)) + (TextSpace * MaxConsoleMsgs) + TextSpace
    If ConsoleHeight > ((frmMain.Height / Screen.TwipsPerPixelY) \ 2) Then ConsoleHeight = ((frmMain.Height / Screen.TwipsPerPixelY) \ 2)
    
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
    
    If Not Bindings Is Nothing Then
        Dim cnt As Integer
        For cnt = 0 To 255
            Bindings(cnt) = ""
        Next
        Set Bindings = Nothing
    End If
    
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
'        Case "debug"
'            DebugMode = Not DebugMode
'            If DebugMode Then
'                AddMessage "Debug mode enabled."
'            Else
'                AddMessage "Debug mode disabled."
'            End If
        Case "goto"
            
            
            Player.Origin = inArg
        Case "parse"
            If PathExists(inArg, True) Then
                inTmp = ReadFile(inArg)
                If inTmp <> "" Then
                    ParseScript inTmp, , 0
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
            If Not (Perspective = Spectator) Then
                Perspective = Spectator
                AddMessage "Changed to spectate mode."
            Else
                AddMessage "Already in spectate mode."
            End If
        Case "join"
            If (Perspective = Spectator) Then
                Perspective = ThirdPerson
                AddMessage "You've entered the game."
            Else
                AddMessage "Already joined the game."
            End If
        Case "eval"
            Process frmMain.Evaluate(inArg)
        Case "echo"
            inArg = Replace(inArg, "\n", vbCrLf)
            Do Until inArg = ""
                AddMessage RemoveNextArg(inArg, vbCrLf)
            Loop
        Case "fade"
            FadeMessage inArg
        Case "clear"
            ClearText


            
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
            AddMessage "Origin X: " & Round(CSng(Player.Origin.X), 3)
            AddMessage "Origin Y: " & Round(CSng(Player.Origin.Y), 3)
            AddMessage "Origin Z: " & Round(CSng(Player.Origin.Z), 3)
            AddMessage "Distance: " & Round(CSng(DistanceEx(Player.Origin, MakePoint(0, 0, 0))), 3)
            AddMessage "Angle: " & Round(CSng(Player.angle), 3)
            AddMessage "Pitch: " & Round(CSng(Player.Pitch), 3)
'        Case "credits"
'            ShowCredits = Not ShowCredits
'        Case "showcredits"
'            ShowCredits = True
'        Case "hidecredits"
'            ShowCredits = False
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
            If PathExists(AppPath & "Levels\" & inArg & ".vbx", True) Then

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

            ElseIf PathExists(AppPath & "Levels\" & CurrentLoadedLevel & ".vbx", True) Then
                CleanupLand
                CleanupMove
                CreateMove
                CreateLand
                AddMessage "Level Reloaded."
            Else
                AddMessage "Invalid Level - [" & AppPath & "Levels\" & inArg & ".vbx" & "]"
            End If


        Case "load"
            If inArg = "" Then
                If EditFileName = "" Then
                    AddMessage "No file is loaded, use ""LOAD <name>"" to load one."
                Else
                    AddMessage "File loaded [" & AppPath & "Levels\" & EditFileName & ".vbx" & "]"
                End If
            Else
                If PathExists(AppPath & "Levels\" & inArg & ".vbx", True) Then
                    EditFileName = inArg
                    EditFileData = ReadFile(AppPath & "Levels\" & inArg & ".vbx")
                    AddMessage "File loaded [" & AppPath & "Levels\" & inArg & ".vbx" & "]"
                Else
                    AddMessage "File not found [" & AppPath & "Levels\" & inArg & ".vbx" & "]"
                End If
            End If
        Case "view"
            If PathExists(AppPath & "Levels\" & EditFileName & ".vbx", True) Then

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
                AddMessage "File not found [" & AppPath & "Levels\" & EditFileName & ".vbx" & "]"
            End If
        Case "lines"
            If PathExists(AppPath & "Levels\" & EditFileName & ".vbx", True) Then
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
                AddMessage "File not found [" & AppPath & "Levels\" & EditFileName & ".vbx" & "]"
            End If
        Case "edit"
            If PathExists(AppPath & "Levels\" & EditFileName & ".vbx", True) Then
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
                AddMessage "File not found [" & AppPath & "Levels\" & EditFileName & ".vbx" & "]"
            End If
        Case "save"
            If PathExists(AppPath & "Levels\" & EditFileName & ".vbx", True) Then
                WriteFile AppPath & "Levels\" & EditFileName & ".vbx", EditFileData
                AddMessage "Saved data file [" & AppPath & "Levels\" & EditFileName & ".vbx" & "]"
            ElseIf EditFileName = "" Then
                AddMessage "File not loaded."
            Else
                AddMessage "File not found [" & AppPath & "Levels\" & EditFileName & ".vbx" & "]"
            End If
        Case ""
        Case Else
            frmMain.ExecuteStatement inCmd & IIf(inArg <> "", " " & inArg, "")
    End Select
End Sub


