Attribute VB_Name = "modCmds"

#Const modCmds = -1
Option Explicit
'TOP DOWN

Option Compare Binary

Private Const MaxConsoleMsgs = 20
Private Const MaxHistoryMsgs = 10

Private ConsoleMsgs As Collection
Private HistoryMsgs As Collection
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

Private lX As Integer
Private lY As Integer
Private lZ As Integer
Private norepeat As String

Private FadeTime As Long
Private FadeText As String

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
            GetBindingText = "NEXTTRACK"
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
            GetBindingText = "NEXT"
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


Public Sub FadeMessage(ByVal txt As String)
    FadeTime = Timer
    txt = Replace(txt, "\n", vbCrLf)
    FadeText = txt
    AddMessage txt
End Sub

Public Function Row(ByVal num As Long) As Long
    Row = ((TextHeight \ VB.Screen.TwipsPerPixelY) * num) + (2 * num)
End Function

Public Function MakeScreen(ByVal X As Single, ByVal Y As Single, ByVal z As Single, Optional ByVal tu As Single = 0, Optional ByVal tv As Single = 0) As MyScreen
    MakeScreen.X = X
    MakeScreen.Y = Y
    MakeScreen.z = z
    MakeScreen.rhw = 1
    MakeScreen.clr = D3DColorARGB(255, 255, 255, 255)
    MakeScreen.tu = tu
    MakeScreen.tv = tv
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

Private Function Toggled(ByVal vkCode As Long) As Boolean
    Toggled = KeyState(vkCode).VKToggle
End Function
Private Function Pressed(ByVal vkCode As Long) As Boolean
    Pressed = KeyState(vkCode).VKPressed
End Function

Public Sub InputScene(ByRef UserControl As Macroscopic)

    On Error GoTo pausing

    DIKeyBoardDevice.GetDeviceStateKeyboard DIKEYBOARDSTATE

    ConsoleInput UserControl, DIKEYBOARDSTATE

    Dim cnt As Long
    Dim uses(0 To 255) As Boolean
    
    If ((GetActiveWindow = UserControl.Parent.hwnd) Or (GetActiveWindow = frmMain.hwnd)) Then
    
        If (FullScreen And Not TrapMouse) Then TrapMouse = True And (Bindings.Controller = Trapping)

        If DIKEYBOARDSTATE.Key(DIK_ESCAPE) Then
            If (Not TogglePress1 = DIK_ESCAPE) Then
                TogglePress1 = DIK_ESCAPE
    
                If ((Not FullScreen) And TrapMouse) Then
                    TrapMouse = False
                ElseIf (FullScreen Or (Not FullScreen And Not TrapMouse)) And (Bindings.Controller = Trapping) Then
                    StopGame = True
                End If
                
            End If
        ElseIf DIKEYBOARDSTATE.Key(DIK_F1) Then
            If (Not TogglePress1 = DIK_F1) Then
                TogglePress1 = DIK_F1
                ShowSetup = True
            End If
        ElseIf DIKEYBOARDSTATE.Key(DIK_F2) Then
            If (Not TogglePress1 = DIK_F2) Then
                TogglePress1 = DIK_F2
                ShowStats = Not ShowStats
            End If
            
        ElseIf DIKEYBOARDSTATE.Key(DIK_LALT) Or DIKEYBOARDSTATE.Key(DIK_RALT) Then
            If DIKEYBOARDSTATE.Key(DIK_TAB) Then
                If (Not TogglePress1 = DIK_TAB) Then
                    TogglePress1 = DIK_TAB
                    TrapMouse = False

                    If FullScreen Then
                        UserControl.Parent.WindowState = 1
                        UserControl.PauseRendering
                    End If
                    
                End If
            End If
        ElseIf DIKEYBOARDSTATE.Key(DIK_GRAVE) Then
            If (Not TogglePress1 = DIK_GRAVE) Then
                TogglePress1 = DIK_GRAVE

                ConsoleToggle
                
            End If
        ElseIf Not (TogglePress1 = 0) Then
            TogglePress1 = 0
        End If
    
        Dim rec  As RECT
        Dim mloc As modDecs.POINTAPI
        GetCursorPos mloc
        GetWindowRect UserControl.hwnd, rec
   
        Dim mX As Integer
        Dim mY As Integer
        Dim mZ As Integer
        DIMouseDevice.GetDeviceStateMouse DIMOUSESTATE
        mX = DIMOUSESTATE.lX
        mY = DIMOUSESTATE.lY
        mZ = DIMOUSESTATE.lZ
        
        If ((TrapMouse Or FullScreen) And (Bindings.Controller = Trapping)) Or _
            ((Bindings.Controller = Hidden) And ((mloc.X >= rec.Left) And (mloc.X <= rec.Right) And (mloc.Y >= rec.Top) And (mloc.Y <= rec.Bottom))) Then
            If (Not (VB.Screen.MousePointer = 99)) Then
                VB.Screen.MousePointer = 99
                Set VB.Screen.MouseIcon = PictureFromByteStream(LoadResData(2, "CUSTOM"))
            End If
        Else
            If (Not (VB.Screen.MousePointer = 0)) Then VB.Screen.MousePointer = 0
        End If
        
        If ((TrapMouse And (Not ConsoleVisible)) And (Bindings.Controller = Trapping)) Or _
            (((Not TrapMouse) And (Not ConsoleVisible)) And (Bindings.Controller = Hidden)) Or _
            ((Bindings.Controller = Visual) And ((mloc.X >= rec.Left) And (mloc.X <= rec.Right) _
            And (mloc.Y >= rec.Top) And (mloc.Y <= rec.Bottom))) Then
        
            For cnt = 0 To 255
                If DIKEYBOARDSTATE.Key(cnt) Then
                    If Not Bindings(cnt) = "" Then
                        uses(cnt) = True
                        frmMain.RunEvent Bindings(cnt)
                    End If
                End If
            Next
                
            If ((mZ < 0) And (Bindings(DIK_PGDN) <> "")) And (Not uses(DIK_PGDN)) Then
                uses(DIK_PGDN) = True
                For cnt = mZ To 0
                    frmMain.RunEvent Bindings(DIK_PGDN)
                Next
            ElseIf ((mZ > 0) And (Bindings(DIK_PGUP) <> "")) And (Not uses(DIK_PGUP)) Then
                uses(DIK_PGDN) = True
                For cnt = 0 To mZ
                    frmMain.RunEvent Bindings(DIK_PGUP)
                Next
            End If
            
            If ((mX < 0) And (Bindings(DIK_LEFT) <> "")) And (Not uses(DIK_LEFT)) Then
                uses(DIK_LEFT) = True
                For cnt = (mX * MouseSensitivity) To 0
                    frmMain.RunEvent Bindings(DIK_LEFT)
                Next
            ElseIf ((mX > 0) And (Bindings(DIK_RIGHT) <> "")) And (Not uses(DIK_RIGHT)) Then
                uses(DIK_RIGHT) = True
                For cnt = 0 To (mX * MouseSensitivity)
                    frmMain.RunEvent Bindings(DIK_RIGHT)
                Next
            End If

            If ((mY < 0) And (Bindings(DIK_DOWN) <> "")) And (Not uses(DIK_DOWN)) Then
                uses(DIK_DOWN) = True
                For cnt = (mY * MouseSensitivity) To 0
                    frmMain.RunEvent Bindings(DIK_DOWN)
                Next
            ElseIf ((mY > 0) And (Bindings(DIK_UP) <> "")) And (Not uses(DIK_UP)) Then
                uses(DIK_UP) = True
                For cnt = 0 To (mY * MouseSensitivity)
                    frmMain.RunEvent Bindings(DIK_UP)
                Next
            End If

            If (DIMOUSESTATE.Buttons(0) And (Bindings(DIK_LCONTROL) <> "")) And (Not uses(DIK_LCONTROL)) Then
                frmMain.RunEvent Bindings(DIK_LCONTROL)
            End If
            If (DIMOUSESTATE.Buttons(1) And (Bindings(DIK_LALT) <> "")) And (Not uses(DIK_LALT)) Then
                frmMain.RunEvent Bindings(DIK_LALT)
            End If

            If (DIMOUSESTATE.Buttons(2) And (Bindings(DIK_RCONTROL) <> "")) And (Not uses(DIK_RCONTROL)) Then
                frmMain.RunEvent Bindings(DIK_RCONTROL)
            End If
            If (DIMOUSESTATE.Buttons(3) And (Bindings(DIK_RALT) <> "")) And (Not uses(DIK_RALT)) Then
                frmMain.RunEvent Bindings(DIK_RALT)
            End If

            If (Bindings.Controller = Trapping) Then
                SetCursorPos (UserControl.Parent.Left / VB.Screen.TwipsPerPixelX) + UserControl.Left + (UserControl.Width / 2), (UserControl.Parent.Top / VB.Screen.TwipsPerPixelY) + UserControl.Top + (UserControl.Height / 2)
            End If
            
        End If
    
        If ((mloc.X > rec.Left) And (mloc.X < rec.Right)) And ((mloc.Y > rec.Top) And (mloc.Y < rec.Bottom)) Then
            If DIMOUSESTATE.Buttons(0) Then
                TrapMouse = True And (Bindings.Controller = Trapping)
            End If
    
        End If

        lX = mX
        lY = mX
        lZ = mZ
        
    End If

    Exit Sub
pausing:
    Err.Clear
    'DoPauseGame UserControl
    UserControl.PauseRendering
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

Public Sub RenderCmds(ByRef UserControl As Macroscopic)

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
    
    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_LINEAR

    DDevice.SetVertexShader FVF_SCREEN
    DDevice.SetRenderState D3DRS_ZENABLE, False
    DDevice.SetRenderState D3DRS_LIGHTING, False
    DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE

    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, False
    
'    If Billboards.Count > 0 Then
'
'        Dim e As Billboard
'
'        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
'        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
'
'        DDevice.SetMaterial GenericMaterial
'
'        For Each e In Billboards
'
'            If ((e.Form And TwoDimensions) = TwoDimensions) Then
'
'                If e.Visible Then
'                    If e.Transparent Then
'                        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
'                        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
'                    ElseIf e.Translucent Then
'                        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTALPHA
'                        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_SRCALPHA
'                    Else
'                        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
'                        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
'                    End If
'
'                    DDevice.SetTexture 0, Files(Faces(e.FaceIndex).Images(e.AnimatePoint)).Data
'                    DDevice.SetTexture 1, Files(Faces(e.FaceIndex).Images(e.AnimatePoint)).Data
'                    DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Faces(e.FaceIndex).Screen2D(0), LenB(Faces(e.FaceIndex).Screen2D(0))
'
'                End If
'            End If
'        Next
'
'    End If


    If ConsoleVisible Then
    
        Dim lastColor As Long
        lastColor = TextColor
        TextColor = D3DColorARGB(255, 0, 0, 0)
        
        DDevice.SetTexture 0, Backdrop
        DDevice.SetTexture 1, Nothing
        DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0), LenB(Vertex(0))
    
        If Len(CommandLine) > 0 Then
            DrawText Replace(CommandLine, vbTab, "     "), TextSpace + 2, Bottom - (TextHeight / VB.Screen.TwipsPerPixelY) - TextSpace + 2
        End If
        
        Static DrawCursor As Double
        Static DrawBlink As Boolean
        If DrawCursor = 0 Or CDbl(((GetTimer * 1000) - DrawCursor)) >= 2000 Then
            DrawCursor = GetTimer
            DrawBlink = Not DrawBlink
            
            If DrawBlink Then DrawText String(CursorPos + (CountWord(Left(CommandLine, CursorPos), vbTab) * 4), " ") & "_", TextSpace + 2, Bottom - (TextHeight / VB.Screen.TwipsPerPixelY) - TextSpace + 2

        End If
        
        If ConsoleMsgs.Count > 0 Then
            Dim ConsoleMsgY As Long
            ConsoleMsgY = (Bottom - (((TextHeight / VB.Screen.TwipsPerPixelY) + TextSpace) * 2))
            Dim cnt As Integer
            For cnt = ConsoleMsgs.Count To 1 Step -1
                DrawText ConsoleMsgs(cnt), TextSpace + 2, ConsoleMsgY - ((ConsoleMsgs.Count - cnt) * (((TextHeight / VB.Screen.TwipsPerPixelY) + TextSpace)))
            Next
        End If
        
        TextColor = lastColor

    Else
        If ShowStats Then DrawText GetStats, 10, 10

        If DrawCount > 0 Then
            For cnt = 1 To DrawCount
                If Draws(1, cnt) <> "" Then DrawText CStr(Draws(1, cnt)), CSng(Draws(2, cnt)), CSng(Draws(3, cnt))
            Next
        End If

    End If

    If Not (FadeText = "") Then
        If (Timer - FadeTime) >= 6 Then
            FadeText = ""
        Else
            DrawText FadeText, ((frmMain.ScaleWidth / VB.Screen.TwipsPerPixelX) / 2) - ((frmMain.TextWidth(FadeText) / VB.Screen.TwipsPerPixelX) / 2), ((frmMain.ScaleHeight / VB.Screen.TwipsPerPixelY) / 2) - (TextHeight / 2)
        End If
    End If
    
    DDevice.SetRenderState D3DRS_ZENABLE, 1
    DDevice.SetRenderState D3DRS_LIGHTING, 1
        
    DDevice.SetPixelShader PixelShaderDefault

End Sub

Public Function AddMessage(ByVal Message As String)

    If ConsoleMsgs.Count > MaxConsoleMsgs Then
        ConsoleMsgs.Remove 1
    End If
    ConsoleMsgs.Add Message
    Debug.Print Message
    
End Function
Public Sub ConsoleInput(ByRef UserControl As Macroscopic, ByRef kState As DIKEYBOARDSTATE)

    If ConsoleVisible And (Not (Shift < 0)) And ((GetActiveWindow = UserControl.Parent.hwnd) Or (GetActiveWindow = frmMain.hwnd)) Then
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
                    
                    Process CommandLine, UserControl
                    
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
                
                CursorX = (frmMain.TextWidth(Left(CommandLine, CursorPos)) / VB.Screen.TwipsPerPixelX)
                CursorWidth = (frmMain.TextWidth(Mid(CommandLine, CursorPos + 1, 1)) / VB.Screen.TwipsPerPixelX)
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

    IdleInput = Timer - 4

    Set ConsoleMsgs = New Collection
    Set HistoryMsgs = New Collection

    Bottom = 0

    Dim measure As Single

    measure = (frmMain.Width / VB.Screen.TwipsPerPixelX) + ((VB.Screen.Width / VB.Screen.TwipsPerPixelX) - (frmMain.Width / VB.Screen.TwipsPerPixelX))
    
    Vertex(0) = MakeScreen(0, 0, -1, 0, 0)
    Vertex(1) = MakeScreen(measure, 0, -1, 1, 0)
    Vertex(2) = MakeScreen(0, 0, -1, 0, 1)
    Vertex(3) = MakeScreen(measure, 0, -1, 1, 1)

    ConsoleWidth = measure
    ConsoleHeight = (MaxConsoleMsgs * (frmMain.TextHeight("A") / VB.Screen.TwipsPerPixelY)) + (TextSpace * MaxConsoleMsgs) + TextSpace
    
    measure = (frmMain.Height / VB.Screen.TwipsPerPixelY) + ((VB.Screen.Height / VB.Screen.TwipsPerPixelY) - (frmMain.Height / VB.Screen.TwipsPerPixelY))
     
    If ConsoleHeight > (measure \ 2) Then ConsoleHeight = (measure \ 2)
        
    InitKeys
    
    Set Backdrop = LoadTextureRes(LoadResData(1, "CUSTOM"))
    
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
    ClearText

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


Public Sub Process(ByVal inArg As String, Optional ByRef UserControl As Macroscopic)


    Dim inX As Single
    Dim inY As Single
    
    Dim inCmd As String
    
    inCmd = RemoveNextArg(inArg, " ")
    If Left(inCmd, 1) = "/" Then inCmd = Mid(inCmd, 2)
    
    Select Case Trim(LCase(inCmd))

        Case "reset"
            ResetGame = True
                        
        Case "exit", "quit", "close"
            StopGame = True

        Case "echo"
            inArg = Replace(inArg, "\n", vbCrLf)
            Do Until inArg = ""
                AddMessage RemoveNextArg(inArg, vbCrLf)
            Loop
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
            PrintText inArg, (((frmMain.ScaleWidth / VB.Screen.TwipsPerPixelX) - (TextSpace * 2)) / ColumnCount) * inX, Row(inY)
        Case "help", "cmdlist", "?", "--?"
            Select Case LCase(inArg)
                Case Else
                    
                    AddMessage ""
                    AddMessage "Console Commands:"
                    AddMessage "   EXIT (completely exit out of the game)"
                    AddMessage "   ECHO<Text> (Displays Text in the console)"
                    AddMessage "   FADE<Text> (Displays Text in center that goes away)"
                    AddMessage "   DRAW<X><Y><Text> (Draw Text on the screen at X,Y)"
                    AddMessage "   PRINT<Row><Col><Text> (Like DRAW but use row,col)"
                    AddMessage "   CLEAR (Clear all text being drawn by DRAW and PRINT)"

            End Select

        Case ""
        Case Else
            AddMessage "Unknown command."
    End Select
End Sub




'Public Function IntentJesus(ByVal WholeTotal As Single, ByVal UnitMeasure As Single, Optional ByVal PercentWhole As Single = 100, Optional ByVal PercentUnit As Single = 100) As Single
'    'returns the ratio percentage to a full wholetotal or unitmeasures
'    IntentJesus = (PercentUnit / WholeTotal * UnitMeasure / PercentWhole)
'End Function
'Public Function InventJesus(ByVal WholeTotal As Single, ByVal UnitMeasure As Single, Optional ByVal PercentWhole As Single = 100, Optional ByVal PercentUnit As Single = 100) As Single
'    InventJesus = Sqr(PercentUnit ^ 2 / WholeTotal ^ 2 * UnitMeasure ^ 2 / PercentWhole ^ 2)
'End Function
'Public Function InvertJesus(ByVal WholeTotal As Single, ByVal UnitMeasure As Single, Optional ByVal PercentWhole As Single = 100, Optional ByVal PercentUnit As Single = 100) As Single
'    InvertJesus = Sqr((PercentUnit ^ 3) / (WholeTotal ^ 2) * (UnitMeasure ^ 3) / (PercentWhole ^ 4))
'End Function
'Public Function InnateJesus(ByVal WholeTotal As Single, ByVal UnitMeasure As Single) As Single
'    InnateJesus = (Sqr((IntentJesus(WholeTotal, UnitMeasure) ^ 3) / (InvertJesus(WholeTotal, UnitMeasure) ^ 2) * _
'                    (InvertJesus(WholeTotal, UnitMeasure) ^ 3) / (IntentJesus(WholeTotal, UnitMeasure) ^ 4)) + _
'                    Sqr((IntentJesus(UnitMeasure, WholeTotal) ^ 3) / (InvertJesus(UnitMeasure, WholeTotal) ^ 2) * _
'                    (InvertJesus(UnitMeasure, WholeTotal) ^ 3) / (IntentJesus(UnitMeasure, WholeTotal) ^ 4))) / 2
'End Function
'Public Function InnertJesus(ByVal WholeTotal As Single, ByVal UnitMeasure As Single) As Single
'    InnertJesus = Sqr((InventJesus(WholeTotal, UnitMeasure) ^ 3) / (InventJesus(WholeTotal, UnitMeasure) ^ 2) * _
'                    (InventJesus(WholeTotal, UnitMeasure) ^ 3) / (InventJesus(WholeTotal, UnitMeasure) ^ 4))
'End Function



'Public Sub DrawPointer(ByRef UserControl As Macroscopic, ByVal X As Single, ByVal Y As Single)
'    DDevice.SetVertexShader FVF_SCREEN
'
'    DDevice.SetRenderState D3DRS_ZENABLE, False
'    DDevice.SetRenderState D3DRS_LIGHTING, False
'    DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
'    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
'
'    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
'    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
'    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
'    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
'
''    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
''    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
''    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
''    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_LINEAR
'
'    DDevice.SetMaterial GenericMat
'    If DDevice.GetRenderState(D3DRS_AMBIENT)<> RGB(255, 255, 255) Then
'        DDevice.SetRenderState D3DRS_AMBIENT, RGB(255, 255, 255)
'    End If
'
'    DDevice.SetTexture 0, CircleImgText
'    DDevice.SetTexture 1, CircleImgText
'
'    Dim rec As RECT
'    GetWindowRect UserControl.hWnd, rec
'
'  '  x = x + rec.left
'  '  y = rec.top + InvertNum(y, rec.Bottom - rec.top)
'  '  x = ((frmMain.Width / VB.Screen.TwipsPerPixelX) / 2)
'  '  y = ((frmMain.Height / VB.Screen.TwipsPerPixelY) / 2)
'
'    CircleImgVert(0) = MakeScreen(X - 5, Y + 5, -1, 1, 1)
'    CircleImgVert(1) = MakeScreen(X + 5, Y + 5, -1, 0, 1)
'    CircleImgVert(2) = MakeScreen(X - 5, Y - 5, -1, 1, 0)
'    CircleImgVert(3) = MakeScreen(X + 5, Y - 5, -1, 0, 0)
'
'    DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, CircleImgVert(0), LenB(CircleImgVert(0))
'End Sub


'Public Sub RenderInfo(ByRef UserControl As Macroscopic)
'
'
''    DrawPointer UserControl, frmMain.LastX, frmMain.LastY
'
'    MainFont.Begin
'
'
''    Dim vDir As D3DVECTOR
''    Dim vIntersect As D3DVECTOR
''
''    MouseX = Tan(FOV / 2) * (frmMain.LastX / (frmMain.ScaleWidth / 2) - 1) / ASPECT
''    MouseY = Tan(FOV / 2) * (1 - frmMain.LastY / (frmMain.ScaleHeight / 2))
'
''    Dim p1 As D3DVECTOR 'StartPoint on the nearplane
''    Dim p2 As D3DVECTOR 'EndPoint on the farplane
''
''    p1.x = MouseX * NEAR
''    p1.y = MouseY * NEAR
''    p1.Z = NEAR
''
''    p2.x = MouseX * FAR
''    p2.y = MouseY * FAR
''    p2.Z = FAR
'
'
'
'
''    Dim wholeX As Single
''    Dim wholeY As Single
''    Dim unitX As Single
''    Dim unitY As Single
''
'    Dim X As Single
'    Dim Y As Single
'
'    Dim dy As Single
'    Dim dx As Single
'
'
'    Dim r1 As Single
'    Dim r2 As Single
'
'
'    Dim pt As POINTAPI
'    GetCursorPos pt
'
'    Dim rec As RECT
'    Dim rec2 As RECT
'    GetWindowRect UserControl.hWnd, rec
'
'    GetWindowRect UserControl.Parent.hWnd, rec2
'
'    If ((pt.X >= rec.Left) And (pt.X<= rec.Right)) And ((pt.Y >= rec.Top) And (pt.Y<= rec.Bottom)) Then
'
'        'Screen.MousePointer = 99
'
'        Dim MouseX As Single
'        Dim MouseY As Single
'
'        Dim vDir As D3DVECTOR
'        Dim vIntersect As D3DVECTOR
'
'        MouseX = Tan(FOV / 2) * (pt.X / ((frmMain.ScaleWidth / VB.Screen.TwipsPerPixelX) / 2) - 1) / ASPECT
'        MouseY = Tan(FOV / 2) * (1 - pt.Y / ((frmMain.ScaleHeight / VB.Screen.TwipsPerPixelY) / 2))
'
'        Dim p1 As D3DVECTOR 'StartPoint on the nearplane
'        Dim p2 As D3DVECTOR 'EndPoint on the farplane
'
'        p1.X = MouseX * NEAR
'        p1.Y = MouseY * NEAR
'        p1.z = NEAR
'
'        p2.X = MouseX * FAR
'        p2.Y = MouseY * FAR
'        p2.z = FAR
'
'        'Inverse the view matrix
'        Dim matInverse As D3DMATRIX
'        DDevice.GetTransform D3DTS_VIEW, matView
'
'        D3DXMatrixInverse matInverse, 0, matView
'
'        VectorMatrixMultiply p1, p1, matInverse
'        VectorMatrixMultiply p2, p2, matInverse
'        D3DXVec3Subtract vDir, p2, p1
'
'        'Check if the points hit
'        Dim v1 As D3DVECTOR
'        Dim v2 As D3DVECTOR
'        Dim v3 As D3DVECTOR
'
'        Dim v4 As D3DVECTOR
'        Dim v5 As D3DVECTOR
'        Dim v6 As D3DVECTOR
'
'        Dim pPlane1 As D3DVECTOR4
'
'        Dim cnt As Long
'        v1.X = frmMain.Width '/ Screen.Width)
'        v1.Y = -frmMain.Height ' / Screen.Height)
'        v1.z = 1
'
'        v2.X = -frmMain.Width '/ Screen.Width)
'        v2.Y = -frmMain.Height ' / Screen.Height)
'        v2.z = 1
'
'        v3.X = -frmMain.Width '/ Screen.Width)
'        v3.Y = frmMain.Height '/ Screen.Height)
'        v3.z = 1
'
'        pPlane1 = Create4DPlaneVectorFromPoints(v1, v2, v3)
'
'        Dim c As D3DVECTOR
'        Dim N As D3DVECTOR
'        Dim P As D3DVECTOR
'        Dim V As D3DVECTOR
'
'        Dim hit As Boolean
'        LastMouseSetX = MouseSetX
'        LastMouseSetY = MouseSetY
'
'        MouseSetX = Round(MouseX, 5)
'        MouseSetY = Round(MouseY, 5)
'
'        hit = RayIntersectPlane(pPlane1, p1, vDir, vIntersect)
'
'       ' dx = ((vIntersect.x / Screen.Width) * VB.Screen.TwipsPerPixelX)
'       ' dy = ((vIntersect.y / Screen.Height) * VB.Screen.TwipsPerPixelY)
'
'        LastScreenSetX = ScreenSetX
'        LastScreenSetY = ScreenSetY
'        ScreenSetX = vIntersect.X
'        ScreenSetY = vIntersect.Y
'
'        If hit Then
'
'            'vIntersect.x = (vIntersect.x - rec2.Left)
'           ' vIntersect.y = (-(vIntersect.y + rec2.Top))
'
'           ' x = vIntersect.x
'           ' y = vIntersect.y
'
'
'          '  V = ScreenXYTo3DZ0(x, y)
'
'
'            DrawTextByCoord "A", X, Y
'
'        End If
'
'   ' Else
'   '     Screen.MousePointer = 0
'    End If
'
'   ' Debug.Print hit; X; Y
'
'    MainFont.End
'
'    DDevice.SetVertexShader FVF_SCREEN
'
'    DDevice.SetRenderState D3DRS_ZENABLE, False
'    DDevice.SetRenderState D3DRS_LIGHTING, False
'    DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
'    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
'
'    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
'    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
'    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
'    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
'
'    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
'    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
'    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
'    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_LINEAR
'
'    DDevice.SetMaterial GenericMat
'    If DDevice.GetRenderState(D3DRS_AMBIENT)<> RGB(255, 255, 255) Then
'        DDevice.SetRenderState D3DRS_AMBIENT, RGB(255, 255, 255)
'    End If
'
'  '  If ((pt.x >= rec.Left) And (pt.x<= rec.Right)) And ((pt.y >= rec.Top) And (pt.y<= rec.Bottom)) Then
'
'
'        DDevice.SetTexture 0, CircleImgText
'        DDevice.SetTexture 1, CircleImgText
'
'       ' x = ((frmMain.Width / VB.Screen.TwipsPerPixelX) / 2)
'       ' y = ((frmMain.Height / VB.Screen.TwipsPerPixelY) / 2)
'
'        CircleImgVert(0) = MakeScreen(X - 5, Y + 5, -1, 1, 1)
'        CircleImgVert(1) = MakeScreen(X + 5, Y + 5, -1, 0, 1)
'        CircleImgVert(2) = MakeScreen(X - 5, Y - 5, -1, 1, 0)
'        CircleImgVert(3) = MakeScreen(X + 5, Y - 5, -1, 0, 0)
'
'        DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, CircleImgVert(0), LenB(CircleImgVert(0))
'
' '   End If
'
'
''    MainFont.Begin
''
''    Dim wholeX As Single
''    Dim wholeY As Single
''    Dim unitX As Single
''    Dim unitY As Single
''
''
''    Dim dy As Single
''    Dim dx As Single
''
''    dx = (Player.Location.X / (Screen.Width * VB.Screen.TwipsPerPixelX)) * VB.Screen.TwipsPerPixelX
''    dy = (-Player.Location.Y / (Screen.Height * VB.Screen.TwipsPerPixelY)) * VB.Screen.TwipsPerPixelY
'''
'''
''    Dim pt As POINTAPI
''    GetCursorPos pt
''''
''    Dim rec As RECT
''    GetWindowRect UserControl.hWnd, rec
''''
''
''    pt.X = (pt.X + rec.Left)
''    pt.Y = (pt.Y + rec.Top)
''
''    Dim V As D3DVECTOR
''    V = ScreenXYTo3DZ0((pt.X - rec.Right) + rec.Left, (pt.Y - rec.Bottom) + rec.Top)
'''
'''    DrawTextByCoord "A",
'''
'''    MainFont.End
''
''
''
''
''    DDevice.SetVertexShader FVF_SCREEN
''
''    DDevice.SetRenderState D3DRS_ZENABLE, False
''    DDevice.SetRenderState D3DRS_LIGHTING, False
''    DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
''    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
''
''    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
''    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
''    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
''    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
''
''    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
''    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
''    DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
''    DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_LINEAR
''
''    DDevice.SetMaterial GenericMat
''    If DDevice.GetRenderState(D3DRS_AMBIENT)<> RGB(255, 255, 255) Then
''        DDevice.SetRenderState D3DRS_AMBIENT, RGB(255, 255, 255)
''    End If
''
''    DDevice.SetTexture 0, CircleImgText
''    DDevice.SetTexture 1, CircleImgText
''
''    Dim X As Single
''    Dim Y As Single
''
''    X = frmMain.LastX + dx
''    Y = frmMain.LastY + dy ' InvertNum(v.y + dy, rec.Bottom - rec.top)
''
''    CircleImgVert(0) = MakeScreen(X - 5, Y + 5, -1, 1, 1)
''    CircleImgVert(1) = MakeScreen(X + 5, Y + 5, -1, 0, 1)
''    CircleImgVert(2) = MakeScreen(X - 5, Y - 5, -1, 1, 0)
''    CircleImgVert(3) = MakeScreen(X + 5, Y - 5, -1, 0, 0)
''
''    DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, CircleImgVert(0), LenB(CircleImgVert(0))
'
'
'
'End Sub
'




