Attribute VB_Name = "modInput"

#Const modInput = -1
Option Explicit
'TOP DOWN
Option Private Module


Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type Msg
    hwnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

' Virtual Keys, Standard Set
Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2
Public Const VK_CANCEL = &H3
Public Const VK_MBUTTON = &H4             '  NOT contiguous with L RBUTTON

Public Const VK_BACK = &H8
Public Const VK_TAB = &H9

Public Const VK_CLEAR = &HC
Public Const VK_RETURN = &HD

Public Const VK_SHIFT = &H10
Public Const VK_CONTROL = &H11
Public Const VK_MENU = &H12
Public Const VK_PAUSE = &H13
Public Const VK_CAPITAL = &H14

Public Const VK_ESCAPE = &H1B

Public Const VK_SPACE = &H20
Public Const VK_PRIOR = &H21
Public Const VK_NEXT = &H22
Public Const VK_END = &H23
Public Const VK_HOME = &H24
Public Const VK_LEFT = &H25
Public Const VK_UP = &H26
Public Const VK_RIGHT = &H27
Public Const VK_DOWN = &H28
Public Const VK_SELECT = &H29
Public Const VK_PRINT = &H2A
Public Const VK_EXECUTE = &H2B
Public Const VK_SNAPSHOT = &H2C
Public Const VK_INSERT = &H2D
Public Const VK_DELETE = &H2E
Public Const VK_HELP = &H2F

' VK_A thru VK_Z are the same as their ASCII equivalents: 'A' thru 'Z'
' VK_0 thru VK_9 are the same as their ASCII equivalents: '0' thru '9'

Public Const VK_NUMPAD0 = &H60
Public Const VK_NUMPAD1 = &H61
Public Const VK_NUMPAD2 = &H62
Public Const VK_NUMPAD3 = &H63
Public Const VK_NUMPAD4 = &H64
Public Const VK_NUMPAD5 = &H65
Public Const VK_NUMPAD6 = &H66
Public Const VK_NUMPAD7 = &H67
Public Const VK_NUMPAD8 = &H68
Public Const VK_NUMPAD9 = &H69
Public Const VK_MULTIPLY = &H6A
Public Const VK_ADD = &H6B
Public Const VK_SEPARATOR = &H6C
Public Const VK_SUBTRACT = &H6D
Public Const VK_DECIMAL = &H6E
Public Const VK_DIVIDE = &H6F
Public Const VK_F1 = &H70
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B
Public Const VK_F13 = &H7C
Public Const VK_F14 = &H7D
Public Const VK_F15 = &H7E
Public Const VK_F16 = &H7F
Public Const VK_F17 = &H80
Public Const VK_F18 = &H81
Public Const VK_F19 = &H82
Public Const VK_F20 = &H83
Public Const VK_F21 = &H84
Public Const VK_F22 = &H85
Public Const VK_F23 = &H86
Public Const VK_F24 = &H87

Public Const VK_NUMLOCK = &H90
Public Const VK_SCROLL = &H91

'
'   VK_L VK_R - left and right Alt, Ctrl and Shift virtual keys.
'   Used only as parameters to GetAsyncKeyState() and GetKeyState().
'   No other API or message will distinguish left and right keys in this way.
'  /
Public Const VK_LSHIFT = &HA0
Public Const VK_RSHIFT = &HA1
Public Const VK_LCONTROL = &HA2
Public Const VK_RCONTROL = &HA3
Public Const VK_LMENU = &HA4
Public Const VK_RMENU = &HA5

Public Const VK_ATTN = &HF6
Public Const VK_CRSEL = &HF7
Public Const VK_EXSEL = &HF8
Public Const VK_EREOF = &HF9
Public Const VK_PLAY = &HFA
Public Const VK_ZOOM = &HFB
Public Const VK_NONAME = &HFC
Public Const VK_PA1 = &HFD
Public Const VK_OEM_CLEAR = &HFE

Private VirtualMouse As POINTAPI

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long

Private Type KeyState
    VKState As Integer
    VKToggle As Boolean
    VKPressed As Boolean
    VKLatency As Single
 End Type

Private VirtualKeys(0 To 255) As KeyState

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Function GetShiftCtrlAlt() As Integer
    Dim iKeys As Integer
    
    Const VK_SHIFT As Long = &H10
    Const VK_CONTROL As Long = &H11
    Const VK_ALT As Long = &H12
    
    If GetAsyncKeyState(VK_SHIFT) <> 0 Then iKeys = iKeys + 1
    If GetAsyncKeyState(VK_CONTROL) <> 0 Then iKeys = iKeys + 1
    If GetAsyncKeyState(VK_ALT) <> 0 Then iKeys = iKeys + 1
    
    GetShiftCtrlAlt = iKeys
End Function

Public Sub ReadyInput()
    Dim cnt As Long
    For cnt = 0 To 255
        VirtualKeys(cnt).VKState = GetKeyState(cnt)
    Next
End Sub
Public Function InputLoop() As Long
    Dim cnt As Long
    Dim ks As Long
    For cnt = 0 To 255
        ks = GetKeyState(cnt)
        If (Not (VirtualKeys(cnt).VKState = ks)) Then
            VirtualKeys(cnt).VKState = ks
            If (VirtualKeys(cnt).VKLatency = 0) Then
                VirtualKeys(cnt).VKLatency = Timer
            End If
            VirtualKeys(cnt).VKPressed = Not VirtualKeys(cnt).VKPressed
            If VirtualKeys(cnt).VKPressed = False Then
                VirtualKeys(cnt).VKToggle = Not VirtualKeys(cnt).VKToggle
            ElseIf VirtualKeys(cnt).VKToggle Then
                VirtualKeys(cnt).VKToggle = VirtualKeys(cnt).VKToggle Xor (Not VirtualKeys(cnt).VKPressed)
            End If
            If Not VirtualKeys(cnt).VKPressed Then VirtualKeys(cnt).VKLatency = 0
        Else
            If (VirtualKeys(cnt).VKLatency <> 0) Then
                If ((Timer - -VirtualKeys(cnt).VKLatency) > 0.15) And (VirtualKeys(cnt).VKLatency < 0) Then
                    VirtualKeys(cnt).VKPressed = True
                    VirtualKeys(cnt).VKLatency = VirtualKeys(cnt).VKLatency + -0.2
                ElseIf ((Timer - VirtualKeys(cnt).VKLatency) > 0.15) And (VirtualKeys(cnt).VKLatency > 0) Then
                    VirtualKeys(cnt).VKPressed = True
                    VirtualKeys(cnt).VKLatency = VirtualKeys(cnt).VKLatency - 0.2
                End If
            Else
                VirtualKeys(cnt).VKLatency = 0
            End If
        End If
    Next
    InputLoop = GetCursorPos(VirtualMouse)
End Function
Public Function EndOfInput() As Boolean
    EndOfInput = VirtualKeys(VK_ESCAPE).VKPressed
End Function
Public Function Toggled(ByVal vkCode As Long) As Boolean
    Toggled = VirtualKeys(vkCode).VKToggle
End Function
Public Function Pressed(ByVal vkCode As Long) As Boolean
    Pressed = VirtualKeys(vkCode).VKPressed
End Function
Public Function MouseXY() As POINTAPI
    MouseXY = VirtualMouse
End Function
