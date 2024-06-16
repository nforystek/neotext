Attribute VB_Name = "modTrayIcon"





#Const modTrayIcon = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONUP = &H205
Private Const WM_MOUSEMOVE = &H200
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Declare Function DestroyMenu Lib "user32" (ByRef hMenu As Long) As Long

Private Declare Function ExtractAssociatedIcon Lib "shell32" Alias "ExtractAssociateIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Private Declare Function ExtractIcon Lib "shell32" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Function ExtractIconEx Lib "shell32" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Private Const LARGE_ICON As Integer = 32
Private Const SMALL_ICON As Integer = 16
Private Const DI_NORMAL = 3
Private Declare Function DrawIconEx Lib "user32" _
    (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, _
    ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, _
    ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, _
    ByVal diFlags As Long) As Long
    
Private Type POINTAPI
      x As Long
      Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3

Private Const GW_CHILD = 5
Private Const GW_MAX = 5
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long

Private TrayInfo As NOTIFYICONDATA
Private IsIconInTray As Boolean

Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Any) As Long

Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Public Enum OverwriteTypes
    or_Prompt = 7
    or_Yes = 0
    or_YesToAll = 1
    or_No = 2
    or_NoToAll = 3
    or_Resume = 4
    or_Cancel = 5
    or_AutoAll = 6
End Enum

Public Sub InitSystemTray()
    TrayInfo.cbSize = Len(TrayInfo)
    TrayInfo.hwnd = frmMain.hwnd
    TrayInfo.uID = vbNull
    TrayInfo.uCallbackMessage = WM_MOUSEMOVE
    TrayInfo.hIcon = frmMain.Icon
    TrayInfo.szTip = AppName & " - Click for menu" & vbNullChar
    SetTrayToolTip
End Sub
Public Sub SetTrayToolTip()
    If dbSettings.GetProfileSetting("ViewToolTips") Then
        TrayInfo.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    Else
        TrayInfo.uFlags = NIF_ICON Or NIF_MESSAGE
    End If
    TrayIcon False
    TrayIcon dbSettings.GetProfileSetting("SystemTray")
End Sub

Public Function PerformInTray()

    Dim cnt As Integer
    
    For cnt = 0 To Forms.count - 1

        If IsHiddenForm(Forms(cnt)) Then

            If Forms(cnt).WindowState = 1 Then
                If Not Forms(cnt).Visible = False Then Forms(cnt).Visible = False
            Else
                If Not Forms(cnt).Visible = True Then Forms(cnt).Visible = True
            End If
        
        End If
    Next

End Function
Public Function PerformOutTray() As Boolean
    
    Dim LoadedForms As Integer
    LoadedForms = 0
    
    Dim cnt As Integer

    For cnt = 0 To Forms.count - 1

        If IsHiddenForm(Forms(cnt)) Then

            If Not Forms(cnt).Visible = True Then Forms(cnt).Visible = True

        End If

        If IsUnloadForm(Forms(cnt)) Then

            LoadedForms = LoadedForms + 1

        End If

    Next

    If LoadedForms = 0 Then
        
        ShutDownMaxFTP
        
        PerformOutTray = False
    Else
        PerformOutTray = True
    End If

End Function

Public Sub TrayIcon(ByVal IsInTray As Boolean)

    If (Not IsInTray) And IsIconInTray Then
        Call Shell_NotifyIcon(NIM_DELETE, TrayInfo)
        IsIconInTray = False
    ElseIf IsInTray Then

        Dim exp As Long
        exp = ProcessRunning("explorer.exe")

        If IsIconInTray And (exp = 0) Then
            Call Shell_NotifyIcon(NIM_DELETE, TrayInfo)
            IsIconInTray = False
        ElseIf (Not IsIconInTray) And (exp > 0) Then
            IsIconInTray = (Shell_NotifyIcon(NIM_ADD, TrayInfo) = 1)
        ElseIf IsIconInTray And (exp > 0) Then
            'Call Shell_NotifyIcon(NIM_MODIFY, TrayInfo)
        End If

    End If

End Sub

Public Sub MouseOverTray(x)
    On Error GoTo trayerror
    Dim lngMsg As Long

    lngMsg = x / Screen.TwipsPerPixelX
        
    Select Case lngMsg
        Case WM_LBUTTONUP
            frmMain.InitSysTrayMenu
        Case WM_RBUTTONUP
            frmMain.InitSysTrayMenu
    End Select
    Exit Sub
trayerror:
    Err.Clear
End Sub

