Attribute VB_Name = "modAPI"
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
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private VBGTray As NOTIFYICONDATA
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private IsIconInTray As Boolean
Private Const SysTrayToolTip = "Max-FTP - Click for a menu"

Public Declare Function DestroyMenu Lib "user32" (ByRef hMenu As Long) As Long

Public Const MAX_PATH = 260

Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_SYSICONINDEX = &H4000
Public Const SHGFI_LARGEICON = &H0
Public Const SHGFI_SMALLICON = &H1
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_TYPENAME = &H400
Public Const SHGFI_USEFILEATTRIBUTES = &H10
Public Const FILE_ATTRIBUTE_NORMAL = &H80

Public Const ILD_TRANSPARENT = &H1

Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Public Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long

Public Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hDCDest As Long, ByVal X As Long, ByVal Y As Long, ByVal flags As Long) As Long

Public Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociateIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Public Const LARGE_ICON As Integer = 32
Public Const SMALL_ICON As Integer = 16
Public Const DI_NORMAL = 3
Public Declare Function DrawIconEx Lib "user32.dll" _
    (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, _
    ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, _
    ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, _
    ByVal diFlags As Long) As Long
    
Type POINTAPI
      X As Long
      Y As Long
End Type
Global MousePos As POINTAPI
Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3

Public Const GW_CHILD = 5
Public Const GW_MAX = 5
Declare Function GetWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function SetFocus Lib "user32.dll" (ByVal hWnd As Long) As Long
Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function GetTopWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Declare Function GetActiveWindow Lib "user32.dll" () As Long

Public Function SetFocusOnView(ByVal frm As frmFTPClientGUI, ByVal Index As Integer) As Long
    Dim hWnd As Long
    Dim className As String * 100
    Dim classLen As Long
    
    hWnd = GetWindow(frm.userGUI(Index).hWnd, GW_CHILD)
    
    hWnd = GetWindow(hWnd, GW_HWNDFIRST)
    classLen = GetClassName(hWnd, className, 100)
    
    Do Until Trim(LCase(Left(className, 8))) = "listview"
        hWnd = GetWindow(hWnd, GW_HWNDNEXT)
        classLen = GetClassName(hWnd, className, 100)
    Loop
        
    SetFocus hWnd

End Function

Public Sub MaxInTray()
    
    On Error Resume Next
    VBGTray.cbSize = Len(VBGTray)
    VBGTray.hWnd = frmMain.hWnd
    VBGTray.uID = vbNull
    VBGTray.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    VBGTray.uCallbackMessage = WM_MOUSEMOVE
    VBGTray.hIcon = frmMain.Icon
    VBGTray.szTip = SysTrayToolTip & vbNullChar
    If IsIconInTray Then
        Call Shell_NotifyIcon(NIM_MODIFY, VBGTray)
    Else
        IsIconInTray = True
        Call Shell_NotifyIcon(NIM_ADD, VBGTray)
    End If
    If Err Then Err.Clear
    On Error GoTo 0
    IsIconInTray = True

End Sub
Public Sub MaxOutTray()
    
    On Error Resume Next
    If IsIconInTray Then
        VBGTray.cbSize = Len(VBGTray)
        VBGTray.hWnd = frmMain.hWnd
        VBGTray.uID = vbNull
        Call Shell_NotifyIcon(NIM_DELETE, VBGTray)
        IsIconInTray = False
    End If
    If Err Then Err.Clear
    On Error GoTo 0

End Sub
Public Sub MaxMouseOverTray(X)
    
    Dim lngMsg As Long
'    Dim blnFlag As Boolean
    lngMsg = X / Screen.TwipsPerPixelX
'    If blnFlag = False Then
'        blnFlag = True
        
        Select Case lngMsg
            Case WM_LBUTTONUP
                frmMain.InitSysTrayMenu
            Case WM_RBUTTONUP
                frmMain.InitSysTrayMenu
        End Select
'        blnFlag = False
'    End If

End Sub



