Attribute VB_Name = "modMWheel"
#Const [True] = -1
#Const [False] = 0
#Const modMWheel = -1
Option Explicit
'TOP DOWN
Option Compare Text

Public Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadId _
    As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long

Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent _
    As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
    (ByVal hwnd As Long, ByVal lpClassName As String, _
    ByVal nMaxCount As Long) As Long

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hwnd As Long, ByVal lpString As String, _
    ByVal cch As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias _
    "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) _
    As Long
    
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As _
    Long, ByVal lParam As Long) As Long

Private Declare Function WindowFromPointXY Lib "user32" _
    Alias "WindowFromPoint" (ByVal xPoint As Long, _
    ByVal yPoint As Long) As Long
               
Private Declare Function SystemParametersInfo Lib "user32" _
    Alias "SystemParametersInfoA" _
    (ByVal uAction As Long, _
    ByVal uParam As Long, _
    lpvParam As Any, _
    ByVal fuWinIni As Long) As Long

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        
Private Declare Function WindowFromPoint Lib "user32" (pt As POINTAPI) As Long

Private Declare Function GetWindowInfo Lib "user32" (ByVal hwnd As Long, ByRef pwi As WINDOWINFO) As Boolean

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" Alias "FreeLibraryA" (ByVal hLibrary As Long) As Boolean

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type WINDOWINFO
    cbSize As Long
    rcWindow As RECT
    rcClient As RECT
    dwStyle As Long
    dwExStyle As Long
    cxWindowBorders As Long
    cyWindowBorders As Long
    atomWindowtype As Long
    wCreatorVersion As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MOUSEHOOKSTRUCT
    pt As POINTAPI
    hwnd As Long
    wHitTestCode As Long
    dwExtraInfo As Long
End Type

Private Type MSLLHOOKSTRUCT
    pt As POINTAPI
    mouseData As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Private Const WM_MOUSEWHEEL = &H20A
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONUP = &H205

Private Const MK_LBUTTON = &H1
Private Const MK_MBUTTON = &H10
Private Const MK_RBUTTON = &H2

Private Const WH_MOUSE = 7
Private Const WHEEL_DELTA = 120

Private Const WM_VSCROLL = &H115
Private Const WM_USER = &H400

Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101

Private Const VK_F1 = &H70
Private Const VK_F2 = &H71
Private Const VK_F3 = &H72
Private Const VK_F4 = &H73
Private Const VK_F5 = &H74
Private Const VK_F6 = &H75
Private Const VK_F7 = &H76
Private Const VK_F8 = &H77
Private Const VK_F9 = &H78
Private Const VK_F10 = &H79
Private Const VK_F11 = &H7A
Private Const VK_F12 = &H7B

Private Const GWL_WNDPROC = -4
Private Const WH_MOUSE_LL = 14

Private Const SB_LINEUP = 0
Private Const SB_LINELEFT = 0
Private Const SB_LINEDOWN = 1
Private Const SB_LINERIGHT = 1
Private Const SB_ENDSCROLL = 8
Private Const WS_VISIBLE = &H10000000
Private Const SBS_VERT = 1
Private Const SBS_HORZ = 0
Private Const WM_HSCROLL = &H114
Private Const SPI_GETWHEELSCROLLLINES = 104

Private Enum mButtons
    LBUTTON = &H1
    MBUTTON = &H10
    RBUTTON = &H2
End Enum

Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003

Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_ARENA_TRASHED = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259

Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_ALL_ACCESS = &H3F

Private Const REG_OPTION_NON_VOLATILE = 0

Private Declare Function RegCloseKey Lib "advapi32" _
        (ByVal hKey As Long) As Long
   
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias _
        "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
        ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions _
        As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes _
        As Long, phkResult As Long, lpdwDisposition As Long) As Long
   
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias _
        "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
        ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As _
        Long) As Long
   
Private Declare Function RegQueryValueExString Lib "advapi32" Alias _
        "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
        String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
        As String, lpcbData As Long) As Long
            
Private Declare Function RegQueryValueExLong Lib "advapi32" Alias _
        "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
        String, ByVal lpReserved As Long, lpType As Long, lpData As _
        Long, lpcbData As Long) As Long
   
Private Declare Function RegQueryValueExNULL Lib "advapi32" Alias _
        "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
        String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
        As Long, lpcbData As Long) As Long
   
Private Declare Function RegSetValueExString Lib "advapi32" Alias _
        "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As _
        String, ByVal cbData As Long) As Long
    
Private Declare Function RegSetValueExLong Lib "advapi32" Alias _
        "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
        ByVal cbData As Long) As Long

Dim nKeys As Long, Delta As Long, XPos As Long, YPos As Long
Dim OriginalWindowProc As Long
Dim pthWnd As Long
Dim lLineNumbers As Long
Dim MainWindowHwnd As Long  ' Main IDE window handle
Dim bHook As Boolean
Dim sLib As String
Dim hLib As Long

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Global gblFile As String
Global gVBInstance  As VBIDE.VBE
Global gwinWindow   As VBIDE.Window
Global gblMouseClick As Integer

Function GetFromIni(strSectionHeader As String, strVariableName As String, strFileName As String) As String
    Dim strReturn As String
    strReturn = String(255, Chr(0))
    GetFromIni = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), strFileName))
End Function

Function WriteToIni(strSectionHeader As String, strVariableName As String, strValue As String, strFileName As String) As Integer
    WriteToIni = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strFileName)
End Function

Public Function FileExists(strFile As String) As String
    On Error Resume Next 'Doesn't raise error - FileExists will be False
    FileExists = Dir(strFile, vbHidden) <> ""
End Function


Private Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, _
                           ByVal wParam As Long, ByVal lParam As Long) _
                           As Long
    On Error GoTo exitthis
    On Local Error GoTo exitthis
    
    Select Case uMsg
      Case WM_MOUSEWHEEL
        nKeys = wParam And 65535
        Delta = wParam / 65536 / WHEEL_DELTA

        XPos = LowWord(lParam)
        YPos = HighWord(lParam)
        
        pthWnd = WindowFromPointXY(XPos, YPos)
                
        ' Get the scroll bar for this window and send the vscroll to it
        Dim lRet As Long
        lRet = EnumChildWindows(pthWnd, AddressOf EnumChildProc, lParam)
        
    End Select

    If OriginalWindowProc <> 0 Then
        WindowProc = CallWindowProc(OriginalWindowProc, hwnd, uMsg, wParam, lParam)
    End If
    
exitthis:
    If Err Then Err.Clear
    On Error GoTo 0
    On Local Error GoTo 0
End Function

Public Sub UnHook()
    On Error GoTo exitthis
    On Local Error GoTo exitthis
    
    'Ensures that you don't try to unsubclass the window when
    'it is not subclassed.
    If OriginalWindowProc = 0 Then Exit Sub
    

    'Reset the window's function back to the original address.
    Dim hr As Long
    hr = SetWindowLong(MainWindowHwnd, GWL_WNDPROC, OriginalWindowProc)
    If hr <> 0 Then
        OriginalWindowProc = 0
        bHook = False
    Else
        Debug.Print "Unable to unhook:  SetWindowLong returns " & vbCrLf & hr & vbCrLf & Err.LastDllError
    End If
    
exitthis:
    If Err Then Err.Clear
    On Error GoTo 0
    On Local Error GoTo 0
End Sub

Public Sub Hook()
    On Error GoTo exitthis
    On Local Error GoTo exitthis
    
    ' GetLine Numbers
    SystemParametersInfo SPI_GETWHEELSCROLLLINES, 0, lLineNumbers, 0
    
    ' Adjust just in case, otherwise we'll never get the scroll notification.
    If lLineNumbers = 0 Then
        lLineNumbers = 1
    End If
    
    OriginalWindowProc = SetWindowLong(MainWindowHwnd, GWL_WNDPROC, AddressOf WindowProc)
    
    ' Set a flag indicating that we are hooking
    bHook = True
    
    ' Find out where we live on the filesystem
    Dim lRetVal As Long
    Dim sKeyName As String
    Dim sValue As String
    sKeyName = "CLSID\{B84F8C6E-BDDE-4384-9946-82EEE7F81D48}\InprocServer32"
    sValue = QueryValue(sKeyName, "")

    ' If we found where we live let's increase our ref count so we can do our own cleanup later
    If Len(sValue) > 0 Then
        sLib = Replace(sValue, ".exe", ".dll")
        hLib = LoadLibrary(sLib)
    End If

    Exit Sub
    
exitthis:
    If Err Then Err.Clear
    On Error GoTo 0
    On Local Error GoTo 0
End Sub

Function EnumChildProc(ByVal lhWnd As Long, ByVal lParam As Long) _
   As Long
   
    On Error GoTo exitthis
    On Local Error GoTo exitthis
    
   Dim retVal As Long
   Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
   Dim WinClass As String, WinTitle As String
   Dim WinRect As RECT
   Dim WinWidth As Long, WinHeight As Long

   retVal = GetClassName(pthWnd, WinClassBuf, 255)
   WinClass = StripNulls(WinClassBuf)  ' remove extra Nulls & spaces
   retVal = GetWindowText(lhWnd, WinTitleBuf, 255)
   WinTitle = StripNulls(WinTitleBuf)
   
   ' see the Windows Class and Title for each Child Window enumerated
   'Debug.Print "   hWnd = " & Hex(lhWnd) & " Child Class = "; WinClass; ", Title = "; WinTitle
   ' You can find any type of Window by searching for its WinClass
   Dim lRet As Long
   Dim i As Long
   
   ' Since we can have split windows we need to figure out which scroll bar to move.
   ' We can do this by comparing the Y position of the cursor against the vertical scrollbars
   ' that are children of the current window
   Dim wi As WINDOWINFO
   wi.cbSize = Len(wi)
   If GetWindowInfo(lhWnd, wi) And WinClass <> "MDIClient" Then
        If IsVerticalScrollBar(lhWnd) = True And wi.rcWindow.Top < YPos And wi.rcWindow.Bottom > YPos Then    ' TextBox Window
          
             If Delta > 0 Then                       ' Scroll Up
                  Do While i < Delta * lLineNumbers
                     lRet = PostMessage(pthWnd, WM_VSCROLL, SB_LINEUP, lhWnd)
                     i = i + 1
                  Loop
              Else                                   ' Scroll Down
                  Do While i > Delta * lLineNumbers
                     lRet = PostMessage(pthWnd, WM_VSCROLL, SB_LINEDOWN, lhWnd)
                     i = i - 1
                  Loop
              End If
        ElseIf IsHorizontalScrollBar(lhWnd) = True Then
             If Delta > 0 Then                       ' Scroll Left
                 Do While i < Delta * lLineNumbers
                     lRet = PostMessage(pthWnd, WM_HSCROLL, SB_LINELEFT, lhWnd)
                     i = i + 1
                 Loop
              Else                                   ' Scroll Right
                 Do While i > Delta * lLineNumbers
                     lRet = PostMessage(pthWnd, WM_HSCROLL, SB_LINERIGHT, lhWnd)
                     i = i - 1
                 Loop
              End If
        End If
   End If
   
   EnumChildProc = bHook                              ' Continue enumerating the windows based on whether we are hooking or not
   
   ' It's possible that the addin has already been requested to unload and in such a case we will call free library on ourselves
   ' to reduce our ref count since we incremented it on our own so we can do a clean shutdown
   If Not bHook Then
        If Not FreeLibrary(hLib) Then
             Debug.Print "Unable to FreeLibrary: " & Err.Number & vbCrLf & Err.LastDllError
        End If
   End If
   
   Exit Function
exitthis:
    If Err Then Err.Clear
    On Error GoTo 0
    On Local Error GoTo 0
   
End Function

Function EnumThreadProc(ByVal lhWnd As Long, ByVal lParam As Long) _
   As Long

    On Error GoTo exitthis
    On Local Error GoTo exitthis
    
   Dim retVal As Long
   Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
   Dim WinClass As String, WinTitle As String

    
   retVal = GetClassName(lhWnd, WinClassBuf, 255)
   WinClass = StripNulls(WinClassBuf)  ' remove extra Nulls & spaces
   retVal = GetWindowText(lhWnd, WinTitleBuf, 255)
   WinTitle = StripNulls(WinTitleBuf)

   ' see the Windows Class and Title for top level Window
   Debug.Print "Thread Window Class = "; WinClass; ", Title = "; _
   WinTitle
   EnumThreadProc = True
   
   If ((InStr(1, WinTitle, "Microsoft Visual Basic") <> 0 Xor WinTitle = "") _
    And (WinClass = "wndclass_desked_gsk" Xor WinClass = "ObtbarWndClass") _
    And MainWindowHwnd = 0) Then
    
    MainWindowHwnd = lhWnd
    ' Setup the windows Hook
    Hook
   
   End If
   
   Exit Function
exitthis:
    If Err Then Err.Clear
    On Error GoTo 0
    On Local Error GoTo 0
   
End Function

Public Function StripNulls(OriginalStr As String) As String
   ' This removes the extra Nulls so String comparisons will work
   If (InStr(OriginalStr, Chr(0)) > 0) Then
      OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
   End If
   StripNulls = OriginalStr
End Function

Public Function IsVerticalScrollBar(hwnd As Long) As Boolean

    ' Check the style of the window specified by hWnd to see if it's a vertical scrollbar

    Dim wi As WINDOWINFO
    wi.cbSize = Len(wi)
    
    If GetWindowInfo(hwnd, wi) Then
        If (wi.dwStyle And WS_VISIBLE) > 0 And (wi.dwStyle And SBS_VERT) > 0 Then
            IsVerticalScrollBar = True
            Exit Function
        End If
    End If
  
    IsVerticalScrollBar = False

End Function

Public Function IsHorizontalScrollBar(hwnd As Long) As Boolean

    ' Check the style of the window specified by hWnd to see if it's a horizontal scrollbar

    Dim wi As WINDOWINFO
    wi.cbSize = Len(wi)
    
    If GetWindowInfo(hwnd, wi) Then
        If (wi.dwStyle And WS_VISIBLE) > 0 And (wi.dwStyle And SBS_HORZ) > 0 Then
            IsHorizontalScrollBar = True
            Exit Function
        End If
    End If
  
    IsHorizontalScrollBar = False

End Function

Public Function QueryValueCU(sKeyName As String, sValueName As String) As Variant
    Dim lRetVal As Long         'result of the API functions
    Dim hKey As Long         'handle of opened key
    Dim vValue As Variant      'setting of queried value

    lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sKeyName, 0, KEY_QUERY_VALUE, hKey)
    lRetVal = QueryValueEx(hKey, sValueName, vValue)
    RegCloseKey (hKey)
    
    QueryValueCU = vValue
End Function

Private Function QueryValue(sKeyName As String, sValueName As String) As Variant
    Dim lRetVal As Long         'result of the API functions
    Dim hKey As Long         'handle of opened key
    Dim vValue As Variant      'setting of queried value

    lRetVal = RegOpenKeyEx(HKEY_CLASSES_ROOT, sKeyName, 0, KEY_QUERY_VALUE, hKey)
    lRetVal = QueryValueEx(hKey, sValueName, vValue)
    RegCloseKey (hKey)
    
    QueryValue = vValue
End Function

Public Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
       Dim lValue As Long
       Dim sValue As String
       Select Case lType
           Case REG_SZ
               sValue = vValue & Chr$(0)
               SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
           Case REG_DWORD
               lValue = vValue
               SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
           End Select
End Function

Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
       Dim cch As Long
       Dim lrc As Long
       Dim lType As Long
       Dim lValue As Long
       Dim sValue As String

    On Error GoTo QueryValueExError
    On Local Error GoTo QueryValueExError
    
    
       ' Determine the size and type of data to be read
       lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
       If lrc <> ERROR_NONE Then Exit Function

       Select Case lType
           ' For strings
           Case REG_SZ:
               sValue = String(cch, 0)

   lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, _
   sValue, cch)
               If lrc = ERROR_NONE Then
                   vValue = Left$(sValue, cch - 1)
               Else
                   vValue = Empty
               End If
           ' For DWORDS
           Case REG_DWORD:
   lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, _
   lValue, cch)
               If lrc = ERROR_NONE Then vValue = lValue
           Case Else
               'all other data types not supported
               lrc = -1
       End Select

QueryValueExExit:
       QueryValueEx = lrc
       Exit Function

QueryValueExError:
       Resume QueryValueExExit
End Function

Private Function LowWord(ByVal inDWord As Long) As Integer
    LowWord = inDWord And &H7FFF&
    If (inDWord And &H8000&) Then LowWord = LowWord Or &H8000
End Function

Private Function HighWord(ByVal inDWord As Long) As Integer
    HighWord = LowWord(((inDWord And &HFFFF0000) \ &H10000) And &HFFFF&)
End Function





