Attribute VB_Name = "modParentProc"

#Const modParentProc = -1
Option Explicit
'TOP DOWN
Option Compare Binary


Option Private Module
Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_WINDOWPOSCHANGED = &H47

Type WINDOWPOS
        hwnd As Long
        hWndInsertAfter As Long
        X As Long
        Y As Long
        cx As Long
        cy As Long
        flags As Long
End Type

Private Const WM_ACTIVATE = &H6
Private Const WM_SIZE = &H5
Private Const WM_CLOSE = &H10
Private Const WM_DESTROY = &H2
Private Const WM_QUIT = &H12
Private Const WM_QUERYENDSESSION = &H11
Private Const WM_ENDSESSION = &H16

Private Const SIZE_RESTORED = 0
Private Const SIZE_MINIMIZED = 1
Private Const SIZE_MAXIMIZED = 2
Private Const SIZE_MAXSHOW = 3
Private Const SIZE_MAXHIDE = 4

Private Const WM_MOVE = &H3
Private Const WM_SIZING = &H214
Private Const WM_MOVING = &H216
Private Const WM_ENTERSIZEMOVE = &H231
Private Const WM_EXITSIZEMOVE = &H232

Private Const GWL_WNDPROC = (-4)

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_HIDE = 0

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lngParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function GetActiveWindow Lib "user32" () As Long

Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const ENDSESSION_LOGOFF As Long = &H80000000

Private UC As New VBA.Collection

Public Function Hook(ByRef Obj As Window) As String
    If Obj.PrevWndProc = 0 Then
        
          Dim NewObj As Window
          Dim UCKey As String
        
          UCKey = "hw" & Obj.ParentHWnd
          
          Set NewObj = Obj
          UC.Add NewObj, UCKey
          
          Set NewObj = Nothing

        Hook = UCKey
        
        Obj.PrevWndProc = SetWindowLong(Obj.ParentHWnd, GWL_WNDPROC, AddressOf WindowProc)
    End If
End Function

Public Sub Unhook(ByRef Obj As Window)
    If Obj.PrevWndProc <> 0 Then
        SetWindowLong Obj.ParentHWnd, GWL_WNDPROC, Obj.PrevWndProc
        Obj.PrevWndProc = 0
        UC.Remove "hw" & Obj.ParentHWnd
    End If
End Sub

Public Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lngParam As Long) As Long
   
    On Error GoTo walkoff
    
        Dim TempUC As Window
        Dim WinPos As WINDOWPOS
        
        Set TempUC = UC.Item("hw" & hw)
        
        Select Case uMsg
            Case WM_QUERYENDSESSION
                TempUC.Visible = False
               ' WindowProc = DefWindowProc(hw, uMsg, wParam, lngParam)
               Unhook TempUC
                DestroyWindow hw
'                Dim isVisible As Boolean
'                isVisible = TempUC.Visible
'                If TempUC.Visible Then TempUC.Visible = False
'                Dim proc As Long
'                proc = TempUC.PrevWndProc
'                Unhook TempUC
'                WindowProc = CallWindowProc(proc, hw, uMsg, wParam, lngParam)
'                If WindowProc Then
'                    DestroyWindow hw
'                Else
'                    Hook TempUC
'                    If isVisible Then TempUC.Visible = True
'                End If
               WindowProc = 1
            Case WM_CLOSE, WM_QUIT, WM_DESTROY
                Unhook TempUC
                DestroyWindow hw
                WindowProc = 1
            Case WM_SIZE
                Select Case wParam
                    Case SIZE_RESTORED
                        TempUC.ParentWindowState = vbNormal
                    Case SIZE_MINIMIZED
                        TempUC.ParentWindowState = vbMinimized
                    Case SIZE_MAXIMIZED
                        TempUC.ParentWindowState = vbMaximized
                    Case SIZE_MAXSHOW
                    Case SIZE_MAXHIDE
                End Select
                
            Case WM_ACTIVATE
                TempUC.ParentIsActive = True
            Case WM_WINDOWPOSCHANGED
                CopyMemory WinPos, ByVal lngParam, LenB(WinPos)
            Case Else
                WindowProc = CallWindowProc(TempUC.PrevWndProc, hw, uMsg, wParam, lngParam)
        End Select
        
        If Not ((uMsg And WM_QUERYENDSESSION) = WM_QUERYENDSESSION) Then
    
            Select Case GetActiveWindow
                Case TempUC.hwnd, TempUC.ParentHWnd
                    If GetForegroundWindow = GetActiveWindow Then TempUC.ParentIsActive = True
                Case 0
                    If TempUC.WindowState = vbNormal Then
                        TempUC.ParentIsActive = False
                    End If
            End Select
    
            Select Case WinPos.flags
                Case 33072
                    TempUC.ParentWindowState = vbMinimized
                Case 33060
                    TempUC.ParentWindowState = vbNormal
                Case 32804
                    TempUC.ParentWindowState = vbMaximized
                Case 6147
                    TempUC.ParentIsActive = True
            End Select
    
            Set TempUC = Nothing
        End If


    If uMsg = WM_ENTERSIZEMOVE Then
        WindowProc = DefWindowProc(hw, uMsg, wParam, lngParam)
    End If
    
    Exit Function
walkoff:
    Err.Clear
    ShowWindow hw, SW_HIDE
    DestroyWindow hw
    WindowProc = 1
End Function
