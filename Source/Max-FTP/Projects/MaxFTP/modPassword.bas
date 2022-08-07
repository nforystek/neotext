Attribute VB_Name = "modPassword"
Option Explicit
'TOP DOWN

Option Compare Binary
Option Private Module
Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_WINDOWPOSCHANGED = &H47

Type WINDOWPOS
        hWnd As Long
        hWndInsertAfter As Long
        X As Long
        Y As Long
        cx As Long
        cy As Long
        flags As Long
End Type

Public Const WM_ACTIVATE = &H6
Private Const WM_SIZE = &H5

Private Const SIZE_RESTORED = 0
Private Const SIZE_MINIMIZED = 1
Private Const SIZE_MAXIMIZED = 2
Private Const SIZE_MAXSHOW = 3
Private Const SIZE_MAXHIDE = 4

Private Const GWL_WNDPROC = (-4)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lngParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function GetActiveWindow Lib "user32" () As Long

Private UC As New Collection

Public Type PasswordInfo
    HostURL As String
    Port As Long
    Pasv As Boolean
    PortRange As String
    Adapter As Long
    
    Username As String
    Password As String
End Type

Public Function ShowPassword(frm As Form, ByRef pwdInfo As PasswordInfo) As Boolean
    
    frm.sInfo.Reset
        
    frm.sInfo.sHostURL.text = pwdInfo.HostURL
    frm.sInfo.sPort.text = pwdInfo.Port
    frm.sInfo.sPassive.Value = -CInt(pwdInfo.Pasv)
    frm.sInfo.sPortRange.text = pwdInfo.PortRange
    frm.sInfo.sAdapter.ListIndex = (pwdInfo.Adapter - 1)
    
    LoadCache frm.sInfo
    
reshow:

    frm.Visible = True
    TopMostForm frm, True
    frm.VisVar = True
    
    If frm.sInfo.sHostURL.text = "" Or frm.sInfo.sHostURL.text = "ftp://" Then
        frm.sInfo.sHostURL.SetFocus
    Else
        frm.sInfo.sUserName.SetFocus
    End If
    
    Do While (Not (frm Is Nothing))
        If (Not (frm Is Nothing)) Then
            If (Not frm.VisVar) Then Exit Do
        End If
        DoEvents
        Sleep 1
    Loop

    If Not (frm Is Nothing) Then
                
        pwdInfo.HostURL = frm.sInfo.sHostURL.text
        If (frm.sInfo.sUserName.text = "") And frm.IsOk Then
            frm.sInfo.sUserName.text = "anonymous"
            frm.sInfo.sSavePass.Value = 1
        End If
        pwdInfo.Username = frm.sInfo.sUserName.text
        If (Trim(LCase(frm.sInfo.sUserName.text)) = "anonymous") And frm.IsOk Then
            If Not (InStr(frm.sInfo.sPassword.text, ".") > InStr(frm.sInfo.sPassword.text, "@")) Then
                MsgBox "Anonymous logins must provide an email address as the password consisting of the following mask *@*.*", vbInformation + vbOK, AppName
                GoTo reshow
            End If
        End If
        pwdInfo.Password = frm.sInfo.sPassword.text
        If IsNumeric(frm.sInfo.sPort.text) Then pwdInfo.Port = CLng(frm.sInfo.sPort.text)
        pwdInfo.Pasv = frm.sInfo.sPassive.Value
        pwdInfo.PortRange = frm.sInfo.sPortRange.text
        pwdInfo.Adapter = (frm.sInfo.sAdapter.ListIndex + 1)
            
        If (Not Left(Trim(LCase(pwdInfo.HostURL)), 6) = "ftp://") And (Not pwdInfo.HostURL = "") Then
            
            pwdInfo.HostURL = "ftp://" & pwdInfo.HostURL
            
        End If

        ShowPassword = frm.IsOk
        
        Unload frm
    Else
        ShowPassword = False
    End If
    
End Function


Public Sub UnSetControlHost(ByVal obj As frmPassword)
  On Error Resume Next
  UC.Remove "hw" & obj.ParentHWnd
  If Err Then Err.Clear
  On Error GoTo 0
End Sub

Private Function SetControlHost(ByRef obj As frmPassword) As String
  
  Dim NewObj As frmPassword
  Dim UCKey As String

  UCKey = "hw" & obj.ParentHWnd
  
  Set NewObj = obj
  UC.Add NewObj, UCKey
  
  Set NewObj = Nothing
  
  SetControlHost = UCKey
      
End Function
Public Function ControlHostExists(ByVal hWnd As String)
    Dim test
    On Error Resume Next
    Set test = UC("hw" & hWnd)
    ControlHostExists = (Not (test Is Nothing)) And (Not Err)
    If Err Then Err.Clear
    On Error GoTo 0
End Function

Public Function Hook(ByRef obj As frmPassword) As String
    If obj.ParentHWnd > 0 Then
        If (Not ControlHostExists(obj.ParentHWnd)) Then
            Hook = SetControlHost(obj)
            obj.PrevWndProc = SetWindowLong(obj.ParentHWnd, GWL_WNDPROC, AddressOf WindowProc)
        End If
    End If
End Function

Public Sub Unhook(ByVal obj As frmPassword)
    If obj.ParentHWnd > 0 Then
        If ControlHostExists(obj.ParentHWnd) Then
            SetWindowLong obj.ParentHWnd, GWL_WNDPROC, obj.PrevWndProc
            UnSetControlHost obj
        End If
    End If
End Sub

Private Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lngParam As Long) As Long
    
    If ControlHostExists(hw) Then
        Dim TempUC As frmPassword
        Dim WinPos As WINDOWPOS
        
        Set TempUC = UC.Item("hw" & hw)
        
        Select Case uMsg
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
        End Select
        
        Select Case GetActiveWindow
            Case TempUC.hWnd, TempUC.ParentHWnd
                TempUC.ParentIsActive = True
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
        
        WindowProc = CallWindowProc(TempUC.PrevWndProc, hw, uMsg, wParam, lngParam)
        
    End If
    
End Function
