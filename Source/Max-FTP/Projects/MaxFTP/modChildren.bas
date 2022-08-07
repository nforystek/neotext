Attribute VB_Name = "modChildren"
#Const [True] = -1
#Const [False] = 0



#Const modChildren = -1
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
        Flags As Long
End Type

Private Const WM_ACTIVATE = &H6
Private Const WM_SIZE = &H5
Private Const WM_CLOSE = &H10
Private Const WM_DESTROY = &H2
Private Const WM_QUIT = &H12
Private Const WM_QUERYENDSESSION = &H11
Private Const WM_ENDSESSION = &H16

Private Const WM_CANCELMODE = &H1F
Private Const WM_COMMAND = &H111
Private Const WM_PARENTNOTIFY = &H210
Private Const WM_SYSCOMMAND = &H112


Private Const SIZE_RESTORED = 0
Private Const SIZE_MINIMIZED = 1
Private Const SIZE_MAXIMIZED = 2
Private Const SIZE_MAXSHOW = 3
Private Const SIZE_MAXHIDE = 4

Private Const GWL_WNDPROC = (-4)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lngParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Const ENDSESSION_LOGOFF As Long = &H80000000

Private UC As New NTNodes10.Collection

Public Type PasswordInfo
    HostURL As String
    Port As Long
    Pasv As Boolean
    PortRange As String
    Adapter As Long
    
    Username As String
    Password As String
    
    SSL As Boolean
End Type

Public Function ShowOverwrite(frm As Form, ByRef OptionClick As OverwriteTypes, ByVal ThisFile As String, ByVal WithFile As String) As Boolean
    
    frm.OptionClick = or_Prompt
    frm.DestFile.Text = ThisFile
    frm.SourceFile.Text = WithFile
   
    Dim tFile As String
    Dim wFile As String
    Dim tSize As Boolean

    tFile = Trim(NextArg(RemoveArg(ThisFile, "FileSize: "), "FileDate:"))
    wFile = Trim(NextArg(RemoveArg(WithFile, "FileSize: "), "FileDate:"))
    
    If IsNumeric(tFile) And IsNumeric(wFile) Then
        tSize = (CDbl(wFile) > CDbl(tFile))
    Else
        tSize = True
    End If
    
    Dim fa As clsFileAssoc
    Set fa = New clsFileAssoc
    frm.Command1(4).enabled = (fa.GetTransferType(GetFileExt(ThisFile)) = TransferModes.Binary) And tSize
    Set fa = Nothing
    
    frm.Visible = True
    TopMostForm frm, True
    frm.VisVar = True

    Do While (Not (frm Is Nothing))
        If (Not (frm Is Nothing)) Then
            If (Not frm.VisVar) Or Not frm.Visible Then Exit Do
        End If

        DoEvents
        
    Loop

    If Not (frm Is Nothing) Then

        OptionClick = frm.OptionClick
        
        Unload frm
    Else
        ShowOverwrite = False
    End If

End Function

Public Function ShowPassword(frm As Form, ByRef pwdInfo As PasswordInfo) As Boolean

    frm.sInfo.Reset
        
    frm.sInfo.sHostURL.Text = pwdInfo.HostURL
    frm.sInfo.sPort.Text = pwdInfo.Port
    frm.sInfo.sPassive.Value = -CInt(pwdInfo.Pasv)
    frm.sInfo.sPortRange.Text = pwdInfo.PortRange
    frm.sInfo.sAdapter.ListIndex = pwdInfo.Adapter
    frm.sInfo.sSSL.Value = -CInt(pwdInfo.SSL)
    LoadCache frm.sInfo
    
reshow:

    frm.Visible = True
    TopMostForm frm, True
    frm.VisVar = True
    
    If frm.sInfo.sHostURL.Text = "" Or frm.sInfo.sHostURL.Text = "ftp://" Or frm.sInfo.sHostURL.Text = "ftps://" Then
        frm.sInfo.sHostURL.SetFocus
    Else
        frm.sInfo.sUserName.SetFocus
    End If
    
    Do While (Not (frm Is Nothing))
        If (Not (frm Is Nothing)) Then
            If (Not frm.VisVar) Or Not frm.Visible Then Exit Do
        End If

        DoEvents
        Sleep 1
        
    Loop

    If Not (frm Is Nothing) Then
                
        pwdInfo.HostURL = frm.sInfo.sHostURL.Text
        If (frm.sInfo.sUserName.Text = "") And frm.IsOk Then
            frm.sInfo.sUserName.Text = "anonymous"
            frm.sInfo.sSavePass.Value = 1
        End If
        pwdInfo.Username = frm.sInfo.sUserName.Text
        If (Trim(LCase(frm.sInfo.sUserName.Text)) = "anonymous") And frm.IsOk Then
            If Not (InStr(frm.sInfo.sPassword.Text, ".") > InStr(frm.sInfo.sPassword.Text, "@")) Then
                MsgBox "Anonymous logins must provide an email address as the password consisting of the following mask *@*.*", vbInformation + vbOK, AppName
                GoTo reshow
            End If
        End If
        pwdInfo.Password = frm.sInfo.sPassword.Text
        If IsNumeric(frm.sInfo.sPort.Text) Then pwdInfo.Port = CLng(frm.sInfo.sPort.Text)
        pwdInfo.Pasv = frm.sInfo.sPassive.Value
        pwdInfo.PortRange = frm.sInfo.sPortRange.Text
        pwdInfo.Adapter = frm.sInfo.sAdapter.ListIndex
        pwdInfo.SSL = (frm.sInfo.sSSL.Value = 1)
        If (Not Left(Trim(LCase(pwdInfo.HostURL)), 6) = "ftp://") And (Not Left(Trim(LCase(pwdInfo.HostURL)), 7) = "ftps://") And (Not pwdInfo.HostURL = "") Then
            
            pwdInfo.HostURL = IIf(dbSettings.GetProfileSetting("SSL") = 1, "ftps://", "ftp://") & pwdInfo.HostURL
            
        End If

        ShowPassword = frm.IsOk
        
        Unload frm
    Else
        ShowPassword = False
    End If

End Function


Public Sub UnSetControlHost(ByVal obj As Object)
  On Error Resume Next
  UC.Remove "hw" & obj.ParentHWnd
  If Err Then Err.Clear
  On Error GoTo 0
End Sub

Private Function SetControlHost(ByRef obj As Object) As String
  
  Dim NewObj As Object
  Dim UCKey As String

  UCKey = "hw" & obj.ParentHWnd
    
  Set NewObj = obj
  UC.Add NewObj, UCKey
  
  Set NewObj = Nothing
  
  SetControlHost = UCKey
      
End Function

Public Function Hook(ByRef obj As Object) As String
    If obj.ParentHWnd > 0 Then
        If (Not UC.Exists("hw" & obj.ParentHWnd)) Then
            Hook = SetControlHost(obj)
            obj.PrevWndProc = SetWindowLong(obj.ParentHWnd, GWL_WNDPROC, AddressOf WindowProc)
        End If
    End If
End Function

Public Sub Unhook(ByVal obj As Object)
    If obj.ParentHWnd > 0 Then
        If UC.Exists("hw" & obj.ParentHWnd) Then
            SetWindowLong obj.ParentHWnd, GWL_WNDPROC, obj.PrevWndProc
            UnSetControlHost obj
        End If
    End If
End Sub

Private Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lngParam As Long) As Long
    On Error GoTo walkoff

    If UC.Exists("hw" & hw) Then
        Dim TempUC As Object
        Dim WinPos As WINDOWPOS
        
        Set TempUC = UC.Item("hw" & hw)
        
        Select Case uMsg
            Case WM_QUERYENDSESSION ', WM_ENDSESSION
                'WindowProc = DefWindowProc(hw, uMsg, wParam, lngParam)
                Unhook TempUC
                DestroyWindow hw
'                Dim proc As Long
'                proc = TempUC.PrevWndProc
'                Unhook TempUC
'                WindowProc = CallWindowProc(proc, hw, uMsg, wParam, lngParam)
'                If WindowProc Then
'                    DestroyWindow hw
'                Else
'                    Hook TempUC
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
            Case WM_ACTIVATE, WM_ACTIVATEAPP
                TempUC.ParentIsActive = True
            Case WM_WINDOWPOSCHANGED
                CopyMemory WinPos, ByVal lngParam, LenB(WinPos)
            Case Else

                WindowProc = CallWindowProc(TempUC.PrevWndProc, hw, uMsg, wParam, lngParam)
               
        End Select
        
        Select Case GetActiveWindow
            Case TempUC.hwnd, TempUC.ParentHWnd
                TempUC.ParentIsActive = True
            Case 0
                If TempUC.WindowState = vbNormal Then
                    TempUC.ParentIsActive = False
                End If
        End Select

        Select Case WinPos.Flags
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
'WindowProc = 1
    Exit Function
walkoff:
'Debug.Print Err.Description

    Err.Clear
    WindowProc = 1
End Function
