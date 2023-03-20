Attribute VB_Name = "Module1"
#Const Module1 = -1

' Root value for hidden window caption
Public Const PROC_CAPTION = "ApartmentDemoProcessWindow"

Public Const ERR_InternalStartup = &H600
Public Const ERR_NoAutomation = &H601

Public Const ENUM_STOP = 0
Public Const ENUM_CONTINUE = 1

Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
   (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindowThreadProcessId Lib "user32" _
   (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function EnumThreadWindows Lib "user32" _
   (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lParam As Long) _
   As Long

' Window handle retrieved by EnumThreadWindows.
Private mhwndVB As Long
' Hidden form used to identify main thread.
Private mfrmProcess As New frmProcess
' Process ID.
Private mlngProcessID As Long

Sub Main()
   Dim ma As MainApp

   ' Borrow a window handle to use to obtain the process
   '   ID (see EnumThreadWndMain call-back, below).
   Call EnumThreadWindows(App.ThreadID, AddressOf EnumThreadWndMain, 0&)
   If mhwndVB = 0 Then
      Err.Raise ERR_InternalStartup + vbObjectError, , _
      "Internal error starting thread"
   Else
      Call GetWindowThreadProcessId(mhwndVB, mlngProcessID)
      ' The process ID makes the hidden window caption unique.
      If 0 = FindWindow(vbNullString, PROC_CAPTION & CStr(mlngProcessID)) Then
            ' The window wasn't found, so this is the first thread.
            If App.StartMode = vbSModeStandalone Then
               ' Create hidden form with unique caption.
               mfrmProcess.Caption = PROC_CAPTION & CStr(mlngProcessID)
               ' The Initialize event of MainApp (Instancing =
               '   PublicNotCreatable) shows the main user interface.
               Set ma = New MainApp
               ' (Application shutdown is simpler if there is no
               '   global reference to MainApp; instead, MainApp
               '   should pass Me to the main user form, so that
               '   the form keeps MainApp from terminating.)
            Else
               Err.Raise ERR_NoAutomation + vbObjectError, , _
                  "Application can't be started with Automation"
            End If
         End If
   End If
End Sub

' Call-back function used by EnumThreadWindows.
Public Function EnumThreadWndMain(ByVal hwnd As Long, ByVal _
   lParam As Long) As Long
   ' Save the window handle.
   mhwndVB = hwnd
   ' The first window is the only one required.
   ' Stop the iteration as soon as a window has been found.
   EnumThreadWndMain = ENUM_STOP
End Function

' MainApp calls this Sub in its Terminate event;
'   otherwise the hidden form will keep the
'   application from closing.
Public Sub FreeProcessWindow()
   Unload mfrmProcess
   Set mfrmProcess = Nothing
End Sub



