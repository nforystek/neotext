VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Controller"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Option Compare Binary

Public Enum SvcCallBacks
    SvcControlCallBack = 1
    SvcStartCallBack = 2
    SvcPauseCallBack = 3
    SvcContinueCallBack = 4
    SvcStopCallBack = 5
    SvcLoggedOffCallBack = 6
End Enum
'
'Public Enum SvcControls
'    SvcControlPauseContinue = 2
'    SvcControlShutDown = 4
'    SvcControlStartStop = 1
'    SvcControlStop = 1
'End Enum

'Public Enum SvcEventID
'    svcMessageDebug = 108
'    svcMessageError = 109
'    svcMessageInfo = 110
'End Enum
'
'Public Enum SvcEventType
'    svcEventAuditFailure = 16
'    svcEventAuditSuccess = 8
'    svcEventError = 1
'    svcEventInformation = 4
'    svcEventSuccess = 0
'    svcEventWarning = 2
'End Enum

Public Enum SvcStartMode
    svcStartAutomatic = 2
    svcStartDisabled = 4
    svcStartManual = 3
End Enum

'Public Enum SvcState
'    svcContinuePending = 5
'    svcPaused = 7
'    svcPausePending = 6
'    svcRunning = 4
'    svcStartPending = 2
'    svcStopped = 1
'    svcStopPending = 3
'End Enum

Private pServiceName As String
Private pDisplayName As String
Private pDescription As String

Private pLoadOrderGroup As String
Private pInteractive As Boolean
Private pAutoStart As Boolean
Private pAccount As String
Private pPassword As String
Private pProcIDCache As Long
Private pWasStarted As Boolean
Private pWasInstall As Boolean
Private pIsLegacyOS As Boolean
Private pSvcRunning As Boolean

Private WithEvents WinServe As frmService
Attribute WinServe.VB_VarHelpID = -1
Private Win98Svc As LegacyOS

Private eventCallBacks() As Long
Private classCallBacks As CallBacks

Public Event ControlService(lEvent As Long)
Public Event StartService(Success As Boolean)
Public Event PauseService(Success As Boolean)
Public Event ContinueService(Success As Boolean)
Public Event StopService()
Public Event LoggedOffService()

Public Sub AddCallBack(ByVal ServiceCallBack As SvcCallBacks, ByVal lpProcAddress As Long)
    If classCallBacks Is Nothing Then
        Set classCallBacks = New CallBacks
        ReDim eventCallBacks(1 To 6) As Long
        memCallBacks = ObjectPointers(classCallBacks, 3, 0, 0, 2)
    End If
    eventCallBacks(ServiceCallBack) = lpProcAddress
End Sub
Private Function RequiredVarsSet() As Boolean
    RequiredVarsSet = Not (pServiceName = "" Or pDisplayName = "" Or pDescription = "")
End Function

Public Property Get IsLegacyOS() As Boolean
    IsLegacyOS = pIsLegacyOS
End Property

Public Property Get Running() As Boolean
    If (Not pIsLegacyOS) Then
        Running = WinServe.NTService1.Running
    Else
        Running = (pWasStarted And pSvcRunning) And (Not pWasInstall)
    End If
End Property

Friend Property Let Running(ByVal newVal As Boolean)
    pSvcRunning = newVal
End Property

Public Property Get ServiceName() As String
    If (Not pIsLegacyOS) And ((Not pWasInstall) Or pWasStarted) Then
        ServiceName = WinServe.NTService1.ServiceName
    Else
        ServiceName = pServiceName
    End If
End Property

Public Property Let ServiceName(ByVal newVal As String)
    If (pWasInstall Or pWasStarted) Then
        Err.Raise vbObjectError, App.EXEName, "Unable to change the ServiceName property; Install or StartService has already been invoked."
    Else
        pServiceName = newVal
        If (Not pIsLegacyOS) Then WinServe.NTService1.ServiceName = newVal
    End If
End Property

Public Property Get DisplayName() As String
    If (Not pIsLegacyOS) And ((Not pWasInstall) Or pWasStarted) Then
        DisplayName = WinServe.NTService1.DisplayName
    Else
        DisplayName = pDisplayName
    End If
End Property

Public Property Let DisplayName(ByVal newVal As String)
    If (pWasInstall Or pWasStarted) Then
        Err.Raise vbObjectError, App.EXEName, "Unable to change the DisplayName property; Install or StartService has already been invoked."
    Else
        pDisplayName = newVal
        If (Not pIsLegacyOS) Then WinServe.NTService1.DisplayName = newVal
    End If
End Property

Public Property Get Description() As String
    If (Not pIsLegacyOS) And ((Not pWasInstall) Or pWasStarted) Then
        Description = WinServe.NTService1.GetSetting("", "Description", pDescription)
    Else
        Description = pDescription
    End If
End Property

Public Property Let Description(ByVal newVal As String)
    If (pWasInstall Or pWasStarted) Then
        Err.Raise vbObjectError, App.EXEName, "Unable to change the Description property; Install or StartService has already been invoked."
    Else
        pDescription = newVal
    End If
End Property

Public Property Get Interactive() As Boolean
    If (Not pIsLegacyOS) And ((Not pWasInstall) Or pWasStarted) Then
        Interactive = WinServe.NTService1.Interactive
    Else
        Interactive = pInteractive
    End If
End Property

Public Property Let Interactive(ByVal newVal As Boolean)
    If (pWasInstall Or pWasStarted) Then
        Err.Raise vbObjectError, App.EXEName, "Unable to change the Interactive property; Install or StartService has already been invoked."
    Else
        pInteractive = newVal
        If (Not pIsLegacyOS) Then WinServe.NTService1.Interactive = newVal
    End If
End Property

Public Property Get LoadOrderGroup() As String
    If (Not pIsLegacyOS) And ((Not pWasInstall) Or pWasStarted) Then
        LoadOrderGroup = WinServe.NTService1.LoadOrderGroup
    Else
        LoadOrderGroup = pLoadOrderGroup
    End If
End Property
Public Property Let LoadOrderGroup(ByVal newVal As String)
    If (pWasInstall Or pWasStarted) Then
        Err.Raise vbObjectError, App.EXEName, "Unable to change the LoadOrderGroup property; Install or StartService has already been invoked."
    Else
        pLoadOrderGroup = newVal
        If (Not pIsLegacyOS) Then WinServe.NTService1.LoadOrderGroup = newVal
    End If
End Property

Public Property Get AutoStart() As Variant
    If (Not pIsLegacyOS) And ((Not pWasInstall) Or pWasStarted) Then
        AutoStart = (WinServe.NTService1.StartMode = svcStartAutomatic)
    Else
        AutoStart = (Not (GetSettingLong(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName, "RegistryRun", 0) = 0)) _
                    And (GetSettingLong(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName, "DisableExec", 0) = 0)
    End If
End Property

Public Property Let AutoStart(ByVal newVal As Variant)
    If (pWasInstall Or pWasStarted) Then
        Err.Raise vbObjectError, App.EXEName, "Unable to change the AutoStart property; Install or StartService has already been invoked."
    Else
        pAutoStart = newVal
        If (Not pIsLegacyOS) Then
            If TypeName(newVal) = "Boolean" Then
                If pAutoStart Then
                    WinServe.NTService1.StartMode = svcStartAutomatic
                Else
                    WinServe.NTService1.StartMode = svcStartManual
                End If
            ElseIf IsNumeric(newVal) Then
                If CLng(newVal) >= 2 And CLng(newVal) <= 4 Then
                    WinServe.NTService1.StartMode = newVal
                Else
                    Err.Raise 8, App.EXEName, "Invalid value for AutoStart, the value must be of SvcStartMode or a Boolean."
                End If
            Else
                Err.Raise 8, App.EXEName, "Invalid value for AutoStart, the value must be of SvcStartMode or a Boolean."
            End If
        Else
            If TypeName(newVal) = "Boolean" Then
                Win98Svc.AutoRun ServiceName, newVal
            ElseIf IsNumeric(newVal) Then
                If CLng(newVal) >= 2 And CLng(newVal) <= 4 Then
                    Win98Svc.Disable ServiceName, (newVal = SvcStartMode.svcStartDisabled)
                    Win98Svc.AutoRun ServiceName, (newVal = SvcStartMode.svcStartAutomatic)
                Else
                    Err.Raise 8, App.EXEName, "Invalid value for AutoStart, the value must be of SvcStartMode or a Boolean."
                End If
            Else
                Err.Raise 8, App.EXEName, "Invalid value for AutoStart, the value must be of SvcStartMode or a Boolean."
            End If
        End If
    End If
End Property

Public Property Get Account() As String
    If (Not pIsLegacyOS) And ((Not pWasInstall) Or pWasStarted) Then
        Account = WinServe.NTService1.Account
    Else
        Account = pAccount
    End If
End Property

Public Property Let Account(ByVal newVal As String)
    If (pWasInstall Or pWasStarted) Then
        Err.Raise vbObjectError, App.EXEName, "Unable to change the Account property; Install or StartService has already been invoked."
    Else
        pAccount = newVal
        If (Not pIsLegacyOS) Then WinServe.NTService1.Account = newVal
    End If
End Property

Public Property Get Password() As String
    If (Not pIsLegacyOS) And ((Not pWasInstall) Or pWasStarted) Then
        Password = WinServe.NTService1.Password
    Else
        Password = pPassword
    End If
End Property

Public Property Let Password(ByVal newVal As String)
    If (pWasInstall Or pWasStarted) Then
        Err.Raise vbObjectError, App.EXEName, "Unable to change the Password property; Install or StartService has already been invoked."
    Else
        pPassword = newVal
        If (Not pIsLegacyOS) Then WinServe.NTService1.Password = newVal
    End If
End Property

Private Sub AddForm(ByRef FormOrHWnd)
    AddServiceForm FormOrHWnd, ObjPtr(Me)
End Sub

Private Sub RemoveForm(ByRef FormOrHWnd)
    RemoveServiceForm FormOrHWnd
End Sub

Friend Sub Procure()
    modMain.EnumAddServiceForms ObjPtr(Me)
End Sub

Public Function Install() As Boolean
    If Not RequiredVarsSet Then
        Err.Raise vbObjectError, App.EXEName, "You must set properties ServiceName, DisplayName and Description before using this function."
    ElseIf pWasStarted Then
        Err.Raise vbObjectError, App.EXEName, "StartService has been invoked, unable to initial properties safely to invoke Install or Uninstall."
    Else
        If (Not pIsLegacyOS) Then
            WinServe.NTService1.Install
            WinServe.NTService1.SaveSetting "", "Description", pDescription
            WinServe.NTService1.SaveSetting "", "ImagePath", AppEXE(False, True)
        Else
            Win98Svc.Install pServiceName, pDisplayName, pDescription, pInteractive
            Win98Svc.AutoRun pServiceName, pAutoStart
        End If
        pWasInstall = True
    End If
End Function

Public Function Uninstall() As Boolean
    If Not RequiredVarsSet Then
        Err.Raise vbObjectError, App.EXEName, "You must set properties ServiceName, DisplayName and Description before using this function."
    ElseIf pWasStarted Then
        Err.Raise vbObjectError, App.EXEName, "StartService has been invoked, unable to initial properties safely to invoke Install or Uninstall."
    Else
        If (Not pIsLegacyOS) Then
            WinServe.NTService1.Uninstall
        Else
            Win98Svc.Uninstall ServiceName
        End If
        pWasInstall = False
    End If
End Function

Public Function StartService()

    If Not RequiredVarsSet Then
        Err.Raise vbObjectError, App.EXEName, "You must set properties ServiceName, DisplayName and Description before using this function."
    ElseIf pWasInstall Then
        Err.Raise vbObjectError, App.EXEName, "Install or Uninstall has been invoked, unable to initial properties safely to invoke StartService."
    ElseIf ServiceFormExists(ProcessRunning(GetCurrentProcessId)) Then
        Err.Raise 8, "NTService20", "This program is already running as a service and can not occur in multiple instances."
    Else
        If Not Disabled Then
    
            pWasStarted = True
                
            If (Not pIsLegacyOS) Then
            
                AddForm WinServe.hwnd
                If WinServe.NTService1.Debug Then
                    WinServe_StartService pSvcRunning
                Else
                    WinServe.NTService1.StartService
                End If
            Else
                If ((GetSettingLong(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName, "NetStarting", 0) = 0) And AutoStart) Or _
                            (GetSettingLong(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName, "NetStarting", 0) = 1) Then
                    AddForm Win98Svc.hwnd
                    CallWindowProc AddressOf WindowProcedure, Win98Svc.hwnd, WM_NULL, Win98Svc.StartService(ServiceName), 0
                Else
                    Win98Svc.StopService
                    pSvcRunning = False
                End If
                SaveSettingLong HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName, "NetStarting", 0
                
            End If
            'If Not LCase(AppEXE(True, True)) = "vb6" Then
             '   Procure
          '  End If
        ElseIf (Not pIsLegacyOS) Then
           ' If Not LCase(AppEXE(True, True)) = "vb6" Then
            '    PostQuit WinServe
           ' End If
            If (LCase(Trim(GetFileTitle(ProcessRunning(CStr(GetCurrentProcessId))))) = "vb6") Then
                Err.Raise 8, App.EXEName, "This service is currently disabled and, if it was not for debug mode, normally would shutdown instead of this message."
            Else
                WinServe.NTService1.StopService
            End If
        ElseIf pIsLegacyOS Then
           ' If Not LCase(AppEXE(True, True)) = "vb6" Then
           '     PostQuit Win98Svc
           ' End If
            If (LCase(Trim(GetFileTitle(ProcessRunning(CStr(GetCurrentProcessId))))) = "vb6") Then
                Err.Raise 8, App.EXEName, "This service is currently disabled and, if it was not for debug mode, normally would shutdown instead of this message."
            Else
                Win98Svc.StopService
            End If
        End If
    End If

End Function
Friend Function StartWin98Service()
    StartWin98Service = Win98Svc.StartService(ServiceName)
End Function
Friend Property Get hwnd() As Long
    If (Not pIsLegacyOS) Then
        hwnd = WinServe.hwnd
    Else
        hwnd = Win98Svc.hwnd
    End If
End Property

Friend Function Disabled() As Boolean
    If (Not pIsLegacyOS) Then
        Disabled = (modRegistry.GetSettingLong(HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Services\" & pServiceName, "Start") = 4) Or (WinServe.NTService1.StartMode = svcStartDisabled)
    ElseIf pIsLegacyOS Then
        Disabled = Not (GetSettingLong(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName, "DisableExec", 0) = 0)
    End If
End Function

Friend Property Get ProcessID() As Long
    ProcessID = pProcIDCache
End Property

Private Sub Class_Initialize()

    pIsLegacyOS = IsWindows98
    pProcIDCache = GetCurrentProcessId
    
    pServiceName = ""
    pDisplayName = ""
    pDescription = ""

    pInteractive = False
    pAutoStart = True
    pAccount = ""
    pPassword = ""
    
    pWasInstall = False
    pWasStarted = False
    
    If (Not pIsLegacyOS) Then
        Set WinServe = New frmService
        Load WinServe
    Else
        Set Win98Svc = New LegacyOS
    End If

End Sub

Private Sub Class_Terminate()
        
    If (Not pIsLegacyOS) Then
        Unload WinServe
        Set WinServe = Nothing
    Else
        Set Win98Svc = Nothing
    End If
    
    IsWindows98
    
    Set classCallBacks = Nothing
End Sub
Private Sub RaiseCallback(ByVal CallbackEvent As SvcCallBacks)
    On Error Resume Next
    On Error GoTo errhandle:
    
    If Not classCallBacks Is Nothing Then
        If eventCallBacks(CallbackEvent) <> 0 Then
            RtlMoveMemory ByVal memCallBacks(0).Pointer, eventCallBacks(CallbackEvent), 4
            classCallBacks.ProcedureDefined1
            RtlMoveMemory ByVal memCallBacks(0).Pointer, memCallBacks(0).Address, 4
        End If
    End If
    
    Exit Sub
    On Error GoTo 0
errhandle:
    RtlMoveMemory ByVal memCallBacks(0).Pointer, memCallBacks(0).Address, 4
    If Err.Number <> -1 Then
        Err.Raise Err.Number, Err.Source, Err.Description
        Resume
    End If
End Sub
Friend Sub WinServe_LoggedOffService()
    If pWasStarted Then
    
    On Error GoTo handlemore:
    GoTo handleless:
handlemore:
    On Error Resume Next
    Err.Raise Err.Number, Err.Source, Err.Description
    On Error GoTo 0
handleless:

        RaiseCallback SvcCallBacks.SvcLoggedOffCallBack
        RaiseEvent LoggedOffService
    End If
End Sub

Private Sub WinServe_StartService(Success As Boolean)

    If Not pWasStarted Then StartService
    If pWasStarted Then
    
    On Error GoTo handlemore:
    GoTo handleless:
handlemore:
    On Error Resume Next
    Err.Raise Err.Number, Err.Source, Err.Description
    On Error GoTo 0
handleless:

        RaiseCallback SvcCallBacks.SvcStartCallBack
        If Err Then Err.Raise Err.Number, Err.Source, Err.Description
        RaiseEvent StartService(Success)
    End If
    If pWasStarted Then Success = True
End Sub

Private Sub WinServe_StopService()
    If pWasStarted Then

    modMain.ClearServiceForms
    On Error GoTo handlemore:
    GoTo handleless:
handlemore:
    On Error Resume Next
    Err.Raise Err.Number, Err.Source, Err.Description
    On Error GoTo 0
handleless:

        RaiseCallback SvcCallBacks.SvcStopCallBack
        RaiseEvent StopService
    End If
End Sub

Private Sub WinServe_ContinueService(Success As Boolean)
    If pWasStarted Then

    On Error GoTo handlemore:
    GoTo handleless:
handlemore:
    On Error Resume Next
    Err.Raise Err.Number, Err.Source, Err.Description
    On Error GoTo 0
handleless:

        RaiseCallback SvcCallBacks.SvcContinueCallBack
        RaiseEvent ContinueService(Success)
        Success = True
    End If
End Sub

Private Sub WinServe_Control(lEvent As Long)
    If pWasStarted Then

    On Error GoTo handlemore:
    GoTo handleless:
handlemore:
    On Error Resume Next
    Err.Raise Err.Number, Err.Source, Err.Description
    On Error GoTo 0
handleless:
    
        RaiseCallback SvcCallBacks.SvcControlCallBack
        RaiseEvent ControlService(lEvent)
    End If
End Sub

Private Sub WinServe_PauseService(Success As Boolean)
    If pWasStarted Then

    On Error GoTo handlemore:
    GoTo handleless:
handlemore:
    On Error Resume Next
    Err.Raise Err.Number, Err.Source, Err.Description
    On Error GoTo 0
handleless:

        RaiseCallback SvcCallBacks.SvcPauseCallBack
        RaiseEvent PauseService(Success)
        Success = True
    End If
End Sub

Friend Function Win98Svc_StartService() As Boolean
    If pWasStarted Then
    
    On Error GoTo handlemore:
    GoTo handleless:
handlemore:
    On Error Resume Next
    Err.Raise Err.Number, Err.Source, Err.Description
    On Error GoTo 0
handleless:

        Win98Svc_StartService = True
        RaiseCallback SvcCallBacks.SvcStartCallBack
        RaiseEvent StartService(Win98Svc_StartService)
    End If
End Function

Friend Sub Win98Svc_StopService()
    If pWasStarted Then
        Win98Svc.StopService
    
    modMain.ClearServiceForms
    On Error GoTo handlemore:
    GoTo handleless:
handlemore:
    On Error Resume Next
    Err.Raise Err.Number, Err.Source, Err.Description
    On Error GoTo 0
handleless:

        RaiseCallback SvcCallBacks.SvcStopCallBack
        RaiseEvent StopService
        pSvcRunning = False
    End If
End Sub




