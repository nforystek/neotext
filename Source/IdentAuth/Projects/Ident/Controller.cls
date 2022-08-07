VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Controller"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private pServiceName As String
Private pDisplayName As String
Private pDescription As String

Private pLoadOrderGroup As String
Private pInteractive As Boolean
Private pAutoStart As Boolean
Private pAccount As String
Private pPassword As String

Private pWasStarted As Boolean
Private pWasInstall As Boolean
Private pIsLegacyOS As Boolean
Private pSvcRunning As Boolean

Private WithEvents WinServe As frmService
Attribute WinServe.VB_VarHelpID = -1
Private Win98Svc As LegacyOS

Public Event UserLoggedOff()
Public Event ControlService(lEvent As Long)
Public Event ContinueService(Success As Boolean)
Public Event PauseService(Success As Boolean)
Public Event StartService(Success As Boolean)
Public Event StopService()

Private Declare Function ActiveEvent Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal lEvent As Long) As Long

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

Public Property Get AutoStart() As Boolean
    If (Not pIsLegacyOS) And ((Not pWasInstall) Or pWasStarted) Then
        AutoStart = (WinServe.NTService1.StartMode = svcStartAutomatic)
    Else
        AutoStart = Not (GetSettingLong(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName, "RegistryRun", 0) = 0)
    End If
End Property

Public Property Let AutoStart(ByVal newVal As Boolean)
    If (pWasInstall Or pWasStarted) Then
        Err.Raise vbObjectError, App.EXEName, "Unable to change the AutoStart property; Install or StartService has already been invoked."
    Else
        pAutoStart = newVal
        If (Not pIsLegacyOS) Then
            If pAutoStart Then
                WinServe.NTService1.StartMode = svcStartAutomatic
            Else
                WinServe.NTService1.StartMode = svcStartManual
            End If
        Else
            Win98Svc.AutoRun ServiceName, pAutoStart
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

'Public Sub Procure()
'    modService.EnumAddServiceForms ObjPtr(Me)
'End Sub

Public Function Install() As Boolean
    If Not RequiredVarsSet Then
        Err.Raise vbObjectError, App.EXEName, "You must set properties ServiceName, DisplayName and Description before using this function."
    ElseIf pWasStarted Then
        Err.Raise vbObjectError, App.EXEName, "StartService has been invoked, unable to initial properties safely to invoke Install or Uninstall."
    Else
        If (Not pIsLegacyOS) Then
            WinServe.NTService1.Install
            WinServe.NTService1.SaveSetting "", "Description", pDescription
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
    Else
        pWasStarted = True
            
        If (Not pIsLegacyOS) Then
            
            AddForm WinServe.hwnd
            If IsDebugger Then
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

Private Sub Class_Initialize()
    
    pIsLegacyOS = is9x

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

    modService.ClearServiceForms
        
    If (Not pIsLegacyOS) Then
        Unload WinServe
        Set WinServe = Nothing
    Else
        Set Win98Svc = Nothing
    End If
    
    modProcess.IsWindows98
End Sub

Friend Sub WinServe_UserLoggedOff()
    If pWasStarted Then RaiseEvent UserLoggedOff
End Sub

Private Sub WinServe_StartService(Success As Boolean)
    If Not pWasStarted Then StartService
    RaiseEvent StartService(Success)
    If pWasStarted Then Success = True
End Sub

Private Sub WinServe_StopService()
    If pWasStarted Then RaiseEvent StopService
End Sub

Private Sub WinServe_ContinueService(Success As Boolean)
    If pWasStarted Then RaiseEvent ContinueService(Success)
End Sub

Private Sub WinServe_Control(lEvent As Long)
   ' lEvent = ActiveEvent(AddressOf WindowProcedure, hWnd, lEvent)
End Sub

Private Sub WinServe_PauseService(Success As Boolean)
    If pWasStarted Then RaiseEvent PauseService(Success)
End Sub

Friend Function Win98Svc_StartService() As Boolean
    If pWasStarted Then RaiseEvent StartService(Win98Svc_StartService)
End Function

Friend Sub Win98Svc_StopService()
    If pWasStarted Then
        Win98Svc.StopService
        RaiseEvent StopService
        pSvcRunning = False
    End If
End Sub






