VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LegacyOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Option Compare Binary

Public hWnd As Long

Public Function StartService(ByVal ServiceName As String) As Boolean
    If Not (GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices", ServiceName, "") = "") Then
        StartService = (Not RegisterServiceProcess(GetCurrentProcessId, 1))
    End If
End Function

Public Sub StopService()
    RegisterServiceProcess GetCurrentProcessId, 0
End Sub

Public Function Install(ByVal ServiceName As String, ByVal DisplayName As String, ByVal Description As String, ByVal Interactive As Boolean) As Boolean
   
    SaveSettingString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", ServiceName, AppEXE(False, True)
    
    CreateKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName
    
    SaveSettingString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName, "DisplayName", DisplayName
    SaveSettingString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName, "Description", Description
    SaveSettingLong HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName, "Interactive", IIf(Interactive, 1, 0)
    SaveSettingLong HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName, "DisableExec", 0
    SaveSettingLong HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName, "NetStarting", 0
    
End Function

Public Sub AutoRun(ByVal ServiceName As String, ByVal Enabled As Boolean)
        
    SaveSettingLong HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName, "RegistryRun", IIf(Enabled, 1, 0)
    
End Sub

Public Sub Disable(ByVal ServiceName As String, ByVal Disable As Boolean)
        
    SaveSettingLong HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName, "DisableExec", IIf(Disable, 1, 0)
    
End Sub

Public Sub Uninstall(ByVal ServiceName As String)
  
    DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName, "DisplayName"
    DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName, "Description"
    DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName, "Interactive"
    DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName, "NetStarting"
    DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName, "RegistryRun"
    DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName, "DisableExec"
    
    DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName
    
    DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", ServiceName
    
End Sub

Private Sub Class_Initialize()
    hWnd = WindowInitialize(AddressOf WindowProcedure)
End Sub

Private Sub Class_Terminate()
    WindowTerminate hWnd
End Sub



