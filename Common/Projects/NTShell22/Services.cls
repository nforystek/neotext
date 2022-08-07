VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Services"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary


Public Sub NetStart(ByVal ServiceName As String)
    Dim exePath As String
    If modProcess.IsWindows98 Then
        exePath = GetSettingString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", ServiceName, "")
    Else
        exePath = GetSettingExpand(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\" & ServiceName, "ImagePath", "")
    End If

    If exePath = "" Then
        Err.Raise 8, App.Title, "ServiceName not found or isn't installed as a service."
    ElseIf Trim(LCase(exePath)) = Trim(LCase(AppEXE(False, True))) Then
        Err.Raise 8, App.Title, "Unable to NetStart for the ServiceName with-in itself."
    ElseIf Not PathExists(exePath, True) Then
        Err.Raise 8, App.Title, "The path specified for ServiceName was found in-valid."
    Else
        modNTServices.NetStart ServiceName, GetFileName(exePath)
    End If
End Sub

Public Sub NetStop(ByVal ServiceName As String)
    Dim exePath As String
    If modProcess.IsWindows98 Then
        exePath = GetSettingString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", ServiceName, "")
    Else
        exePath = GetSettingExpand(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\" & ServiceName, "ImagePath", "")
    End If

    If exePath = "" Then
        Err.Raise 8, App.Title, "ServiceName not found or isn't installed as a service."
    ElseIf Trim(LCase(exePath)) = Trim(LCase(AppEXE(False, True))) Then
        Err.Raise 8, App.Title, "Unable to NetStop for the ServiceName with-in itself."
    ElseIf Not PathExists(exePath, True) Then
        Err.Raise 8, App.Title, "The path specified for ServiceName was found in-valid."
    Else
        modNTServices.NetStop ServiceName, GetFileName(exePath)
    End If
End Sub
