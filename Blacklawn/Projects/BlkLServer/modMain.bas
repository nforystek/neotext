Attribute VB_Name = "modMain"
#Const modMain = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module

#If Not (DLL = -1) Then
Public wService As clsServer
#End If

Public DisplayMsg As String
Public ListenPort As Long
Public ClientCount As Long

Public Sub Main()

#If Not (DLL = -1) Then
    Set wService = New clsServer
    If Not wService.Controller Is Nothing Then
        
        wService.Controller.ServiceName = "BlkLServer"
        wService.Controller.Account = ".\LocalSystem"
        wService.Controller.Password = "*"
        wService.Controller.DisplayName = "Blacklawn Server"
        wService.Controller.Description = "Multiplayer game server for Blacklawn."
        wService.Controller.AutoStart = False
        wService.Controller.Interactive = False
    
        Select Case Command
            Case "install"
                wService.Controller.Install
                Set wService = Nothing
            Case "uninstall"
                wService.Controller.Uninstall
                Set wService = Nothing
            Case Else
                LoadINI
    
                wService.Controller.StartService
#End If
    
#If (DLL = -1) Then
                LoadTxt "displaymsg = Welcome to SoSouiX.net's Blacklawn" & vbCrLf & _
                        "clientcount = 15" & vbCrLf & _
                        "listenport = False"
#End If
                Load frmBlkLServer
    
#If Not (DLL = -1) Then
        End Select
#End If
    End If
    
End Sub

Public Sub StopService()
#If Not (DLL = -1) Then
    Set wService = Nothing
#End If
    Unload frmBlkLServer

End Sub

Private Sub LoadINI()
    LoadTxt ReadFile(AppPath & "BlkLServer.ini")
End Sub

Public Sub LoadTxt(ByVal inData As String)
    Dim inLine As String
    
    Do Until (inData = "")
        inLine = RemoveNextArg(inData, vbCrLf)
        If Not ((Left(inLine, 1)) = ";") Then
        
            Select Case LCase(RemoveNextArg(inLine, "="))
                Case "clientcount"
                    If IsNumeric(inLine) Then ClientCount = CLng(inLine)
                Case "listenport"
                    If IsNumeric(inLine) Then ListenPort = CLng(inLine)
                Case "displaymsg"
                    DisplayMsg = Left(inLine, 512)
            End Select
        End If
    Loop
End Sub

