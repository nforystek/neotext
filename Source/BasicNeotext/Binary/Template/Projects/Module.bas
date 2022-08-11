Attribute VB_Name = "Module"
Option Private Module

Public MyService As Class

Public Sub Main()

    Set MyService = New Class
    MyService.Controller.ServiceName = "ServiceProject" 'this is how the system refers too
    MyService.Controller.Account = ".\LocalSystem" 'can be a user account and system
    MyService.Controller.Password = "*" 'this is for root accounts else use password
    MyService.Controller.DisplayName = "ServiceProject Test" 'this is a visual description
    MyService.Controller.Description = "A example of a service using a BasicNeotext Object"
    MyService.Controller.AutoStart = False 'whether it starts automatically on system
    MyService.Controller.Interactive = False 'if any form will be visible set to true

    Select Case Command
        Case "install"
            MyService.Controller.Install 'this puts it into the service registry
            Set MyService = Nothing 'install should be solo, deinitialize an end
        Case "uninstall"
            MyService.Controller.Uninstall 'remove us from services by ServiceName
            Set MyService = Nothing 'uninstall should be solo, deinitialize an end
        Case Else
            MyService.Controller.StartService 'called to ackknowledge start events
    End Select
    
End Sub

Public Sub StopService()
    'drop all objects initialized in the order they come out
    'conflict and then wont need to use End which will fault
    Set MyService = Nothing
End Sub
