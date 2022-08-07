Attribute VB_Name = "modService"
#Const modService = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module
Public clsService As WindowsService
Public dbSettings As clsSettings

Public MaxEvents As New clsEventLog
Public ThreadManager As clsThreadManager

Public Const AppName = "Max-FTP Service"

Public Sub MaxDBError(ByVal Number As Long, ByVal Description As String, ByRef Retry As Boolean)

    If (Number = -2147467259) Or (Number = 3709) Then
        StopService
    Else
        Retry = True
    End If
End Sub

Public Function IsServiceFormStarted() As Boolean
    IsServiceFormStarted = False
    Dim frm As Form
    For Each frm In Forms
        If TypeName(frm) = "frmService" Then
            IsServiceFormStarted = True
            Exit For
        End If
    Next
End Function
Public Sub Main()

    Set dbSettings = New clsSettings
    Set clsService = New WindowsService
    If Not clsService.Controller Is Nothing Then
        
tryit: On Error GoTo catch
        
            
        With clsService.Controller
            .ServiceName = MaxServiceName
            .DisplayName = MaxServiceDisplay
            .Description = "Pools, maintains and manages scheduled sequential lists of synchronized file actions administrated by the incorporated Max-FTP interface."
            .Account = ".\LocalSystem"
            .Password = "*"
            .Interactive = False 'CBool(dbSettings.GetPublicSetting("ServiceInterface"))
            
        End With
    
        Dim hwnd As Long
        hwnd = FindWindow(vbNullString, ServiceFormCaption)
        If Not ((Not hwnd = 0) Or App.PrevInstance) Then
    
            If Not ExecuteFunction(Command) Then
                
                
                StartService
                
                clsService.Controller.StartService
                
                If Not clsService.ServiceTimer Is Nothing Then
                    clsService.ServiceTimer.Enabled = True
                Else
                    MaxEvents.AddEvent dbSettings, "Service Startup", AppPath, "Error: Unable to initialize scheduling."
                End If
                
            Else
                StopService
            End If
            
        ElseIf Command <> "" Then
            MessageQueueAdd ServiceFileName, Command
        End If
   ' Else'If Not dbSettings Is Nothing Then
      '  MaxEvents.AddEvent dbSettings, "Service Startup", AppPath, "Error: Unable to initialize controller."
    End If
    
GoTo final
catch: On Error GoTo 0
        
    'If Err Then MsgBox Err.Description, vbExclamation, App.EXEName
    
    StopService
    
final: On Error Resume Next
    
On Error GoTo -1
End Sub

Public Function StartService() As Boolean
        
    Load frmService
    LoadSchedule
    StartService = True

End Function
Public Sub StopService()

    If Not clsService Is Nothing Then
        If Not clsService.ServiceTimer Is Nothing Then
            clsService.ServiceTimer.Enabled = False
        End If
    End If
            
    UnloadSchedules
    
    Dim frm As Form
    For Each frm In Forms
        Unload frm
    Next
    
    Set dbSettings = Nothing
    
    Set clsService = Nothing
    
End Sub

Public Function ExecuteFunction(ByVal CommandLine As String)
    Dim hasCmd As Boolean
    If Trim(CommandLine) <> "" Then
        Dim InParams As String
        Dim InCOmmand As String
        
        CommandLine = Replace(Replace(Replace(CommandLine, "+", " "), "%20", " "), "%", " ")
        
        Do Until CommandLine = ""
        
            If InStr(CommandLine, vbCrLf) > 0 Then
                InParams = RemoveNextArg(CommandLine, vbCrLf)
            ElseIf InStr(CommandLine, "|") > 0 Then
                InParams = RemoveNextArg(CommandLine, "|")
            ElseIf Left(CommandLine, 1) = "/" Then
                CommandLine = Mid(CommandLine, 2)
                InParams = RemoveNextArg(CommandLine, "/")
            ElseIf Left(CommandLine, 1) = "-" Then
                CommandLine = Mid(CommandLine, 2)
                InParams = RemoveNextArg(CommandLine, "-")
            Else
                InParams = CommandLine
                CommandLine = ""
            End If
            
            InCOmmand = LCase(RemoveNextArg(InParams, " "))
            
            hasCmd = True
            Select Case Replace(InCOmmand, " ", "")
            
                Case "interactive"
                    clsService.Controller.Uninstall
                    If IsNumeric(InParams) Then
                        clsService.Controller.Interactive = CBool(InParams)
                        dbSettings.SetPublicSetting "ServiceInterface", CBool(InParams)
                    End If
                    clsService.Controller.Install
                    
                Case "startup"
                    clsService.Controller.Uninstall
                    clsService.Controller.AutoStart = CBool(InParams)
                    clsService.Controller.Install
                Case "install"
                    If IsNumeric(InParams) Then
                        clsService.Controller.AutoStart = CBool(InParams)
                    End If
                    clsService.Controller.Install
                Case "uninstall"
                    clsService.Controller.Uninstall

                Case "stopschedule"
                    If IsNumeric(InParams) Then
                        If ScheduleExists(CLng(InParams)) Then
                            GetSchedule(CLng(InParams)).StopSchedule
                        End If
                    End If
                Case "runschedule"
                    If IsNumeric(InParams) Then
                        If ScheduleExists(CLng(InParams)) Then
                            GetSchedule(CLng(InParams)).RunSchedule
                        Else
                            LoadSchedule CLng(InParams)
                            If ScheduleExists(CLng(InParams)) Then
                                GetSchedule(CLng(InParams)).RunSchedule
                            End If
                            UnloadSchedule CLng(InParams)
                        End If
                    End If
                Case "copy", "move", "delete", "folder", "rename", "script"
                    LoadCmdLine InCOmmand, InParams
                    
                Case Else
                    hasCmd = False
                    
            End Select

        Loop
        
        If ScheduleExists(-1) Then
            Dim cmdSchedule As Schedule
            Set cmdSchedule = GetSchedule(-1)
            cmdSchedule.RunSchedule
            cmdSchedule.ClearOperations
            Set cmdSchedule = Nothing
        End If
        
    End If

    ExecuteFunction = hasCmd
    
End Function

Public Function ProcessMessage()

    If (MessageQueueLog(ServiceFileName) > 0) Then
    
        Dim Messages As New VBA.Collection
        Dim Msg
        Set Messages = MessageQueueGet(ServiceFileName)
        
        For Each Msg In Messages

            If Trim(Msg) <> "" Then
                Dim InParams As String
                Dim InCOmmand As String
                
                Do Until Msg = ""
                
                    If InStr(Msg, vbCrLf) > 0 Then
                        InParams = RemoveNextArg(Msg, vbCrLf)
                    ElseIf Left(Msg, 1) = "/" Then
                        Msg = Mid(Msg, 2)
                        InParams = RemoveNextArg(Msg, "/")
                    ElseIf Left(Msg, 1) = "-" Then
                        Msg = Mid(Msg, 2)
                        InParams = RemoveNextArg(Msg, "-")
                    Else
                        InParams = LCase(Msg)
                        Msg = ""
                    End If
                    
                    InCOmmand = LCase(RemoveNextArg(InParams, " "))
                    
                    Select Case InCOmmand

                        Case "runschedule"
                            If ScheduleExists(CLng(InParams)) Then
                                GetSchedule(CLng(InParams)).RunSchedule
                            End If
                        Case "stopschedule"
                            If ScheduleExists(CLng(InParams)) Then
                                GetSchedule(CLng(InParams)).StopSchedule
                            End If
                        Case "loadschedules"
                            LoadSchedule
                        Case "loadschedule"
                            LoadSchedule CLng(InParams)
                        Case "unloadschedule"
                            UnloadSchedule CLng(InParams)
                        Case "unloadschedules"
                            UnloadSchedules

                    End Select
        
                Loop
            End If
        Next
        Set Messages = Nothing
        
    End If

End Function

Public Function IsEmptyArray(ByRef InArray() As String) As Boolean
    On Error Resume Next
    Dim Test As String
    Test = InArray(0)
    If Err = 0 And Not Test = "" Then
        IsEmptyArray = False
    Else
        IsEmptyArray = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

Public Function IsFileOnArray(ByRef InArray() As String, ByVal FileName As String, Optional ByRef FileSize As String, Optional ByRef FileDate As String) As Boolean
    Dim retVal As Boolean
    FileSize = ""
    FileDate = ""
    
    retVal = False
    If Not InArray(0) = "" Then
        Dim cnt As Long
        Dim inFile As String
        cnt = 0
        Do
            inFile = InArray(cnt)
            If LCase(Trim(RemoveNextArg(inFile, "|"))) = LCase(Trim(FileName)) Then
                FileSize = RemoveNextArg(inFile, "|")
                FileDate = RemoveNextArg(inFile, "|")
                retVal = True
            End If
            cnt = cnt + 1
        Loop Until cnt > UBound(InArray) Or retVal
    End If

    IsFileOnArray = retVal
End Function
