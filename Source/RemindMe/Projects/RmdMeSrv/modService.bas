#Const [True] = -1
#Const [False] = 0
Attribute VB_Name = "modService"
#Const modService = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module

Public DBConn As clsDBConnection
Public dbSettings As clsDBSettings
Public wService As WindowsService
Public Operations As clsOperationList

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Any) As Long

Public Sub MethodInvoke()

    Dim hwnd As Long
    hwnd = WindowInitialize
  
    SetTimer hwnd, 0, 3, AddressOf MethodCallback
    
End Sub

Public Sub MethodCallback(ByVal hwnd As Long, ByVal uMsg As Long, ByVal mTimerID As Long, ByVal dwTime As Long)

    KillTimer hwnd, mTimerID
    WindowTerminate hwnd
    
    ProcessMessage
    
End Sub

Public Sub Main()
tryit: On Error GoTo catch
    
'    Dim hwnd As Long
'    hwnd = FindWindow(vbNullString, "RemindMe Service")
    If App.PrevInstance Then
        End
    Else
    
        Set wService = New WindowsService
        With wService.Controller
            .ServiceName = ServiceName
            .DisplayName = "RemindMe Operations"
            .Description = "Pools maintains and manages scheduled script operations incorporated with the RemindMe interface."
            .Account = ".\LocalSystem"
            .Password = "*"
            .Interactive = True
            .AutoStart = True
        End With
    
        If Not ExecuteFunction(Command) Then
        
            StartService
            
        End If
    
    End If

GoTo final
catch: On Error GoTo 0

    'If Err Then MsgBox Err.Description, vbExclamation, App.EXEName

final: On Error Resume Next

On Error GoTo -1
End Sub

Public Sub StartService()
    
    Load frmService
    
    wService.Controller.StartService
    
    Set DBConn = New clsDBConnection
    Set dbSettings = New clsDBSettings
    Set Operations = New clsOperationList

    frmService.UpdateScript
    Operations.Load
    
    wService.ServiceTimer.Enabled = True
    
End Sub

Public Sub StopService()

    wService.ServiceTimer.Enabled = False

    Operations.Clear
    frmService.EndScript

    Set Operations = Nothing
    Set dbSettings = Nothing
    Set DBConn = Nothing
    
    Set wService = Nothing
    
    Unload frmService
    
End Sub

Public Function ExecuteFunction(ByVal CommandLine As String) As Boolean
    Dim HasCmd As Boolean
    
    If Trim(CommandLine) <> "" Then
        Dim InParams As String
        Dim InCommand As String
        CommandLine = Replace(Replace(Replace(CommandLine, "+", " "), "%20", " "), "%", " ")
        
        Do Until CommandLine = ""
        
            If InStr(CommandLine, vbCrLf) > 0 Then
                InParams = RemoveNextArg(CommandLine, vbCrLf)
            ElseIf Left(CommandLine, 1) = "/" Then
                CommandLine = Mid(CommandLine, 2)
                InParams = RemoveNextArg(CommandLine, "/")
            ElseIf Left(CommandLine, 1) = "-" Then
                CommandLine = Mid(CommandLine, 2)
                InParams = RemoveNextArg(CommandLine, "-")
            Else
                InParams = LCase(CommandLine)
                CommandLine = ""
            End If
            
            InCommand = LCase(RemoveNextArg(InParams, " "))
            
            HasCmd = True
            Select Case Replace(InCommand, " ", "")
                Case "install"
                    wService.Controller.Install
                Case "uninstall"
                    wService.Controller.Uninstall
                Case Else
                    
                    HasCmd = False
            End Select
        Loop
    End If
    
    ExecuteFunction = HasCmd
End Function

Public Function ProcessMessage()
   
    If dbSettings.MessageWaiting(ServiceFileName) Then
        Dim NextMsg As String
        Dim Msgs As Collection
        Set Msgs = dbSettings.MessageQueue(ServiceFileName)
        
        Do Until Msgs.Count = 0
            NextMsg = CStr(Msgs(1))
            
            Dim inCmd As String
            Dim inParam As String
            
            inParam = NextMsg
            inCmd = RemoveNextArg(inParam, ":")
            
            Select Case LCase(inCmd)
                Case "updateoperation"
                    Operations.Load CLng(inParam)
                    
                Case "removeoperation"
                    Operations.Remove CLng(inParam)
                    
                Case "updateoperations"
                    Operations.Clear
                    Operations.Load
                    
                Case "updatescript"
                    frmService.UpdateScript
                    
                Case "stopoperations"
                    Operations.StopOperations
                    
                Case "startoperation"
                    Operations.StartOperation CLng(inParam)
                    
            End Select
        
            Msgs.Remove 1
        Loop
        Set Msgs = Nothing
    End If
    
End Function

Public Function SendMessage(ByVal Message As String)
    If (ProcessRunning(RemindMeFileName) > 0) Then
        Dim rs As New ADODB.Recordset
        DBConn.rsQuery rs, "INSERT INTO MessageQueue (MessageTo, MessageText) VALUES ('" & RemindMeFileName & "','" & Message & "');"
        If Not rs.State = 0 Then rs.Close
        Set rs = Nothing
    End If
End Function
