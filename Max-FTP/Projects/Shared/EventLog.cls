VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEventLog"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
                
'**************************************
' Name: logging events - extended functionality
' Description:This is a module with one function that will give you some more functionality than what you can do with the App.LogEvent method in VB6.
'1. You will be able to specify the EventLog (Application, Security or System)
'2. You are able to specify the Source (Your own Application identifier instead of the VBRuntime)
'This module is also used in the ExeptionHandler dll. that is posted earlier today.
' By: Edwin Vermeer. (from psc cd)
'**************************************

' Using this module is easy. Just Call it like this:
' writelog EventLog_Application,"My Special APP",vbLogEventTypeError, "Oep, Something went wrong :)"
'Functions and type for logging events
Option Explicit
'TOP DOWN
Private LogRootPath As String

Private Declare Function RegisterEventSource Lib "ADVAPI32.DLL" Alias "RegisterEventSourceA" (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As Long
Private Declare Function DeregisterEventSource Lib "ADVAPI32.DLL" (ByVal hEventLog As Long) As Long
Private Declare Function ReportEvent Lib "ADVAPI32.DLL" Alias "ReportEventA" (ByVal hEventLog As Long, ByVal wType As Integer, ByVal wCategory As Integer, ByVal dwEventID As Long, ByVal lpUserSid As Any, ByVal wNumStrings As Integer, ByVal dwDataSize As Long, plpStrings As Long, lpRawData As Any) As Boolean
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Enum EventLog
    EventLog_Application
    EventLog_Security
    EventLog_System
End Enum
' This function will give you some more functionality over the App.LogEvent method.
' You will be able to specify the EventLog (Application, Security or System)
' And you are able to specify the Source (Your own Application identifier instead of the VBRuntime)
Friend Function WriteLog(intEventLogID As EventLog, strEventSource As String, intEventType As LogEventTypeConstants, strEventString As String) As Boolean
1 On Error GoTo ErrHandler
2     Dim intEventStringsCount As Integer
3     Dim hEventLog As Long
4     Dim hMsgs As Long
5     Dim lngEventStringSize As Long
6     Dim objRegistry As Object
7     Dim strEventLogDescription As String
8     WriteLog = False
    ' In case we have a new source we make sure it finds the VBRuntime DLL for handeling the event description.
9     Select Case intEventLogID
    Case EventLog_Application
10     strEventLogDescription = "Application"
11     Case EventLog_Security
12     strEventLogDescription = "Security"
13     Case EventLog_System
14     strEventLogDescription = "System"
15     End Select
16     Set objRegistry = CreateObject("Wscript.Shell")
17     objRegistry.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\EventLog\" & strEventLogDescription & "\" & strEventSource & "\EventMessageFile", objRegistry.regread("HKLM\SYSTEM\CurrentControlSet\Services\EventLog\Application\VBRuntime\EventMessageFile"), "REG_SZ"
18     objRegistry.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\EventLog\" & strEventLogDescription & "\" & strEventSource & "\TypesSupported", 4, "REG_DWORD"
    ' Set the event source and report the event
19     hEventLog = RegisterEventSource("", strEventSource)
20     strEventString = ":" & vbCrLf & vbCrLf & strEventString
21     lngEventStringSize = Len(strEventString) + 1
22     hMsgs = GlobalAlloc(&H40, lngEventStringSize)
23     CopyMemory ByVal hMsgs, ByVal strEventString, lngEventStringSize
24     intEventStringsCount = 1
25     If ReportEvent(hEventLog, intEventType, 0, 1, GetUserSIDByProcessID(GetCurrentProcessId), intEventStringsCount, lngEventStringSize, hMsgs, hMsgs) = 0 Then
26     WriteLog = True
27     End If
28     GlobalFree (hMsgs)
29     DeregisterEventSource (hEventLog)
30     Exit Function
31 ErrHandler:
End Function

Private Function GetLogFilename() As String

    GetLogFilename = LogRootPath & "ex" & Padding(2, Right(Year(Now), 2), "0") & Padding(2, Month(Now), "0") & Padding(2, Day(Now), "0") & ".log"
'    Dim fNum As Long
'    fNum = FreeFile
'    If Not PathExists(GetLogFilename, True) Then
'        Open GetLogFilename For Output As #fNum
'        Close #fNum
'        Open GetLogFilename For Append Shared As #fNum
'            Print #fNum, "#Software: Max-FTP Client"
'            Print #fNum, "#Version: 6.1.0"
'            Print #fNum, "#Date: " & Year(Now) & "-" & Month(Now) & "-" & Day(Now) & " " & Format(Now, "hh:mm:ss")
'        Close #fNum
'    End If
End Function
Public Function GetEvents() As VBA.Collection
    Dim fNum As Long
    Dim strEvent As String
    Set GetEvents = New VBA.Collection
    
    'If PathExists(GetLogFilename, True) Then
        Dim strFile As String
        strFile = Dir(Left(LogRootPath, Len(LogRootPath) - 2) & "C\*.log")
        Do While PathExists(Left(LogRootPath, Len(LogRootPath) - 2) & "C\" & strFile, True)
            fNum = FreeFile
            Open Left(LogRootPath, Len(LogRootPath) - 2) & "C\" & strFile For Input As #fNum
                Do Until EOF(fNum)
                    Line Input #fNum, strEvent
                    If Not Left(strEvent, 1) = "#" Then
                        GetEvents.Add strEvent
                    End If
                Loop
            Close #fNum
            strFile = Dir
        Loop
        strFile = Dir(Left(LogRootPath, Len(LogRootPath) - 2) & "S\*.log")
        Do While PathExists(Left(LogRootPath, Len(LogRootPath) - 2) & "S\" & strFile, True)
            fNum = FreeFile
            Open Left(LogRootPath, Len(LogRootPath) - 2) & "S\" & strFile For Input As #fNum
                Do Until EOF(fNum)
                    Line Input #fNum, strEvent
                    If Not Left(strEvent, 1) = "#" Then
                        GetEvents.Add strEvent
                    End If
                Loop
            Close #fNum
            strFile = Dir
        Loop
    'End If
    
End Function
'c-ip cs-username s-computername s-ip s-port s-uri s-method d-ip d-port d-url

Public Sub AddEvent(ByRef dbSettings As clsSettings, ByVal Source As String, ByVal Location As String, ByVal Message As String, Optional ByRef sClient As Object = Nothing, Optional ByRef dClient As Object = Nothing)

    Dim doLog As Boolean
    If dbSettings Is Nothing Then
        doLog = True
    Else

        Select Case LCase(Trim(App.EXEName))
            Case "maxservice"
                doLog = CBool(dbSettings.GetProfileSetting("EventLog"))
            Case "maxftp"
                doLog = CBool(dbSettings.GetClientSetting("EventLog"))
            Case "maxide"
                doLog = True
          End Select
    End If
    
    If Not dClient Is Nothing Then
        If dClient.GetLastError <> "" Then
            Message = " [Destination Error: " & dClient.GetLastError & "]"
        End If
    End If
    If Not sClient Is Nothing Then
        If sClient.GetLastError <> "" Then
            Message = Message & " [Source Error: " & sClient.GetLastError & "]"
        End If
    End If
    If Err.Number <> 0 Then
        Message = Message & " [Program Error" & Err.Description & "]"
    End If
    
    If doLog Then
    
        Dim newHeader As String
        Dim strEvents As String
        strEvents = Year(Now) & "-" & Month(Now) & "-" & Day(Now) & "," & Format(Now, "hh:mm:ss") & "," & GetUserLoginName & "," & GetMachineName
        
        If Not Trim(Message) = "" Then
            strEvents = strEvents & ",""" & Replace(Message, """", """""") & """"
        Else
            strEvents = strEvents & ","
        End If
        newHeader = newHeader & ",cs-method"

        newHeader = "date,time,username,computername,method,l-ip,l-port,l-uri-stem,l-uri-path,r-ip,r-port,r-uri-stem,r-uri-path"
        If Not sClient Is Nothing Then
            strEvents = strEvents & "," & sClient.Server & "," & sClient.Port & ",""" & Replace(sClient.URL, """", """""") & """,""" & Replace(sClient.Folder, """", """""") & """"
        Else
            strEvents = strEvents & ",,,,"
        End If

        If Not dClient Is Nothing Then
            strEvents = strEvents & "," & dClient.Server & "," & dClient.Port & ",""" & Replace(dClient.URL, """", """""") & """,""" & Replace(dClient.Folder, """", """""") & """"
        Else
            strEvents = strEvents & ",,,,"
        End If
        
        Dim fNum As Long
        fNum = FreeFile
        If Not PathExists(GetLogFilename, True) Then
            Open GetLogFilename For Output As #fNum
            Print #fNum, newHeader
            Close #fNum
        End If
        
        fNum = FreeFile
        Open GetLogFilename For Append Shared As #fNum
            Print #fNum, strEvents
        Close #fNum
        
        
    End If
    
    If Not dbSettings Is Nothing Then
        If isNT Then
        
            If dbSettings.GetPublicSetting("ServiceEventLog") Then
                    WriteLog EventLog_Application, App.Title, IIf(LCase(Left(Trim(Message), 5)) = "error", vbLogEventTypeWarning, vbLogEventTypeInformation), "[" & vbCrLf & vbTab & Source & "; " & vbCrLf & vbTab & Location & ";" & vbCrLf & vbTab & Message & ";" & vbCrLf & "]" & vbCrLf
            End If
        End If
    End If
End Sub

Public Sub ClearEvents()

    Dim strFile As String
    strFile = Dir(Left(LogRootPath, Len(LogRootPath) - 2) & "C\*.log")
    Do While PathExists(Left(LogRootPath, Len(LogRootPath) - 2) & "C\" & strFile, True)
        Kill Left(LogRootPath, Len(LogRootPath) - 2) & "C\" & strFile
    Loop
    strFile = Dir(Left(LogRootPath, Len(LogRootPath) - 2) & "S\*.log")
    Do While PathExists(Left(LogRootPath, Len(LogRootPath) - 2) & "S\" & strFile, True)
        Kill Left(LogRootPath, Len(LogRootPath) - 2) & "S\" & strFile
        strFile = Dir
    Loop
    
End Sub

Private Sub Class_Initialize()
    If Not PathExists(SysPath & "LogFiles", False) Then MkDir SysPath & "LogFiles"
    If Not PathExists(SysPath & "LogFiles\NTFTP" & IIf(LCase(Trim(App.EXEName)) = "maxftp", "C", "S"), False) Then
        MkDir SysPath & "LogFiles\NTFTP" & IIf(LCase(Trim(App.EXEName)) = "maxftp", "C", "S")
    End If
    LogRootPath = SysPath & "LogFiles\NTFTP" & IIf(LCase(Trim(App.EXEName)) = "maxftp", "C", "S") & "\"

End Sub


