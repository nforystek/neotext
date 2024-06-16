VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmProjectCompiler 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1050
   ControlBox      =   0   'False
   Icon            =   "frmProjectCompiler.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   1050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   210
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
      UseSafeSubset   =   -1  'True
   End
End
Attribute VB_Name = "frmProjectCompiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public Event Started()
Public Event Finished()

Private nGUID As String
Private nFull As String

Private nDims As String
Private nBase As String
Private nInit As String
Private nExec As String
Private nTerm As String
Private nMain As String

Private HasExec As String
Private HasInit As String
Private HasTerm As String

Private HasCall As String

Private Objects As New Collection

Private WithEvents timMe As NTSchedule20.Timer
Attribute timMe.VB_VarHelpID = -1

Public Function SafeStr(ByVal nStr As String, Optional ByVal NoQuote As Boolean = False) As String
    Select Case Project.Language
        Case "VBScript"
            nStr = Replace(nStr, """", """""")
            nStr = Replace(nStr, vbCrLf, """ & vbCrLf & """)
            nStr = Replace(nStr, vbCr, """ & vbCr & """)
            nStr = Replace(nStr, vbLf, """ & vbLf & """)
        Case "JScript"
            nStr = Replace(nStr, "'", "\'")
            nStr = Replace(nStr, vbCrLf, "\r\n")
            nStr = Replace(nStr, vbCr, "\r")
            nStr = Replace(nStr, vbLf, "\n")
    End Select
    If Not NoQuote Then
        SafeStr = QuoteStr(nStr)
    Else
        SafeStr = nStr
    End If
End Function

Public Function QuoteStr(ByVal nStr As String) As String
    Select Case Project.Language
        Case "VBScript"
            nStr = """" & nStr & """"
        Case "JScript"
            nStr = "'" + nStr + "'"
    End Select
    QuoteStr = nStr
End Function

Private Function FindLine(ByRef cProject As clsProject, ByVal nText As String, ByRef nLine As String) As Long
    Dim pos As Long
    Dim tst As Long
    
    Do
        pos = InStr(tst + 1, nText, nLine, vbTextCompare)
        If pos > 0 Then

            If (cProject.Language = "VBScript") Then
                tst = InStrRev(nText, vbCrLf, pos)
                If (InStrRev(nText, "'", pos) > tst) Then
                    tst = pos
                ElseIf (InStrRev(nText, """", pos) > tst) Then
                    tst = pos
                Else
                    tst = -1
                End If
            Else
                If (InStrRev(nText, "/*", pos) > 0) Then
                    If InStrRev(nText, "*/", pos) = 0 Then
                        tst = pos + 1
                    ElseIf InStr(pos, nText, "*/") > 0 Then
                        tst = pos + 1
                    Else
                        tst = -1
                    End If
                Else
                    tst = -1
                End If
                
                If tst = -1 Then
                    tst = InStrRev(nText, vbCrLf, pos)
                    If (InStrRev(nText, "//", pos) > tst) Then
                        tst = pos + 1
                    ElseIf (InStrRev(nText, """", pos) > tst) Then
                        tst = pos
                    ElseIf (InStrRev(nText, "'", pos) > tst) Then
                        tst = pos
                    Else
                        tst = -1
                    End If
                End If
                
            End If
        Else
            tst = 0
        End If
    Loop Until tst < 1
    If tst = -1 Then
        nLine = Mid(nText, pos, Len(nLine))
        FindLine = pos
    Else
        FindLine = 0
    End If
End Function

Private Function EqualArguments(ByVal ArgList1 As String, ByVal ArgList2 As String) As Boolean

    Dim para1 As String
    Dim para2 As String
    
    ArgList1 = TrimStrip(Trim(LCase(ArgList1)), vbTab)
    If (Left(ArgList1, 1) = "(" And Right(ArgList1, 1) = ")") Then
        ArgList1 = Mid(ArgList1, 2, Len(ArgList1) - 2)
    End If
    
    ArgList2 = TrimStrip(Trim(LCase(ArgList2)), vbTab)
    If (Left(ArgList2, 1) = "(" And Right(ArgList2, 1) = ")") Then
        ArgList2 = Mid(ArgList2, 2, Len(ArgList2) - 2)
    End If
    
    If (Not (ArgList1 = "")) And (Not (ArgList2 = "")) Then
        
        Do
            para1 = TrimStrip(Trim(RemoveNextArg(ArgList1, ",")), vbTab)
            para2 = TrimStrip(Trim(RemoveNextArg(ArgList2, ",")), vbTab)
        Loop Until (ArgList1 = "") Or (Not (para1 = para2))

        EqualArguments = ((ArgList1 = "") And (ArgList2 = "")) And (para1 = para2)
    Else
        EqualArguments = (ArgList1 = ArgList2)
    End If
    
End Function

Private Function ConvertEvent(ByRef cProject As clsProject, ByVal EventIdentity As String, ByVal ObjectParameters As String, ByVal ObjectIdentity As String, ByRef nText As String, Optional ByRef nType As Long = 0, Optional ByRef nCall As String = "*") As String
    Dim pos As Long
    pos = FindLine(cProject, nText, EventIdentity)
        
    If pos > 0 Then
    
        Dim Line As Long
        Line = CountWord(Left(nText, pos), vbCrLf) + 1
        
        Dim Index As Long
        Index = InStrRev(nText, vbCrLf, pos, vbTextCompare)
        If Index = 0 Then
            Index = 1
        Else
            Index = Index + 2
        End If
        
        Dim g As String
        g = "a" & Replace(GUID, "-", "")
       
        If Not (ObjectParameters = "") Then
            Dim asline As String
            Dim Params As String
            Dim tested As Long
            
            asline = Mid(nText, Index)
            tested = InStr(asline, vbCrLf)
            If tested > 0 Then asline = Left(asline, tested - 1)
            tested = (pos - Index) + Len(EventIdentity)
            
            Params = Replace(Replace(asline, EventIdentity, "", , , vbTextCompare), "function ", "", , , vbTextCompare)
            Params = TrimStrip(Trim(Params), vbTab)
            If cProject.Language = "JScript" Then
                If InStr(Params, "{") > 0 Then Params = Left(Params, InStr(Params, "{") - 1)
                Params = TrimStrip(Trim(Params), vbTab)
            End If
            
            If Not Left(Params, 1) = "(" Then
                Err.Raise 440, "#" & ObjectIdentity & ":" & Line & ":" & tested & ":" & asline, "Automation error: Expected '('"
            ElseIf Not Right(Params, 1) = ")" Then
                Err.Raise 440, "#" & ObjectIdentity & ":" & Line & ":" & Len(asline) & ":" & asline, "Automation error: Expected ')'"
            ElseIf Not EqualArguments(ObjectParameters, Params) Then
                Err.Raise 450, "#" & ObjectIdentity & ":" & Line & ":" & tested & ":" & asline, "Wrong number of, or Invalid; arguments"
            End If
        
        End If
        
        nText = Left(nText, pos - 1) & Replace(nText, EventIdentity, g, pos, 1, vbTextCompare)
        Select Case cProject.Language
            Case "VBScript"
            
                If (Not (nCall = "*")) And (Not (nType = 0)) Then
                    If nType = 1 Then
                        nCall = nCall & vbTab & g & " " & ObjectIdentity & vbCrLf
                    ElseIf nType = 2 Then
                        nCall = vbTab & g & " " & ObjectIdentity & vbCrLf & nCall
                    ElseIf nType = 3 Then
                        nCall = vbTab & g & " """ & ObjectIdentity & """" & vbCrLf & nCall
                    ElseIf nType = 4 Then
                        nCall = nCall & vbTab & g & " """ & ObjectIdentity & """" & vbCrLf
                    ElseIf nType = 5 Then
                        nCall = nCall & vbTab & g & vbCrLf
                    End If
                End If
        
            Case "JScript"
                If (Not (nCall = "*")) And (Not (nType = 0)) Then
                    If nType = 1 Then
                        nCall = nCall & vbTab & g & "(" & ObjectIdentity & ");" & vbCrLf
                    ElseIf nType = 2 Then
                        nCall = vbTab & g & "(" & ObjectIdentity & ");" & vbCrLf & nCall
                    ElseIf nType = 3 Then
                        nCall = vbTab & g & "(""" & ObjectIdentity & """);" & vbCrLf & nCall
                    ElseIf nType = 4 Then
                        nCall = nCall & vbTab & g & "(""" & ObjectIdentity & """);" & vbCrLf
                    ElseIf nType = 5 Then
                        nCall = nCall & vbTab & g & "();" & vbCrLf
                    End If
                End If
                    
        End Select
        
        If nType = 4 Then
            nType = Line

            pos = FindLine(cProject, nText, EventIdentity)
            If (pos > 0) Then
                nText = Left(nText, pos - 1) & Replace(nText, EventIdentity, IIf(cProject.Language = "VBScript", g, "return "), pos, 1, vbTextCompare)
            End If
        End If
        
        ConvertEvent = g
    Else
        ConvertEvent = ""
    End If

End Function
Private Function NewCollect(ByRef cProject As clsProject, ByVal nName As String, ByRef nText As String, Optional ByVal AddObjects As Boolean = True) As objCollect
    Dim NewObj As New objCollect
    With NewObj
        If AddObjects Then .InitObject nName, Me
    End With
    Set NewCollect = NewObj
    Set NewObj = Nothing
End Function

Private Function NewTimer(ByRef cProject As clsProject, ByVal nName As String, ByRef nText As String, Optional ByVal AddObjects As Boolean = True) As objTimer
    Dim NewObj As New objTimer
    With NewObj
        .HasOnTicking = ConvertEvent(cProject, "Object:OnTicking", "(Self)", nName, nText)
        If AddObjects Then .InitObject nName, Me
    End With
    Set NewTimer = NewObj
    Set NewObj = Nothing
End Function
Private Function NewSchedule(ByRef cProject As clsProject, ByVal nName As String, ByRef nText As String, Optional ByVal AddObjects As Boolean = True) As objSchedule
    Dim NewObj As New objSchedule
    With NewObj
        .HasScheduledEvent = ConvertEvent(cProject, "Object:ScheduledEvent", "(Self)", nName, nText)
        If AddObjects Then .InitObject nName, Me
    End With
    Set NewSchedule = NewObj
    Set NewObj = Nothing
End Function
Private Function NewClient(ByRef cProject As clsProject, ByVal nName As String, ByRef nText As String, Optional ByVal AddObjects As Boolean = True) As objClient
    Dim NewObj As New objClient
    With NewObj
        .HasError = ConvertEvent(cProject, "Object:Error", "(Self, Number, Source, Description)", nName, nText)
        .HasLogMessage = ConvertEvent(cProject, "Object:LogMessage", "(Self, MessageType, AddedText)", nName, nText)
        .HasItemListing = ConvertEvent(cProject, "Object:ItemListing", "(Self, ItemName, ItemSize, ItemDate, ItemAccess)", nName, nText)
        .HasDataProgress = ConvertEvent(cProject, "Object:DataProgress", "(Self, ProgressType, ReceivedBytes)", nName, nText)
        .HasDataComplete = ConvertEvent(cProject, "Object:DataComplete", "(Self, ProgressType)", nName, nText)
        If AddObjects Then .InitObject nName, Me
    End With
    Set NewClient = NewObj
    Set NewObj = Nothing
End Function
Private Function NewSocket(ByRef cProject As clsProject, ByVal nName As String, ByRef nText As String, Optional ByVal AddObjects As Boolean = True) As objSocket
    Dim NewObj As New objSocket
    With NewObj
        .HasConnected = ConvertEvent(cProject, "Object:Connected", "(Self)", nName, nText)
        .HasDataArriving = ConvertEvent(cProject, "Object:DataArriving", "(Self)", nName, nText)
        .HasConnection = ConvertEvent(cProject, "Object:Connection", "(Self, Handle)", nName, nText)
        .HasDisconnected = ConvertEvent(cProject, "Object:Disconnected", "(Self)", nName, nText)
        .HasSendComplete = ConvertEvent(cProject, "Object:SendComplete", "(Self)", nName, nText)
        .HasError = ConvertEvent(cProject, "Object:Error", "(Self, Number, Source, Description)", nName, nText)
        If AddObjects Then .InitObject nName, Me
    End With
    Set NewSocket = NewObj
    Set NewObj = Nothing
End Function

Private Function NewUUCode(ByRef cProject As clsProject, ByVal nName As String, ByRef nText As String, Optional ByVal AddObjects As Boolean = True) As objUUCode
    Dim NewObj As New objUUCode
    With NewObj
        If AddObjects Then .InitObject nName, Me
    End With
    Set NewUUCode = NewObj
    Set NewObj = Nothing
End Function

Private Function NewNCode(ByRef cProject As clsProject, ByVal nName As String, ByRef nText As String, Optional ByVal AddObjects As Boolean = True) As objNCode
    Dim NewObj As New objNCode
    With NewObj
        If AddObjects Then .InitObject nName, Me
    End With
    Set NewNCode = NewObj
    Set NewObj = Nothing
End Function

Private Function NewBWord(ByRef cProject As clsProject, ByVal nName As String, ByRef nText As String, Optional ByVal AddObjects As Boolean = True) As objBWord
    Dim NewObj As New objBWord
    With NewObj
        If AddObjects Then .InitObject nName, Me
    End With
    Set NewBWord = NewObj
    Set NewObj = Nothing
End Function


Private Function NewPoolID(ByRef cProject As clsProject, ByVal nName As String, ByRef nText As String, Optional ByVal AddObjects As Boolean = True) As objPoolID
    Dim NewObj As New objPoolID
    With NewObj
        If AddObjects Then .InitObject nName, Me
    End With
    Set NewPoolID = NewObj
    Set NewObj = Nothing
End Function


Private Function NewGUID(ByRef cProject As clsProject, ByVal nName As String, ByRef nText As String, Optional ByVal AddObjects As Boolean = True) As objGUID
    Dim NewObj As New objGUID
    With NewObj
        If AddObjects Then .InitObject nName, Me
    End With
    Set NewGUID = NewObj
    Set NewObj = Nothing
End Function

Private Function NewPlayer(ByRef cProject As clsProject, ByVal nName As String, ByRef nText As String, Optional ByVal AddObjects As Boolean = True) As objPlayer
    Dim NewObj As New objPlayer
    With NewObj
        .HasSoundNotify = ConvertEvent(cProject, "Object:SoundNotify", "(Self)", nName, nText)
        If AddObjects Then .InitObject nName, Me
    End With
    Set NewPlayer = NewObj
    Set NewObj = Nothing
End Function

Private Function CreateNewObject(ByRef cProject As clsProject, ByRef cItem As clsItem, ByRef nText As String, Optional ByVal AddObjects As Boolean = True)
    Dim NewObj As Object
    Select Case cItem.ItemClass

        Case "MaxIDE.Project"
            If AddObjects Then Set NewObj = New objProject
        Case "MaxIDE.Debug"
            If AddObjects Then Set NewObj = New objDebug
        Case "MaxIDE.Events"
            If AddObjects Then Set NewObj = New objEvents
            
        Case "NTAdvFTP61.URL"
            If AddObjects Then Set NewObj = New NTAdvFTP61.URL
        Case "NTPopup21.Window"
            If AddObjects Then Set NewObj = New NTPopup21.Window
        Case "NTShell22.Process"
            If AddObjects Then Set NewObj = New NTShell22.Process
        Case "NTShell22.Internet"
            If AddObjects Then Set NewObj = New NTShell22.Internet
        Case "NTCipher10.GUID"
            Set NewObj = NewGUID(cProject, cItem.ItemName, nText)
        Case "NTCipher10.BWord"
            Set NewObj = NewBWord(cProject, cItem.ItemName, nText)
        Case "NTCipher10.PoolID"
            Set NewObj = NewPoolID(cProject, cItem.ItemName, nText)
        Case "NTSound20.Player"
            Set NewObj = NewPlayer(cProject, cItem.ItemName, nText)
        Case "MaxIDE.Collect"
            Set NewObj = NewCollect(cProject, cItem.ItemName, nText)
        Case "NTAdvFTP61.Client"
            Set NewObj = NewClient(cProject, cItem.ItemName, nText)
        Case "NTSchedule20.Timer"
            Set NewObj = NewTimer(cProject, cItem.ItemName, nText)
        Case "NTSchedule20.Schedule"
            Set NewObj = NewSchedule(cProject, cItem.ItemName, nText)
        Case "NTCipher10.NCode"
            Set NewObj = NewNCode(cProject, cItem.ItemName, nText)
        Case "NTCipher10.UUCode"
            Set NewObj = NewUUCode(cProject, cItem.ItemName, nText)
        Case "NTAdvFTP61.Socket"
            Set NewObj = NewSocket(cProject, cItem.ItemName, nText)
    End Select

    Objects.Add NewObj, cItem.ItemName
    If AddObjects Then
        Select Case cItem.ItemClass

            Case "NTAdvFTP61.Client", "NTSchedule20.Timer", "NTSchedule20.Schedule", "NTAdvFTP61.Socket", "NTCipher10.NCode", "NTCipher10.GUID", "NTCipher10.PoolID", "NTCipher10.BWord", "NTCipher10.UUCode", "MaxIDE.Collect", "NTSound20.Player"
                ScriptControl1.AddObject cItem.ItemName, NewObj.GetObject, True
            Case Else
                ScriptControl1.AddObject cItem.ItemName, NewObj, True
        End Select
    End If
    Set NewObj = Nothing
    
End Function

Public Function ObjectClassExists(ByVal Class As String) As Boolean
    ObjectClassExists = False
    Dim m As Object
    For Each m In Objects
        If (TypeName(m) = Class) Then
            ObjectClassExists = True
            Exit Function
        End If
    Next
End Function

Public Function ObjectExists(ByVal name As String) As Boolean
    Dim str As Object
    On Error Resume Next
    Set str = Objects(name)
    If Err Then
        Err.Clear
        ObjectExists = False
    Else
        ObjectExists = True
    End If
    On Error GoTo 0
End Function

Public Function ModuleExists(ByVal name As String) As Boolean
    Dim str As String
    On Error Resume Next
    str = ScriptControl1.Modules.Item(name)
    If Err Then
        Err.Clear
        ModuleExists = False
    Else
        ModuleExists = True
    End If
    On Error GoTo 0
End Function

Private Function TestActiveXObject(ByVal name As String) As Boolean
    On Error GoTo catch
    Dim tmp As Object
    Set tmp = CreateObject(name)
    Set tmp = Nothing
    TestActiveXObject = True
    Exit Function
catch:
    Err.Clear
    TestActiveXObject = False
End Function
Private Function GenerateObjects(ByRef cProject As clsProject, Optional ByVal AddObjects As Boolean = True)

    Dim nText As String
    Dim nName As String
    Dim nLine As Long
    
    Dim cItem As clsItem
    For Each cItem In cProject.Items
        
        Select Case cItem.ItemClass
            Case "MaxIDE.Project"
                nMain = cItem.ItemSource

                ConvertEvent cProject, "Project:Exec", "()", cItem.ItemName, nMain, 5, nExec
                
                CreateNewObject cProject, cItem, nMain, AddObjects
                
            Case "MaxIDE.Generic"
                nText = cItem.ItemSource
                
                nLine = 4
                nName = ConvertEvent(cProject, "Object:Name", "()", cItem.ItemName, nText, nLine)
                If Not (nName = "") Then
                    nDims = nDims & IIf(cProject.Language = "VBScript", "'", "//") & "BREAKPOINT:" & nGUID & ":" & cItem.ItemName & ":" & nLine & vbCrLf
                    Select Case cProject.Language
                        Case "VBScript"
                            nDims = nDims & "Dim " & cItem.ItemName & " : Set " & cItem.ItemName & IIf(nName = "", " = Nothing", " = CreateObject(" & nName & ")") & vbCrLf
                        Case "JScript"
                            nDims = nDims & "var " & cItem.ItemName & IIf(nName = "", " = null;", " = new ActiveXObject(" & nName & "());") & vbCrLf
                    End Select
                End If
                
                ConvertEvent cProject, "Object:Init", "(Self)", cItem.ItemName, nText, 1, nInit
                ConvertEvent cProject, "Object:Term", "(Self)", cItem.ItemName, nText, 2, nTerm
                
                nBase = nBase & IIf(cProject.Language = "VBScript", "'", "//") & "BREAKPOINT:" & nGUID & ":" & cItem.ItemName & vbCrLf & nText & vbCrLf
                
            Case "MaxIDE.Module"
                nText = cItem.ItemSource

                ConvertEvent cProject, "Module:First", "(Name)", cItem.ItemName, nText, 3, nInit
                ConvertEvent cProject, "Module:Default", "(Name)", cItem.ItemName, nText, 3, nExec
                ConvertEvent cProject, "Module:Final", "(Name)", cItem.ItemName, nText, 4, nTerm

                nBase = nBase & IIf(cProject.Language = "VBScript", "'", "//") & "BREAKPOINT:" & nGUID & ":" & cItem.ItemName & vbCrLf & nText & vbCrLf
                
            Case Else
                nText = cItem.ItemSource

                ConvertEvent cProject, "Object:Init", "(Self)", cItem.ItemName, nText, 1, nInit
                ConvertEvent cProject, "Object:Term", "(Self)", cItem.ItemName, nText, 2, nTerm

                CreateNewObject cProject, cItem, nText, AddObjects

                nBase = nBase & IIf(cProject.Language = "VBScript", "'", "//") & "BREAKPOINT:" & nGUID & ":" & cItem.ItemName & vbCrLf & nText & vbCrLf

        End Select

    Next
   
    Dim tmp As Object
    
    If ObjectClassExists("objPoolID") Or ObjectClassExists("PoolID") Then
        Set tmp = New enuBit
        Objects.Add tmp, "CheckSums"
        ScriptControl1.AddObject "CheckSums", tmp, True
        Set tmp = Nothing
    End If
    
    
    If ObjectClassExists("objBWrod") Or ObjectClassExists("BWord") Then
        Set tmp = New enuBit
        Objects.Add tmp, "Bit"
        ScriptControl1.AddObject "Bit", tmp, True
        Set tmp = Nothing
    End If
    
    If ObjectClassExists("objURLTypes") Or ObjectClassExists("URLTypes") Then
        Set tmp = New enuURLTypes
        Objects.Add tmp, "URLTypes"
        ScriptControl1.AddObject "URLTypes", tmp, True
        Set tmp = Nothing
    End If
    
    If ObjectClassExists("objSchedule") Or ObjectClassExists("Schedule") Then
        Set tmp = New enuScheduleTypes
        Objects.Add tmp, "ScheduleTypes"
        ScriptControl1.AddObject "ScheduleTypes", tmp, True
        Set tmp = Nothing
        
        Set tmp = New enuIncrementTypes
        Objects.Add tmp, "IncrementTypes"
        ScriptControl1.AddObject "IncrementTypes", tmp, True
        Set tmp = Nothing
    End If
    
    
    If ObjectClassExists("objClient") Or ObjectClassExists("Client") Then
        Set tmp = New enuAllocateSides
        Objects.Add tmp, "AllocateSides"
        ScriptControl1.AddObject "AllocateSides", tmp, True
        Set tmp = Nothing

        Set tmp = New enuConnectedStates
        Objects.Add tmp, "ConnectedStates"
        ScriptControl1.AddObject "ConnectedStates", tmp, True
        Set tmp = Nothing
        
        Set tmp = New enuConnectionModes
        Objects.Add tmp, "ConnectionModes"
        ScriptControl1.AddObject "ConnectionModes", tmp, True
        Set tmp = Nothing
        
        Set tmp = New enuRateTypes
        Objects.Add tmp, "RateTypes"
        ScriptControl1.AddObject "RateTypes", tmp, True
        Set tmp = Nothing
        
        Set tmp = New enuListSystems
        Objects.Add tmp, "ListSystems"
        ScriptControl1.AddObject "ListSystems", tmp, True
        Set tmp = Nothing
        
        Set tmp = New enuMessageTypes
        Objects.Add tmp, "MessageTypes"
        ScriptControl1.AddObject "MessageTypes", tmp, True
        Set tmp = Nothing
        
        Set tmp = New enuErrorReturns
        Objects.Add tmp, "ErrorReturns"
        ScriptControl1.AddObject "ErrorReturns", tmp, True
        Set tmp = Nothing
        
        Set tmp = New enuTransferModes
        Objects.Add tmp, "TransferModes"
        ScriptControl1.AddObject "TransferModes", tmp, True
        Set tmp = Nothing
        
        Set tmp = New enuProgressTypes
        Objects.Add tmp, "ProgressTypes"
        ScriptControl1.AddObject "ProgressTypes", tmp, True
        Set tmp = Nothing

    End If
    
    If ObjectClassExists("objSocket") Or ObjectClassExists("Socket") Then
        Set tmp = New enuDirection
        Objects.Add tmp, "Direction"
        ScriptControl1.AddObject "Direction", tmp, True
        Set tmp = Nothing
        
        Set tmp = New enuStates
        Objects.Add tmp, "States"
        ScriptControl1.AddObject "States", tmp, True
        Set tmp = Nothing
    End If

End Function

Public Function ResetScript()
 
    If Objects.Count > 0 Then
    
        Dim delObj As Object
        Dim cnt As Long
        For cnt = 1 To Objects.Count
        
            Set delObj = Objects(cnt)
            Select Case TypeName(delObj)
                Case "objTimer", "objSocket", "objSchedule", "objUUCode", "objNCode", "objBWord", "objPoolID", "objGUID", "objCollect", "objClient", "objPlayer"
                    delObj.TermObject
                Case "objProject"
                    delObj.SetTimeout "", 0
                    delObj.SetInterval "", 0
            End Select
        Next
    
        Do Until Objects.Count = 0
            Objects.Remove 1
        Loop
        
    End If
    
    nFull = ""
    
    nDims = ""
    nBase = ""
    nInit = ""
    nExec = ""
    nTerm = ""
    nMain = ""

    HasExec = ""
    HasInit = ""
    HasTerm = ""
    
    On Error Resume Next
    ScriptControl1.Reset
    Project.Debugger.Reset
    If Not (Err.Number = 0) Then Err.Clear
    frmDebug.ResetDialect Project.Language, Nothing
    If Not (Err.Number = 0) Then Err.Clear
    On Error GoTo 0
    
End Function

Public Function MapProject(ByRef cProject As clsProject) As String
On Error GoTo catch:

    MapProject = GenerateProject(cProject, False)

Exit Function
catch:
    
    HandleNormalError

End Function

Public Function RunProject(ByRef cProject As clsProject)
On Error GoTo catch:

    ResetScript
    RaiseEvent Started

    ScriptControl1.Language = cProject.Language
    ScriptControl1.AllowUI = cProject.AllowUI
    ScriptControl1.timeout = 60000
    ScriptControl1.UseSafeSubset = False
    frmDebug.ResetDialect cProject.Language, ScriptControl1
    nFull = GenerateProject(cProject, True)
    ScriptControl1.AddCode nFull
    
    If Not HasInit = "" Then ScriptControl1.ExecuteStatement HasInit & IIf(cProject.Language = "JScript", "();", "")

    If Not HasExec = "" Then ScriptControl1.ExecuteStatement HasExec & IIf(cProject.Language = "JScript", "();", "")
    
Exit Function
catch:
    
    HandleNormalError
    
    CallUnload

End Function

Public Sub StopProject()
On Error GoTo catch:

    If (Not (HasTerm = "")) And ((ScriptControl1.Error.Number = 0) And (Err.Number = 0)) Then
        ScriptControl1.ExecuteStatement HasTerm & IIf(Project.Language = "JScript", "();", "")
    End If
    
catch:
    frmDebug.ResetDialect Project.Language, Nothing
    
    If Not ((ScriptControl1.Error.Number = 0) And (Err.Number = 0)) Then
        HandleNormalError
    End If

    ResetScript
    RaiseEvent Finished

End Sub

Public Function GenerateProject(ByRef cProject As clsProject, Optional ByVal AddObjects As Boolean = True)
    
    nGUID = Replace(GUID, "-", "")
    
    GenerateObjects cProject, AddObjects

    Dim ret As String
    
    ret = nDims & vbCrLf & nBase & vbCrLf
    
    ret = ret & IIf(cProject.Language = "VBScript", "'", "//") & "BREAKPOINT:" & nGUID & ":Project" & vbCrLf & nMain & vbCrLf

    HasInit = "a" & Replace(GUID, "-", "")
    HasExec = "a" & Replace(GUID, "-", "")
    HasTerm = "a" & Replace(GUID, "-", "")
    
    HasCall = "a" & Replace(GUID, "-", "")
    HasCall = "a" & Replace(GUID, "-", "")
    
    Select Case cProject.Language
        Case "VBScript"
            nInit = "Function " & HasInit & "()" & vbCrLf & nInit & "End Function" & vbCrLf
            nExec = "Function " & HasExec & "()" & vbCrLf & nExec & "End Function" & vbCrLf
            nTerm = "Function " & HasTerm & "()" & vbCrLf & nTerm & "End Function" & vbCrLf
        Case "JScript"
            nInit = "function " & HasInit & "() {" & vbCrLf & nInit & "}" & vbCrLf
            nExec = "function " & HasExec & "() {" & vbCrLf & nExec & "}" & vbCrLf
            nTerm = "function " & HasTerm & "() {" & vbCrLf & nTerm & "}" & vbCrLf
    End Select
    
    ret = ret & nInit & vbCrLf & nExec & vbCrLf & nTerm

    Dim nKeyword As String
    
    nKeyword = "iif"
    If FindLine(cProject, ret, nKeyword) > 0 Then
        Select Case cProject.Language
            Case "VBScript"
    
                ret = "Function Iif(Expression, TrueValue, FalseValue)" & vbCrLf & _
                    vbTab & "On Error Resume Next" & vbCrLf & _
                    vbTab & "If CBool(-Abs(Expression)) Then" & vbCrLf & _
                    vbTab & vbTab & "If Project.IsObject(TrueValue) Then" & vbCrLf & _
                    vbTab & vbTab & vbTab & "If Not (TrueValue Is Nothing) Then" & vbCrLf & _
                    vbTab & vbTab & vbTab & vbTab & "Set Iif = TrueValue" & vbCrLf & _
                    vbTab & vbTab & vbTab & "End If" & vbCrLf & _
                    vbTab & vbTab & "ElseIf Not Project.IsMissing(TrueValue) Then" & vbCrLf & _
                    vbTab & vbTab & vbTab & "Iif = TrueValue" & vbCrLf & _
                    vbTab & vbTab & "End If" & vbCrLf & _
                    vbTab & "ElseIf (Abs(Expression) = 0) Then" & vbCrLf & _
                    vbTab & vbTab & "If Project.IsObject(FalseValue) Then" & vbCrLf & _
                    vbTab & vbTab & vbTab & "If Not (FalseValue Is Nothing) Then" & vbCrLf & _
                    vbTab & vbTab & vbTab & vbTab & "Set Iif = FalseValue" & vbCrLf & _
                    vbTab & vbTab & vbTab & "End If" & vbCrLf & _
                    vbTab & vbTab & "ElseIf Not Project.IsMissing(FalseValue) Then" & vbCrLf & _
                    vbTab & vbTab & vbTab & "Iif = FalseValue" & vbCrLf & _
                    vbTab & vbTab & "End If" & vbCrLf & _
                    vbTab & "End If" & vbCrLf & _
                    vbTab & "On Error Goto 0" & vbCrLf & _
                    "End Function" & vbCrLf & ret
    
        End Select
    End If
    
    nKeyword = "nor"
    If FindLine(cProject, ret, nKeyword) > 0 Then
        Select Case cProject.Language
            Case "VBScript"
                ret = "Function Nor(Exp1, Exp2)" & vbCrLf & _
                    "Nor = (((((Not Exp1) And (Not Exp1)) = ((Not -Exp2) And (Not -Exp2)))) And ((((Not Exp1) And (Not Exp1)) = ((Not -Exp2) Or (Not -Exp2)))))" & vbCrLf & _
                    "End Function" & vbCrLf & ret
            Case "JScript"
                ret = "function Nor(Exp1, Exp2) { return Project.Nor(Exp1, Exp2) }" & vbCrLf & ret
        End Select
    End If
    
    nKeyword = "doevents"
    If FindLine(cProject, ret, nKeyword) > 0 Then
        Select Case cProject.Language
            Case "VBScript"
                ret = "Sub " & nKeyword & "()" & vbCrLf & vbTab & "Project.DoTasks" & vbCrLf & "End Sub" & vbCrLf & ret
            Case "JScript"
                ret = "function " & nKeyword & "() {" & vbCrLf & vbTab & "Project.DoTasks();" & vbCrLf & "}" & vbCrLf & ret
        End Select
    End If
    
    nKeyword = "alert"
    If FindLine(cProject, ret, nKeyword) > 0 Then
        Select Case cProject.Language
            Case "VBScript"
                ret = "Function Alert(Text)" & vbCrLf & vbTab & "Alert = Project.Alert(Text)" & vbCrLf & "End Function" & vbCrLf & ret
            Case "JScript"
                ret = "function " & nKeyword & "(Text, Buttons, Title) {" & vbCrLf & vbTab & "return Project.Alert(Text, Buttons, Title);" & vbCrLf & "}" & vbCrLf & ret
                ret = "function " & nKeyword & "(Text, Buttons) {" & vbCrLf & vbTab & "return Project.Alert(Text, Buttons);" & vbCrLf & "}" & vbCrLf & ret
                ret = "function " & nKeyword & "(Text) {" & vbCrLf & vbTab & "return Project.Alert(Text);" & vbCrLf & "}" & vbCrLf & ret

        End Select
    End If
    
    nKeyword = "msgbox"
    If FindLine(cProject, ret, nKeyword) > 0 Then
        Select Case cProject.Language
            Case "JScript"
                ret = "function " & nKeyword & "(Text, Buttons, Title) {" & vbCrLf & vbTab & "return Project.Alert(Text, Buttons, Title);" & vbCrLf & "}" & vbCrLf & ret
                ret = "function " & nKeyword & "(Text, Buttons) {" & vbCrLf & vbTab & "return Project.Alert(Text, Buttons);" & vbCrLf & "}" & vbCrLf & ret
                ret = "function " & nKeyword & "(Text) {" & vbCrLf & vbTab & "return Project.Alert(Text);" & vbCrLf & "}" & vbCrLf & ret

        End Select
    End If
                
    GenerateProject = ret
End Function

Private Function IronFaultEventCheck()
    'assuming userdefined data in making a pregenerated middle tier of routing event like driven function calls to the users custom
    'build requirements, we have worst case scenario in need of least possible interruption on behalf the internal workings of the IDE
    'for said forwarded events, thus assuming escape sequences are all met with no smoking embdeding of those to other envrionments
    'then the length is the other projected possible instance of sufforing IDE fault not the users script issue, nor their able debug
    'we would then assume, line too long and such type of out of memory events of those that allow streaming data such as a socket
    'need to be predisposed of the error (as data is not executable this is possible) before any other even triggered in assortment
    'of allowance to the memory throughput while the escape sequences may had to be placed on user data will also increase the size
    
    'overview of process:
    ' eval completion with out native vb error, nor script error, where as when occurs, the event stream is then broken down into
    ' parts that do not hinder the throughput least a single character is causing the error which is a scenario that shouldn't meet
    ' and if were, a passing it, likely is by time eased of maxed out cruch so that a single character can become through put
    ' at anytime the event has to cut the call into more then was was tested for, a single, to consecutive call of, there of the
    ' size from then on also is parted at same memory cut back, until another whole test is committed, remaining data keeps cutting
    ' equal to the last passed size of, to acheive this, a native vb and script "pad" wrapper, as well as object initited script
    ' "set" wrapper is duely escapeable to the parameters passed, that will be used by two script controls, in effect of one eval
    ' that in turn causes the call of two script controls uses of eval, to tandem the "set" and "pad" wrappers by way of internal
    ' generating with the native vb component.  This makes any executed script error, with in the script control contained, but
    ' generated of in native vb component, raises to the script control calling the eval, and the native vb errors to the eval level.
    ' in which, a attempted test on a single even call with a single data parameter, will be evaluation "line to long" a non-runtime
    ' error with in the scripting function to be eratic capture of the native vb highest call heirarchy, as any error is reduce mem
    ' then making shift into two calls, a 4th portion of the memory,a nd the remainder, logically, half on first receit of even raise
    ' and 1/3rd from then on in any subsequent error on the same event casting break up, or it passes with out error, and is called
    ' for the actual users code of scriptinog, this eliminates the possibility the IDE faults with no ability the users defined error
    ' which is allotted to the development of their own accord, and it essentially is the improbable, runtime compilation of script
    ' so with in reason to escaping the users data sequence with quotations, multiple embeded escape build up has to be created at
    ' the correct native vb or script in eval, andor "pad" and "set" style of casting script to the scriptcontrol such that they
    ' functionally happen at their embedeed experience and with out condone of the users escape sequence, not adding to sum of size
End Function

Public Sub RaiseCallBack(ByVal ObjectName As String, ByVal CallBackName As String, Optional ByVal Params As String = "")
On Error GoTo catch

    Select Case Project.Language
        Case "VBScript"
            ScriptControl1.ExecuteStatement CallBackName & " " & ObjectName & IIf(Params = "", "", ", " & Params)
        Case "JScript"
            ScriptControl1.ExecuteStatement CallBackName & "(" & ObjectName & IIf(Params = "", ");", ", " & Params & ");")
            
    End Select

Exit Sub
catch:
    HandleNormalError
    Err.Clear
End Sub

Public Function RaiseEvalExec(ByVal IsE As Boolean, ByVal Expression As String, ByVal StopOnError As Boolean) As Variant
On Error GoTo catch

    If IsE Then
        ScriptControl1.ExecuteStatement Expression
    Else
        RaiseEvalExec = ScriptControl1.Eval(Expression)
    End If

Exit Function
catch:

    Err.Source = "#:0:0:" & Expression

    HandleNormalError Not StopOnError

    Err.Clear
    
    If StopOnError Then
        ScriptControl1.ExecuteStatement "Project.Finish" & IIf(Project.Language = "JScript", "();", "")
    End If

End Function

Friend Function HandleNormalError(Optional ByVal IgnoreErrors As Boolean = False)

    If Not (ScriptControl1.Error.Number = 0) Then
    
        Dim cnt As Long
        Dim pos As Long
        Dim tmp As String
        Dim Line As Long
        Dim page As String
        
        cnt = ScriptControl1.Error.Line
        tmp = nFull
        
        Do While cnt > 1
            pos = pos + Len(RemoveNextArg(tmp, vbCrLf) + vbCrLf)
            cnt = cnt - 1
        Loop
        
        If (pos > 0) Then

            pos = InStrRev(nFull, IIf(Project.Language = "VBScript", "'", "//") & "BREAKPOINT:", pos)
            If (pos > 0) Then
                tmp = Mid(nFull, pos)
                tmp = RemoveNextArg(tmp, vbCrLf)
                RemoveNextArg tmp, ":"

                If nGUID = RemoveNextArg(tmp, ":") Then
   
                    If InStr(tmp, ":") > 0 Then
                        page = RemoveNextArg(tmp, ":")
                        Line = CLng(tmp)
                    Else
                        page = tmp
                        Line = CountWord(Left(nFull, pos), vbCrLf)
                        Line = (ScriptControl1.Error.Line - Line) - 1
                    End If

                    Project_Error IgnoreErrors, True, ScriptControl1.Error.Number, ScriptControl1.Error.Source, ScriptControl1.Error.Description, page, Line, ScriptControl1.Error.Column, ScriptControl1.Error.Text
                                       
                End If
            End If
            
        End If
        
        ScriptControl1.Error.Clear
        
        If Not (Err.Number = 0) Then
            CustomError IgnoreErrors
        End If
        
    ElseIf Not (Err.Number = 0) Then
    
        CustomError IgnoreErrors
            
    End If

End Function

Private Function CustomError(ByVal IgnoreErrors As Boolean)
    Dim lin As Long
    Dim col As Long
    Dim page As String
    Dim Text As String
    
    If Left(Err.Source, 1) = "#" Then
        Text = Mid(Err.Source, 2)
        
        page = RemoveNextArg(Text, ":")
        lin = CLng(RemoveNextArg(Text, ":"))
        col = CLng(RemoveNextArg(Text, ":"))
        
        Project_Error IgnoreErrors, False, Err.Number, "MaxIDE VBScript compilation error", Err.Description, page, lin, col, Text
    Else
        Project_Error IgnoreErrors, False, Err.Number, "MaxIDE VBScript compilation error", Err.Description, "", 0, 0, ""
    End If
    
    Err.Clear
    
End Function

Private Function Project_Error(ByVal IgnoreErrors As Boolean, ByVal Runtime As Boolean, ByVal Number As Long, ByVal Source As String, ByVal Description As String, ByVal PageName As String, ByVal LineNum As Long, ByVal Column As Long, ByVal LineText As String)

    DebugWinPrint vbCrLf & IIf(Runtime, "         Script: ", "        Compile: ") & IIf(Not Source = "", Source & vbCrLf, vbCrLf) & _
                    IIf(Not Description = "", "          Error: " & Description & IIf(Not Number = 0, " (Num: " & Number & ")", "") & vbCrLf, "") & _
                    IIf(Not LineText = "", "           Code: """ & Trim(LineText) & """" & IIf(Column > 1, " (Column: " & Column & ")", "") & vbCrLf, "")

    If Not IgnoreErrors Then
        If (Not frmMainIDE.Visible) And Project.AllowUI Then frmMainIDE.ShowForm
        If Not PageName = "" Then
            frmProjectExplorer.ShowScriptPage PageName, LineNum, ScriptControl1.Error.Column
        End If
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    If Not timMe Is Nothing Then
        timMe.Enabled = False
        Set timMe = Nothing
    End If
End Sub

Private Sub timMe_OnTicking()
On Error Resume Next

    timMe.Enabled = False
    
    ResetScript
    RaiseEvent Finished

Err.Clear
End Sub

Private Sub CallUnload()
    If timMe Is Nothing Then
        Set timMe = New NTSchedule20.Timer
    End If
    timMe.Interval = 10
    timMe.Enabled = True
End Sub




