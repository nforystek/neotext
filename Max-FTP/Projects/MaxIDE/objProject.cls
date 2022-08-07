VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "objProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private Const Base36$ = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"

Private TimeoutEval As String
Private Timeouteval_Break As Boolean

Private IntervalEval As String
Private IntervalEval_Break As Boolean

Private WithEvents sTimeout As NTSchedule20.Timer
Attribute sTimeout.VB_VarHelpID = -1
Private WithEvents sInterval As NTSchedule20.Timer
Attribute sInterval.VB_VarHelpID = -1

Private Function Decim&(Sessa2$)
    Dim Posiz%, CifraT&, ValC&
    CifraT& = 0
    For Posiz% = 1 To Len(Sessa2$)
        ValC& = (InStr(Base36$, Mid$(Sessa2$, Len(Sessa2$) - Posiz% + 1, 1)) - 1) * 36 ^ (Posiz% - 1)
        CifraT& = CifraT& + ValC&
    Next Posiz%
    Decim& = CifraT&
End Function

Private Function Sessa2$(Decim&)
    Dim DeScr&, CifraT$, Cifra%
    DeScr& = Decim&
    CifraT$ = ""
    Do
        Cifra% = DeScr& Mod 36
        DeScr& = DeScr& \ 36
        CifraT$ = Mid$(Base36$, Cifra% + 1, 1) + CifraT$
    Loop Until DeScr& = 0
    Sessa2$ = CifraT$
End Function

Public Function GetMachineName() As String
    GetMachineName = dbSettings.GetMachineName
End Function

Public Function Switches() As String
    Switches = Command
End Function

Public Function MessageID() As String
    MessageID = Sessa2$(Int(1 + ((LongBound - 1) * Rnd))) & "." & Sessa2$(Int(1 + ((LongBound - 1) * Rnd))) & "@" & Sessa2$(Int(1 + ((LongBound - 1) * Rnd))) & Sessa2$(Int(1 + ((LongBound - 1) * Rnd)))
End Function

Public Function MathRandom() As Variant
    Randomize
    MathRandom = Rnd
End Function

Public Function MathFloor(X) As Variant
    Dim str As String
    str = CStr(X)
    If InStr(X, ".") > 0 Then
        str = Left(str, InStr(str, ".") - 1)
    End If
    MathFloor = CDbl(str)
End Function

Public Function Format(ByVal InDate As Variant, ByVal Mask As Variant) As Variant
    Format = VBA.Format(InDate, Mask)
End Function
Public Function Nor(exp1, exp2)
    Nor = modCommon.Nor(exp1, exp2)
End Function

Public Function IsObject(Optional Variable)
    
    IsObject = modInclude.CheckVar(Variable)
        
End Function

Public Function IsMissing(Optional Argument)
    
    IsMissing = modInclude.CheckParam(Argument)
        
End Function

Public Sub SetTimeout(ByVal Statement As String, ByVal Milliseconds As Long, Optional ByVal StopOnError As Boolean = True)
    
    sTimeout.Enabled = False
    TimeoutEval = Statement
    Timeouteval_Break = StopOnError
    sTimeout.Interval = Milliseconds
    sTimeout.Enabled = (Milliseconds > 0) And Not (Statement = "")
        
End Sub

Public Sub SetInterval(ByVal Statement As String, ByVal Milliseconds As Long, Optional ByVal StopOnError As Boolean = True)

    sInterval.Enabled = False
    IntervalEval = Statement
    IntervalEval_Break = StopOnError
    sInterval.Interval = Milliseconds
    sInterval.Enabled = (Milliseconds > 0) And Not (Statement = "")
    
End Sub

Public Function Evaluate(ByVal Expression As String, Optional ByVal StopOnError As Boolean = True) As Variant
    If Not (Project.Compiler Is Nothing) Then
        Evaluate = Project.Compiler.Parent.RaiseEvalExec(False, Expression, StopOnError)
    End If
End Function

Public Sub Execute(ByVal Statement As String, Optional ByVal StopOnError As Boolean = True)
    If Not (Project.Compiler Is Nothing) Then
        Project.Compiler.Parent.RaiseEvalExec True, Statement, StopOnError
    End If
End Sub

Public Function Finish()
    sTimeout.Enabled = False
    sInterval.Enabled = False
    
    Project.StopProject
End Function

Public Sub DoTasks()
   ' On Error Resume Next
    modCommon.DoTasks
   ' If Not (Err.Number = 0) Then Err.Clear
   ' On Error GoTo 0
End Sub

Public Function GUID() As String
    GUID = modGuid.GUID
End Function

Public Function Alert(ByVal Text As Variant, Optional Buttons As Variant = vbOKOnly, Optional Title As Variant = "") As Variant
    If Project.AllowUI Then
        On Error Resume Next
        Alert = MsgBox(CStr(Text), CLng(Buttons), CStr(Title))
        If Not (Err.Number = 0) Then Err.Clear
        On Error GoTo 0
    End If
End Function

Public Function path(Optional ByVal VarPath As String = "") As String
    If VarPath = "" Then
        path = AppPath
    Else
        Dim tmp As String
        tmp = MapFolderVariables(VarPath)
        If modCommon.PathExists(AppPath & tmp, False) Then
            path = AppPath & tmp
        ElseIf modCommon.PathExists(tmp, False) Then
            path = tmp
        ElseIf modCommon.PathExists(VarPath, False) Then
            path = VarPath
        Else
            path = AppPath
        End If
    End If
End Function

Public Function VarPath(ByVal inPath As String) As String
    VarPath = MapFolderVariables(inPath)
End Function

Public Function ReadTextFile(ByVal FileName As String) As Variant
    ReadTextFile = ReadFile(FileName)
End Function
Public Sub WriteTextFile(ByVal FileName As String, ByVal Text As Variant)
    Dim str As String
    str = CStr(Text)
    WriteFile FileName, str
End Sub

Public Function PathExists(ByVal FileName As String, Optional ByVal IsFile As Boolean = False) As Boolean
    PathExists = modCommon.PathExists(FileName, IsFile)
End Function

Public Function GetFilePath(ByVal URL As String) As String
    GetFilePath = modCommon.GetFilePath(URL)
End Function

Public Function GetFileTitle(ByVal URL As String) As String
    GetFileTitle = modCommon.GetFileTitle(URL)
End Function

Public Function GetFileName(ByVal URL As String) As String
    GetFileName = modCommon.GetFileName(URL)
End Function

Public Function GetFileExt(ByVal URL As String, Optional ByVal Lower As Boolean = True, Optional ByVal RemoveDot As Boolean = False) As String
    GetFileExt = modCommon.GetFileExt(URL, Lower, RemoveDot)
End Function

Public Function GetFileSize(ByVal URL As String) As Long
    GetFileSize = modCommon.GetFileSize(URL)
End Function

Public Function GetFileDate(ByVal URL As String) As Date
    GetFileDate = modCommon.GetFileDate(URL)
End Function

Public Function FormatFileSize(ByVal FileSize As String) As String
    FormatFileSize = modCommon.FormatFileSize(FileSize)
End Function

Public Function TrimStrip(ByVal TheStr As String, ByVal TheChar As String) As String
    TrimStrip = modCommon.TrimStrip(TheStr, TheChar)
End Function

Public Function RemoveNextArg(ByRef TheParams As Variant, ByVal TheSeperator As String) As String
    RemoveNextArg = modCommon.RemoveNextArg(TheParams, TheSeperator)
End Function

Public Function NextArg(ByVal TheParams As String, ByVal TheSeperator As String) As String
    NextArg = modCommon.NextArg(TheParams, TheSeperator)
End Function

Public Function RemoveArg(ByVal TheParams As String, ByVal TheSeperator As String) As String
    RemoveArg = modCommon.RemoveArg(TheParams, TheSeperator)
End Function

Public Function IsAlphaNumeric(ByVal Text As String) As Boolean
    IsAlphaNumeric = modCommon.IsAlphaNumeric(Text)
End Function

Private Sub Class_Initialize()
    Set sTimeout = New NTSchedule20.Timer
    Set sInterval = New NTSchedule20.Timer
End Sub

Private Sub Class_Terminate()
    sTimeout.Enabled = False
    sInterval.Enabled = False
    Set sTimeout = Nothing
    Set sInterval = Nothing
End Sub

Private Sub sInterval_OnTicking()
    If Not (Project.Compiler Is Nothing) Then
        Project.Compiler.Parent.RaiseEvalExec True, IntervalEval, IntervalEval_Break
    End If
End Sub

Private Sub sTimeout_OnTicking()
    sTimeout.Enabled = False
    If Not (Project.Compiler Is Nothing) Then
        Project.Compiler.Parent.RaiseEvalExec True, TimeoutEval, Timeouteval_Break
    End If
End Sub

Attribute 