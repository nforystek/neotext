Attribute VB_Name = "modMain"
#Const modMain = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module

Public Const AppWebsite = "http://www.neotext.org"
Public Const AppName = "Creata-Tree"

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Ini As clsIniFile
Public DefaultIni As String

Public TreeFileExt As String
Public ItemFileExt As String

Public EngineFolder As String
    Public MediaFolder As String
    Public TemplateFolder As String
        Public BlankTreeFile As String
        Public BlankItemFile As String
        
Public MyTreeFolder As String
Public ExportFolder As String
Public ExampleFolder As String

Public Const ErrorImageKey = "error"
Public Const BlankImageText = "(Blank)"
Public Const BlankImageKey = "blank"

Public Sub Main()
    '%LICENSE%
tryit: On Error GoTo catch

    
    InitMainGlobals
    If Not ExecuteFunction(Command) Then
    
        frmMain.Show
                        
        If Not Command = "" Then
            If PathExists(Command, True) Then
                frmMain.nForm.OpenFile Command
            End If
        End If
            
    End If
    
GoTo final
catch: On Error Resume Next

    'If Err Then MsgBox Err.Description, vbExclamation, App.EXEName

final: On Error Resume Next

On Error GoTo -1
End Sub

Public Function ExecuteFunction(ByVal CommandLine As String) As Boolean
On Error GoTo errcmd

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
                Case "hidefolder"
                    SetAttr EngineFolder, VbFileAttribute.vbHidden
                    Exit Do
                Case "showfolder"
                    SetAttr EngineFolder, VbFileAttribute.vbNormal
                    Exit Do
                Case "setupreset"
                    WriteFile AppPath & "CreataTree.ini", DefaultIni
                    Exit Do
                Case Else
                    HasCmd = False
            End Select
        Loop
    End If
    
    ExecuteFunction = HasCmd
    
errcmd:
    If Err Then Err.Clear
    On Error GoTo 0
End Function
Public Sub ColumnSortClick(ByVal ListView1 As ListView, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ListView1.SortKey <> ColumnHeader.SubItemIndex Then
        ListView1.SortKey = ColumnHeader.SubItemIndex
    
    Else
        If ListView1.SortOrder = 0 Then
            ListView1.SortOrder = 1
        Else
            ListView1.SortOrder = 0
        End If
    
    End If
End Sub

Public Function IsOnList(ByRef tList As Control, ByVal Item As String) As Integer

    Dim cnt As Integer
    Dim found As Integer
    cnt = 0
    found = -1
    Do Until cnt = tList.ListCount Or found > -1
        If LCase(Trim(tList.List(cnt))) = LCase(Trim(Item)) Then found = cnt
        cnt = cnt + 1
    Loop
    IsOnList = found

End Function


Public Function NodeExists(ByRef mCol As Variant, ByVal mKey As Variant) As Boolean
    On Error Resume Next
    NodeExists = True
    mKey = SafeKey(mKey)
    Select Case TypeName(mCol)
        Case "ListImages", "IImages", "Nodes", "INodes", "ListItems", "IListItems"
            Dim Test As Integer
            Test = mCol(mKey).Index
        Case "Collection"
            Dim test2 As Variant
            test2 = mCol(mKey)
    End Select
    If Err Then
        Err.Clear
        NodeExists = False
    End If
    On Error GoTo 0
End Function

Public Function FormalWord(ByVal mKey As String) As String
    FormalWord = UCase(Left(mKey, 1)) & LCase(Mid(mKey, 2))
End Function
Public Function SafeKey(ByVal mKey As Variant) As Variant
    If IsNumeric(mKey) Then
        SafeKey = CLng(mKey)
    Else
        Dim Ret As String
        Ret = Trim(mKey)
        
        Ret = Replace(Ret, "/", CStr(Asc("/")))
        Ret = Replace(Ret, "\", CStr(Asc("\")))
        Ret = Replace(Ret, "|", CStr(Asc("|")))
        Ret = Replace(Ret, "[", CStr(Asc("[")))
        Ret = Replace(Ret, "]", CStr(Asc("]")))
        Ret = Replace(Ret, " ", CStr(Asc(" ")))
        Ret = Replace(Ret, "&", CStr(Asc("&")))
        Ret = Replace(Ret, "%", CStr(Asc("%")))
        Ret = Replace(Ret, "$", CStr(Asc("$")))
        Ret = Replace(Ret, ",", CStr(Asc(",")))
        Ret = Replace(Ret, ".", CStr(Asc(".")))
        Ret = Replace(Ret, "-", CStr(Asc("-")))
        Ret = Replace(Ret, "_", CStr(Asc("_")))
        Ret = Replace(Ret, "+", CStr(Asc("+")))
        Ret = Replace(Ret, "=", CStr(Asc("=")))
        Ret = Replace(Ret, "!", CStr(Asc("!")))
        Ret = Replace(Ret, "@", CStr(Asc("@")))
        Ret = Replace(Ret, "#", CStr(Asc("#")))
        Ret = Replace(Ret, "^", CStr(Asc("^")))
        Ret = Replace(Ret, "*", CStr(Asc("*")))
        Ret = Replace(Ret, "(", CStr(Asc(")")))
        Ret = Replace(Ret, "`", CStr(Asc("`")))
        Ret = Replace(Ret, "~", CStr(Asc("~")))
        Ret = Replace(Ret, "<", CStr(Asc("<")))
        Ret = Replace(Ret, ">", CStr(Asc(">")))
        Ret = Replace(Ret, "'", CStr(Asc("'")))
        Ret = Replace(Ret, """", CStr(Asc("""")))
        
        SafeKey = LCase(Ret)
    End If
End Function
