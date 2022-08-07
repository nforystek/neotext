#Const [True] = -1
#Const [False] = 0
Attribute VB_Name = "modMainIDE"
#Const modMainIDE = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module

Public Const AppName = "Max-FTP IDE"

Public dbSettings As clsSettings

Public Const ScriptFile = "VBScript"

Public Const ObjectFile = "Object"
Public Const ModuleFile = "Module"
Public Const ProjectFile = "Project"

Public MaxEvents As clsEventLog
Public Project As clsProject

Public Sub MaxDBError(ByVal Number As Long, ByVal Description As String, ByRef Retry As Boolean)

    If (Number = -2147467259) Or (Number = 3709) Then
    Debug.Print Err.Description
            MsgBox "User permissions insufficient to run Max-FTP, your user account must be part of a group with" & vbCrLf & _
                "permissions to access and modify the Max-FTP database located under the installation directory." & vbCrLf & _
                "Contact your administrator to set up proper user group privlidges for Max-FTP or your account." & vbCrLf & vbCrLf & _
                "(Unable to write to sub folders or database: " & AppPath & ")", vbInformation + vbOKOnly, AppName
        
        End
    Else
        
        frmDBError.ShowError Description
                
        Do Until frmDBError.Visible = False
            DoTasks
        Loop
        Select Case frmDBError.IsOk
            Case 0

                Unload frmMainIDE
                                
            Case 2
                Retry = True
        End Select
    End If
End Sub

Public Function GetTextObj() As Object
    If Not frmMainIDE.ActiveForm Is Nothing Then
        If TypeName(frmMainIDE.ActiveForm) = "frmScriptPage" Then
            Set GetTextObj = frmMainIDE.ActiveForm.CodeEdit1
        Else
            Set GetTextObj = Nothing
        End If
    Else
        Set GetTextObj = Nothing
    End If
End Function

Public Sub Main()
    '%LICENSE%

tryit: On Error GoTo catch

    Set dbSettings = New clsSettings

    If SetupUser(dbSettings.GetUserLoginName) Then
        
      
        Set MaxEvents = New clsEventLog
        Set Project = New clsProject
        
        Load frmMainIDE

        If Not ExecuteFunction(Command) And Project.AllowUI Then
            frmMainIDE.ShowForm
        ElseIf Not frmMainIDE.Visible Or Not Project.AllowUI Then
            Unload frmMainIDE
        End If


    Else
        MsgBox "Error creating or loading User Profile for '" + dbSettings.GetUserLoginName + "'.  You may need to reinstall " & AppName & ".", vbCritical, AppName
    End If

GoTo final
catch: On Error GoTo 0

    If Err Then MsgBox Err.Description, vbExclamation, App.EXEName

final: On Error Resume Next

On Error GoTo -1
End Sub

Public Function ExecuteFunction(ByVal CommandLine As String)
    Dim hasCmd As Boolean
    If Trim(CommandLine) <> "" Then
        Dim InParams As String
        Dim InCommand As String
        Dim inProject As String
        
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
            
            hasCmd = True
            Select Case Replace(InCommand, " ", "")
                Case "noui"
                    Project.AllowUI = False
                Case "run"
                    If InStr(InParams, """") > 0 Then
                        inProject = RemoveQuotedArg(InParams, """", """")
                    Else
                        inProject = InParams
                    End If
                    
                    If Not PathExists(inProject, True) Then
                        DebugWinPrint vbCrLf & "File not found: " & inProject & vbCrLf
                        hasCmd = False
                        Exit Do
                    Else
                        frmMainIDE.ShowForm
                        frmMainIDE.OpenProject inProject
                        Project.RunProject
                        Do While Project.IsRunning
                            DoTasks
                        Loop
                        
                        hasCmd = False
                        Exit Do
                    End If
                Case "exec"
                    If InStr(InParams, """") > 0 Then
                        inProject = RemoveQuotedArg(InParams, """", """")
                    Else
                        inProject = InParams
                    End If
                    
                    If Not PathExists(inProject, True) Then
                        DebugWinPrint vbCrLf & "File not found: " & inProject & vbCrLf
                        hasCmd = False
                        Exit Do
                    Else
                        frmMainIDE.OpenProject inProject
                        Project.RunProject
                        Do While Project.IsRunning
                            DoTasks
                        Loop
    
                        Exit Do
                    End If
                Case "open"
                    If InStr(InParams, """") > 0 Then
                        inProject = RemoveQuotedArg(InParams, """", """")
                    Else
                        inProject = InParams
                    End If
                    
                    If Not PathExists(inProject, True) Then
                        DebugWinPrint vbCrLf & "File not found: " & inProject & vbCrLf
                        hasCmd = False
                        Exit Do
                    Else
                        frmMainIDE.ShowForm
                        frmMainIDE.OpenProject inProject
                        
                        hasCmd = False
                        Exit Do
                    End If
                
                Case Else
                    hasCmd = False
            End Select

        Loop
    End If
    
    ExecuteFunction = hasCmd
    
End Function

Public Function ValidNameSpace(ByVal Text As String) As Boolean
    If IsAlphaNumeric(Text) Then
        ValidNameSpace = True
    Else
        MsgBox "You must enter an alpha numeric name with no spaces.", vbInformation, AppName
        ValidNameSpace = False
    End If
End Function

Public Sub LoadScriptGraphics()
    
    With frmMainIDE
        .ToolBarOut.ListImages.Clear
        .ToolBarOut.ImageHeight = GetSkinDimension("toolbarbutton_height")
        .ToolBarOut.ImageWidth = GetSkinDimension("toolbarbutton_width")
        
        .ToolBarOver.ListImages.Clear
        .ToolBarOver.ImageHeight = GetSkinDimension("toolbarbutton_height")
        .ToolBarOver.ImageWidth = GetSkinDimension("toolbarbutton_width")
        
        LoadButton .ToolBarOver, "script_newproject", 1
        LoadButton .ToolBarOver, "script_openproject", 2
        LoadButton .ToolBarOver, "script_saveproject", 3
        
        LoadButton .ToolBarOver, "script_addfile", 4
        LoadButton .ToolBarOver, "script_removefile", 5
    
        LoadButton .ToolBarOver, "script_undo", 6
        LoadButton .ToolBarOver, "script_redo", 7
    
        LoadButton .ToolBarOver, "script_cut", 8
        LoadButton .ToolBarOver, "script_copy", 9
        LoadButton .ToolBarOver, "script_paste", 10
        LoadButton .ToolBarOver, "script_find", 11
        
        LoadButton .ToolBarOver, "script_stop", 12
        LoadButton .ToolBarOver, "script_run", 13
    
        LoadButton .ToolBarOut, "script_newprojectout", 1
        LoadButton .ToolBarOut, "script_openprojectout", 2
        LoadButton .ToolBarOut, "script_saveprojectout", 3
        
        LoadButton .ToolBarOut, "script_addfileout", 4
        LoadButton .ToolBarOut, "script_removefileout", 5
    
        LoadButton .ToolBarOut, "script_undoout", 6
        LoadButton .ToolBarOut, "script_redoout", 7
    
        LoadButton .ToolBarOut, "script_cutout", 8
        LoadButton .ToolBarOut, "script_copyout", 9
        LoadButton .ToolBarOut, "script_pasteout", 10
        LoadButton .ToolBarOut, "script_findout", 11
        
        LoadButton .ToolBarOut, "script_stopout", 12
        LoadButton .ToolBarOut, "script_runout", 13
    
        .ToolBarOut.MaskColor = GetSkinColor("script_toolout_transparentcolor")
        .ToolBarOut.UseMaskColor = True
        .ToolBarOver.MaskColor = GetSkinColor("script_toolover_transparentcolor")
        .ToolBarOver.UseMaskColor = True
    End With
End Sub

Public Function DebugWinPrint(ByVal inText As String)
    Dim obj As New objDebug
    obj.PrintLine inText
    Set obj = Nothing
End Function



Attribute 