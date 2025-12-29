VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MaxLand"
   ClientHeight    =   570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1590
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmMain.frx":0D4A
   MousePointer    =   1  'Arrow
   ScaleHeight     =   570
   ScaleWidth      =   1590
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSScriptControlCtl.ScriptControl ScriptControl2 
      Left            =   840
      Tag             =   "ScriptControl2"
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   120
      Tag             =   "ScriptControl1"
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

#If VBA = -1 Then

'--------------------------------------------------
'**** VBA Integration code, begin insert
Public m_apcInt As APCIntegration
'**** VBA Integration code, end insert
'--------------------------------------------------


'--------------------------------------------------
'**** VBA Integration code, begin insert
Public Sub ShowVBE()
    ' Show the VBE now
    m_apcInt.ShowVBE
End Sub
'**** VBA Integration code, end insert
'--------------------------------------------------



'--------------------------------------------------
'**** VBA Integration code, begin insert
Public Sub ShowMacroDialog()
    ' The macros dialog will only be viewable if you have a project loaded
    m_apcInt.ShowMacroDialog
End Sub
'**** VBA Integration code, end insert
'--------------------------------------------------


Private Sub Form_Load()

'--------------------------------------------------
'**** VBA Integration code, begin insert
    Dim appObj As Instance

    Set m_apcInt = New APCIntegration
    ' If you already have an existing application object in your
    '  original source, replace "new Instance" here with a
    '  reference to it.
    Set appObj = New Instance
    m_apcInt.Initialize appObj, Me.hwnd
'**** VBA Integration code, end insert
'--------------------------------------------------

End Sub

#End If

'--------------------------------------------------
'**** VBA Integration code, begin insert
Private Sub Form_QueryUnLoad(cancel As Integer, unloadmode As Integer)
#If VBA = -1 Then
    m_apcInt.QueryUnload cancel, unloadmode
#End If
    If Not cancel Then StopGame = True
    If unloadmode = 0 Then cancel = True
End Sub '**** VBA Integration code, end insert
'--------------------------------------------------


Public Property Get ScriptControl() As ScriptControl
    Set ScriptControl = ScriptControl1
End Property

Private Sub Form_Click()
    MouseTrapped = True
End Sub



Public Sub Reset()

    ScriptControl.Reset

    If PathExists(AppPath & "Script.log", True) Then Kill AppPath & "Script.log"
End Sub
Public Sub Startup()
    
    With frmMain.ScriptControl
        .Language = "VBScript"
        'add members of non collections so far
        .AddObject "Include", modParse.Include, True
        .AddObject "Bindings", modParse.Bindings, True
        .AddObject "Camera", modParse.Camera, True
        .AddObject "Player", modParse.Player, True
        
        'unsure benefits so all false for now
        .AddObject "All", modParse.All, False
        .AddObject "Beacons", modParse.Beacons, False
        .AddObject "Boards", modParse.Boards, False
        .AddObject "Cameras", modParse.Cameras, False
        .AddObject "Elements", modParse.Elements, False
        .AddObject "Lights", modParse.Lights, False
        .AddObject "Portals", modParse.Portals, False
        .AddObject "Screens", modParse.Screens, False
        .AddObject "Sounds", modParse.Sounds, False
        .AddObject "Spaces", modParse.Spaces, False
        .AddObject "Tracks", modParse.Tracks, False
        .AddObject "Players", modParse.Players, False

    End With
End Sub


Public Sub AddCode(ByVal Code As String, Optional ByVal source As String = "AddCode", Optional ByVal LineNumber As Long = 0)
'    On Error GoTo tryforgiegner
    

    frmMain.ScriptControl.AddCode Code
    If ScriptDebug Then
        Do Until Code = ""
            DebugPrint "Addcode: " & RemoveNextArg(Code, vbCrLf)
        Loop
    End If
    
'    Exit Sub
'tryforgiegner:
'    If Err.Number <> 0 Then
'        Dim Num As Long
'        Dim des As String
'        Dim src As String
'        Num = Err.Number
'        src = Err.source
'        des = Err.Description
'        Err.Clear
'        On Error GoTo 0
'        '"An error occured while setting up an object." & vbCrLf
'        Err.Raise Num, src, "Line: " & (LineNumber + 1) & " Error: " & des
'    End If
End Sub


Public Function Evaluate(ByVal Expression As Variant, Optional ByVal source As String = "Evaluate", Optional ByVal LineNumber As Long = 0) As Variant
'    On Error GoTo tryforgiegner
    

    Evaluate = frmMain.ScriptControl.Eval(Expression)
    If ScriptDebug Then DebugPrint "Eval: " & Expression & " = " & Evaluate

'    Exit Sub
'tryforgiegner:
'    If Err.Number <> 0 Then
'        Dim Num As Long
'        Dim des As String
'        Dim src As String
'        Num = Err.Number
'        src = Err.source
'        des = Err.Description
'        Err.Clear
'        On Error GoTo 0
'        '"An error occured while setting up an object." & vbCrLf
'        Err.Raise Num, src, "Line: " & (LineNumber + 1) & " Error: " & des
'    End If
End Function

Public Sub ExecuteStatement(ByVal Statement As String, Optional ByVal source As String = "ExecuteStatement", Optional ByVal LineNumber As Long = 0)
'    On Error GoTo tryforgiegner
    

    frmMain.ScriptControl.ExecuteStatement Statement
    If ScriptDebug Then DebugPrint "Execute: " & Statement
    
    
'tryforgiegner:
'    Dim Num As Long
'    Dim des As String
'    Dim src As String
'    If Err.Number <> 0 Then
'        Num = Err.Number
'        src = Err.source
'        des = Err.Description
'        Err.Clear
'        If ScriptControl1.Error.Number <> 0 Then
'            ScriptControl1.Error.Clear
'        End If
'    ElseIf ScriptControl1.Error.Number <> 0 Then
'        Num = ScriptControl1.Error.Number
'        src = ScriptControl1.Error.source
'        des = ScriptControl1.Error.Description
'        ScriptControl1.Error.Clear
'    End If
'
'
'
'    On Error GoTo 0
'
'    If Num = 91 Then
'        scriptcontrol1.
'        '"An error occured while setting up an object." & vbCrLf
'        Err.Raise Num, src, "Line: " & (LineNumber + 1) & " Error: " & des
'    End If
'
'    If Not Num = 91 And Not Err = 0 Then
'        Err.Raise Num, src, "Line: " & (LineNumber + 1) & " Error: " & des
'    End If
End Sub


Public Function Run(ByRef ProcedureName As Variant, Optional ByVal source As String = "Run", Optional ByVal LineNumber As Long = 0) As Variant
'    On Error GoTo tryforgiegner
    


    If frmMain.ScriptControl.Procedures.Count > 0 Then
        Dim i As Long
        For i = 1 To frmMain.ScriptControl.Procedures.Count
            If LCase(frmMain.ScriptControl.Procedures(i).Name) = LCase(ProcedureName) Then
                frmMain.ScriptControl.Run ProcedureName
                If ScriptDebug Then DebugPrint "Run: " & ProcedureName
                Exit For
            End If
        Next
    End If
    

'tryforgiegner:
'        Dim Num As Long
'        Dim des As String
'        Dim src As String
'        Num = Err.Number
'        src = Err.source
'        des = Err.Description
'        Err.Clear
'        On Error GoTo 0
'
'    If Err <> 0 Then
'
'        '"An error occured while setting up an object." & vbCrLf
'        Err.Raise Num, src, "Line: " & (LineNumber + 1) & " Error: " & des
'    End If
End Function

Private Sub DebugPrint(ByVal txt As String)
    Dim fn As Integer
    fn = FreeFile
    Open AppPath & "Script.log" For Append As #fn
        Print #fn, txt
    Close #fn
    Debug.Print txt
End Sub

Private Sub Form_Resize()
    MouseOverCanvas 0, 0
End Sub
