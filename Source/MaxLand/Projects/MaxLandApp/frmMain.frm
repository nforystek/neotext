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
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      UseSafeSubset   =   -1  'True
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

Private Sub Form_Click()
    TrapMouse = True
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    StopGame = True
End Sub


Public Sub Startup()
    With ScriptControl1
        
        Static AlreadyRan As Boolean
        If AlreadyRan Then
            .Reset
        Else
            AlreadyRan = True
        End If

        If PathExists(AppPath & "Script.log", True) Then Kill AppPath & "Script.log"
        
        .Language = "VBScript"
        'only the global add the code members of
        .AddObject "Include", modParse.Include, True
        'the rest are builds of and not code based
        .AddObject "All", modParse.All, False
        .AddObject "Beacons", modParse.Beacons, False
        .AddObject "Bindings", modParse.Bindings, False
        .AddObject "Boards", modParse.Boards, False
        .AddObject "Cameras", modParse.Cameras, False
        .AddObject "Elements", modParse.Elements, False
        .AddObject "Lights", modParse.Lights, False
        .AddObject "Player", modParse.Player, False
        .AddObject "Portals", modParse.Portals, False
        .AddObject "Screens", modParse.Screens, False
        .AddObject "Sounds", modParse.Sounds, False
        .AddObject "Space", modParse.Space, False
        .AddObject "Tracks", modParse.Tracks, False


    End With
End Sub


Public Sub AddCode(ByVal Code As String, Optional ByVal source As String = "AddCode", Optional ByVal LineNumber As Long = 0)

    ScriptControl1.AddCode Code
'    Do Until Code = ""
'        DebugPrint "Addcode: " & RemoveNextArg(Code, vbCrLf)
'    Loop
End Sub


Public Function Evaluate(ByVal Expression As Variant, Optional ByVal source As String = "Evaluate", Optional ByVal LineNumber As Long = 0) As Variant
    Evaluate = ScriptControl1.Eval(Expression)
   ' DebugPrint "Eval: " & Expression & " = " & Evaluate
End Function

Public Sub ExecuteStatement(ByVal Statement As String, Optional ByVal source As String = "ExecuteStatement", Optional ByVal LineNumber As Long = 0)
    ScriptControl1.ExecuteStatement Statement
   ' DebugPrint "Execute: " & Statement
End Sub


Public Function Run(ByRef ProcedureName As Variant, Optional ByVal source As String = "Run", Optional ByVal LineNumber As Long = 0) As Variant

    If ScriptControl1.Procedures.Count > 0 Then
        Dim i As Long
        For i = 1 To ScriptControl1.Procedures.Count
            If LCase(ScriptControl1.Procedures(i).Name) = LCase(ProcedureName) Then
                ScriptControl1.Run ProcedureName
               ' DebugPrint "Run: " & ProcedureName
                Exit For
            End If
        Next
    End If

End Function

Private Sub DebugPrint(ByVal txt As String)
    Dim fn As Integer
    fn = FreeFile
    Open AppPath & "Script.log" For Append As #fn
        Print #fn, txt
    Close #fn
    'Debug.Print txt
End Sub

