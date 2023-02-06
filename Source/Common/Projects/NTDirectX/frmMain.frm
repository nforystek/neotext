VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "KadPatch"
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6090
   ClipControls    =   0   'False
   FillColor       =   &H8000000F&
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1305
      Left            =   4515
      ScaleHeight     =   1305
      ScaleWidth      =   1260
      TabIndex        =   2
      Top             =   540
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1530
      Left            =   1230
      ScaleHeight     =   1530
      ScaleWidth      =   1755
      TabIndex        =   1
      Top             =   2955
      Visible         =   0   'False
      Width           =   1755
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   4020
      Top             =   2700
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   0   'False
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1740
      Left            =   0
      ScaleHeight     =   1740
      ScaleWidth      =   3720
      TabIndex        =   0
      Top             =   0
      Width           =   3720
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

Public LastX As Long
Public LastY As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 112 Then ShowSetup Picture1.Parent
End Sub

Private Sub Form_Load()
    Startup
End Sub

Public Sub Startup()
    With ScriptControl1
       
       
        .Language = "VBScript"
        'only the global add the code members of
        .AddObject "Include", modParse.Include, True
        'the rest are builds of and not code based
        .AddObject "All", modParse.All
        '.AddObject "Camera", modParse.Camera
        .AddObject "Motions", modParse.Motions
        .AddObject "Brilliants", modParse.Brilliants
        .AddObject "Molecules", modParse.Molecules
        .AddObject "Billboards", modParse.Billboards
        .AddObject "Bindings", modParse.Bindings
        .AddObject "Planets", modParse.Planets

    End With
End Sub
Public Function Serialize(Optional ByVal Deserialize As Variant) As String
    On Error GoTo errcatch:
    On Local Error GoTo errcatch:
    With ScriptControl1
        If .Procedures.Count > 0 Then
            Dim cnt As Long
            For cnt = 1 To .Procedures.Count
                If ((Not IsMissing(Deserialize)) And (LCase(.Procedures.Item(cnt).Name) = "deserialize")) Then
                    .Run "Deserialize"
                    If Deserialize <> "" Then
                        .ExecuteStatement Deserialize
                    End If
                ElseIf (IsMissing(Deserialize) And (LCase(.Procedures.Item(cnt).Name) = "serialize")) Then
                    Serialize = .Eval("Serialize")
                End If
            Next
        End If
    Exit Function
errcatch:
        If Not ConsoleVisible Then
            ConsoleToggle
        End If
        Process "echo An error " & Err.Number & " occurd in " & Err.Source & _
                "\n" & "Description: " & Err.Description
        If .Error.Number <> 0 Then .Error.Clear
        If Err.Number <> 0 Then Err.Clear
    End With
End Function
Public Sub AddCode(ByVal Code As String)
    On Error GoTo errcatch:
    On Local Error GoTo errcatch:
    With ScriptControl1
        .AddCode Code
    Exit Sub
errcatch:
        If Not ConsoleVisible Then
            ConsoleToggle
        End If
        Process "echo An error " & Err.Number & " occurd in " & Err.Source & _
                "\n" & "Description: " & Err.Description
        If .Error.Number <> 0 Then .Error.Clear
        If Err.Number <> 0 Then Err.Clear
    End With
End Sub
Public Function Evaluate(ByVal Expression As Variant) As Variant
    On Error GoTo errcatch:
    On Local Error GoTo errcatch:
    With ScriptControl1
        Evaluate = .Eval(Expression)
    Exit Function
errcatch:
        If Not ConsoleVisible Then
            ConsoleToggle
        End If
        Process "echo An error " & Err.Number & " occurd in " & Err.Source & _
                "\n" & "Description: " & Err.Description
        If .Error.Number <> 0 Then .Error.Clear
        If Err.Number <> 0 Then Err.Clear
    End With
End Function
Public Sub ExecuteStatement(ByVal Statement As String)
    On Error GoTo errcatch:
    On Local Error GoTo errcatch:
    With ScriptControl1
        .ExecuteStatement Statement
    Exit Sub
errcatch:
        If Not ConsoleVisible Then
            ConsoleToggle
        End If
        Process "echo An error " & Err.Number & " occurd in " & Err.Source & _
                "\n" & "Description: " & Err.Description
        If .Error.Number <> 0 Then .Error.Clear
        If Err.Number <> 0 Then Err.Clear
    End With
End Sub
Public Sub RunEvent(ByRef EventText As String)
    On Error GoTo errcatch:
    On Local Error GoTo errcatch:
    'any event not yet run is given a guid procedure name and added
    'then subsequent calls are addressing only the guid, changing it
    'to code will create it new, increasing ghosted memory procedures
    With ScriptControl1
        If EventText <> "" Then
            If Not IsGuid(EventText) Then
                Dim proc As String
                proc = modGuid.GUID
                .AddCode "Sub b" & Replace(proc, "-", "") & _
                    "()" & vbCrLf & Replace(EventText, "Debug.Print", "DebugPrint", , , vbTextCompare) & vbCrLf & "End Sub" & vbCrLf
                EventText = proc
            End If
            frmMain.Run "b" & Replace(EventText, "-", "")
        End If
    Exit Sub
errcatch:
        If Not ConsoleVisible Then
            ConsoleToggle
        End If
        Process "echo An error " & Err.Number & " occurd in " & Err.Source & _
                "\n" & "Description: " & Err.Description
        If .Error.Number <> 0 Then .Error.Clear
        If Err.Number <> 0 Then Err.Clear
    End With
End Sub

Public Function Run(ByRef ProcedureName As Variant) As Variant
    On Error GoTo errcatch:
    On Local Error GoTo errcatch:
    With ScriptControl1
        .Run ProcedureName
        If .Error.Number <> 0 Then
            Err.Raise .Error.Number, .Error.Source, .Error.Description & vbCrLf & _
                "At line " & .Error.Line & " column " & .Error.Column & " of sniplet; " & vbCrLf & .Error.Text
        End If
    Exit Function
errcatch:
        If Not ConsoleVisible Then
            ConsoleToggle
        End If
        Process "echo An error " & Err.Number & " occurd in " & Err.Source & _
                "\n" & "Description: " & Err.Description
        If .Error.Number <> 0 Then .Error.Clear
        If Err.Number <> 0 Then Err.Clear
    End With
End Function

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 112 Then ShowSetup Picture1.Parent
End Sub

Private Sub ScriptControl1_Error()
    With ScriptControl1
        If .Error.Number <> 0 Then
            If Not ConsoleVisible Then
                ConsoleToggle
            End If
            Debug.Print "echo An error " & Err.Number & " occurd in " & Err.Source & _
                    vbCrLf & "Description: " & .Error.Description & vbCrLf & _
                    "At line " & .Error.Line & " code sniplet; " & vbCrLf & .Error.Text
            Process "echo An error " & Err.Number & " occurd in " & Err.Source & _
                    vbCrLf & "Description: " & .Error.Description & vbCrLf & _
                    "At line " & .Error.Line & " code sniplet; " & vbCrLf & .Error.Text
            If .Error.Number <> 0 Then .Error.Clear
        End If
    End With
End Sub

Private Sub Form_Resize()
    Picture1.Top = 0
    Picture1.Left = 0
    Picture1.Width = Me.Width
    Picture1.Height = Me.Height
End Sub

