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
   Begin MSScriptControlCtl.ScriptControl ScriptControl2 
      Left            =   4875
      Top             =   2715
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   0   'False
      UseSafeSubset   =   -1  'True
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
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
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
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
    If KeyCode = 112 Then ShowSetup = True
End Sub

Public Sub Startup()
    With ScriptControl1
        
        Static AlreadyRan As Boolean
        If AlreadyRan Then .Reset
        AlreadyRan = True
        
        .Language = "VBScript"
        'only the global add the code members of
        .AddObject "Include", modParse.Include, True
        'the rest are builds of and not code based
       
        .AddObject "All", modParse.All, True
        .AddObject "Camera", modParse.Camera, True
        .AddObject "Motions", modParse.Motions, True
        .AddObject "Brilliants", modParse.Brilliants, True
        .AddObject "Molecules", modParse.Molecules, True
'        .AddObject "Billboards", modParse.Billboards, True
        .AddObject "Bindings", modParse.Bindings, True
        .AddObject "Planets", modParse.Planets, True
        .AddObject "OnEvents", modParse.OnEvents, True
    
    End With
End Sub
Public Function Deserialize() As String
    With ScriptControl1
        If .Procedures.Count > 0 Then
            Dim cnt As Long
            For cnt = 1 To .Procedures.Count
                If (LCase(.Procedures.Item(cnt).Name) = "deserialize") Then
                    .Run "Deserialize"
                    If .Error.number <> 0 Then Err.Raise .Error.number, "Deserialize", .Error.description & vbCrLf & "Line Number: " & .Error.line & " of sniplet; " & vbCrLf & .Error.Text
                End If
            Next
        End If
    End With
End Function

Public Function Serialize() As String
    With ScriptControl1
'        If .Procedures.Count > 0 Then
'            Dim cnt As Long
'            For cnt = 1 To .Procedures.Count
'                If (LCase(.Procedures.Item(cnt).Name) = "serialize") Then
                    Serialize = .Eval("Serialize")
                    If .Error.number <> 0 Then Err.Raise .Error.number, "Serialize", .Error.description & vbCrLf & "Line Number: " & .Error.line & " of sniplet; " & vbCrLf & .Error.Text
'                End If
'            Next
'        End If
    End With
End Function

Public Sub AddCode(ByVal Code As String, Optional ByVal source As String = "AddCode", Optional ByVal LineNumber As Long = 0)
    With ScriptControl1
        .AddCode Code
        If .Error.number <> 0 Then Err.Raise .Error.number, source, .Error.description & vbCrLf & "Line Number: " & LineNumber & " of sniplet; " & vbCrLf & Code
    End With
End Sub

Public Function Evaluate(ByVal Expression As Variant, Optional ByVal source As String = "Evaluate", Optional ByVal LineNumber As Long = 0) As Variant
    With ScriptControl1
        Evaluate = .Eval(Expression)
        If .Error.number <> 0 Then Err.Raise .Error.number, source, .Error.description & vbCrLf & "Line Number: " & LineNumber & " of sniplet; " & vbCrLf & Expression
    End With
End Function

Public Sub ExecuteStatement(ByVal Statement As String, Optional ByVal source As String = "ExecuteStatement", Optional ByVal LineNumber As Long = 0)
    With ScriptControl1
        .ExecuteStatement Statement
        If .Error.number <> 0 Then Err.Raise .Error.number, source, .Error.description & vbCrLf & "Line Number: " & LineNumber & " of sniplet; " & vbCrLf & Statement
    End With
End Sub

Public Sub RunEvent(ByRef EventText As String, Optional ByVal source As String = "RunEvent", Optional ByVal LineNumber As Long = "0")
    'any event not yet run is given a guid procedure name and added
    'then subsequent calls are addressing only the guid, changing it
    'to code will create it new, increasing ghosted memory procedures
    With ScriptControl1
'        If EventText <> "" Then
'            If Not IsGuid(EventText) Then
'                Dim proc As String
'                proc = modGuid.GUID
'                .AddCode "Sub b" & Replace(proc, "-", "") & _
'                    "()" & vbCrLf & Replace(EventText, "Debug.Print", "DebugPrint", , , vbTextCompare) & vbCrLf & "End Sub" & vbCrLf
'                EventText = proc
'                If .Error.number <> 0 Then Err.Raise .Error.number, source, .Error.description & vbCrLf & "Line Number: " & LineNumber & " of sniplet; " & vbCrLf & EventText
'            Else
'                Stop
'            End If
'            frmMain.Run "b" & Replace(EventText, "-", ""), source, LineNumber
'        End If


        frmMain.ExecuteStatement Replace(EventText, "Debug.Print", "DebugPrint", , , vbTextCompare) & vbCrLf, source, LineNumber

    End With
End Sub

Public Function Run(ByRef ProcedureName As Variant, Optional ByVal source As String = "Run", Optional ByVal LineNumber As Long = 0) As Variant
    With ScriptControl1
        .Run ProcedureName
        If .Error.number <> 0 Then Err.Raise .Error.number, source, .Error.description & vbCrLf & "Line Number: " & LineNumber & " of sniplet; " & vbCrLf & ProcedureName
    End With
End Function


Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 112 Then ShowSetup = True
End Sub

Private Sub Form_Resize()
    Picture1.Top = 0
    Picture1.Left = 0
    Picture1.Width = Me.Width
    Picture1.Height = Me.Height
End Sub

