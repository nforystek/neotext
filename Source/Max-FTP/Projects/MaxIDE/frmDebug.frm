VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmDebug 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4770
   ControlBox      =   0   'False
   Icon            =   "frmDebug.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSScriptControlCtl.ScriptControl ScriptControl3 
      Left            =   3870
      Top             =   630
      _ExtentX        =   1005
      _ExtentY        =   1005
      UseSafeSubset   =   -1  'True
   End
   Begin MaxIDE.ctlDragger ctlDragger1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      Top             =   0
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   688
      Dockable        =   0   'False
      Resizable       =   0   'False
      Movable         =   0   'False
      Docked          =   0   'False
      Caption         =   "Debug Window"
      RepositionForm  =   0   'False
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   1710
      Left            =   270
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1050
      Width           =   2985
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private WithEvents ScriptControl2 As ScriptControl
Attribute ScriptControl2.VB_VarHelpID = -1

Public Property Get ScriptControl1() As ScriptControl
    If (ScriptControl2 Is Nothing) Then
        Set ScriptControl1 = ScriptControl3
    Else
        Set ScriptControl1 = ScriptControl2
    End If
End Property
Public Property Set ScriptControl1(newVal As ScriptControl)
    Set ScriptControl2 = newVal
End Property

Public Sub ResetDialect(ByVal Language As String, ByRef sControl As ScriptControl)
    If Not Project.IsRunning Then
        ScriptControl1.Language = Language
        ScriptControl1.AllowUI = Project.AllowUI
        ScriptControl1.timeout = 60000
        ScriptControl1.UseSafeSubset = True
    End If
    If (sControl Is Nothing) And Project.IsRunning Then
        Set ScriptControl2 = ScriptControl3
    Else
        Set ScriptControl2 = sControl
    End If
End Sub

Private Sub ctlDragger1_Resize(Left As Long, Top As Long, Width As Long, ByVal Height As Long)
    On Error Resume Next
    Text1.Move Left, Top, Width, Height
    Err.Clear
End Sub

Private Sub Form_Load()
    With ctlDragger1
        .Docked = dbSettings.GetScriptingSetting("dDocked")
    
        .DockedWidth = dbSettings.GetScriptingSetting("dDockedWidth")
        .DockedHeight = dbSettings.GetScriptingSetting("dDockedHeight")
    
        .FloatingTop = dbSettings.GetScriptingSetting("dFloatTop")
        .FloatingLeft = dbSettings.GetScriptingSetting("dFloatLeft")
        .FloatingWidth = dbSettings.GetScriptingSetting("dFloatWidth")
        .FloatingHeight = dbSettings.GetScriptingSetting("dFloatHeight")

        .SetupDockedForm frmMainIDE, Me, 2

    End With
    
End Sub

Public Sub PrintLine(ByVal Text As String)
    If Len(Text1.Text & Text) + 2 >= 30000 Then
        Text1.Text = Right(Text1.Text, Len(Text1.Text) - (Len(Text) + 2)) & Text & vbCrLf
    Else
        Text1.Text = Text1.Text & Text & vbCrLf
    End If
    Text1.SelStart = Len(Text1.Text)
End Sub

Public Sub PrintText(ByVal Text As String)
    If Len(Text1.Text & Text) >= 30000 Then
        Text1.Text = Right(Text1.Text, Len(Text1.Text) - Len(Text)) & Text
    Else
        Text1.Text = Text1.Text & Text
    End If
    Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ScriptControl2 = Nothing
    On Error Resume Next
    ScriptControl3.Reset
    If Err.Number Then Err.Clear
    On Error GoTo 0
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo catch

    If KeyCode = 13 Then
        Dim tmp As String
        tmp = Left(Text1.Text, Text1.SelStart - 2)
        tmp = StrReverse(tmp)
        tmp = RemoveNextArg(tmp, vbLf + vbCr)
        tmp = StrReverse(tmp)
        If Not Trim(tmp) = "" Then
            
            If Left(tmp, 1) = "?" And (Len(tmp) > 1) Then
    
                    PrintText Me.ScriptControl1.Eval(Mid(tmp, 2))
                    'PrintText "Automation Exception:  Action not allowed."

            ElseIf Left(tmp, 1) = "!" And (Len(tmp) > 1) Then

                    Me.ScriptControl1.ExecuteStatement Mid(tmp, 2)
                    'PrintText "Automation Exception:  Action not allowed."

            End If
            
        End If
    End If

Exit Sub
catch:
    If Project.IsRunning Then
        Project.HandleError
        frmDebug.PrintText "            Use ?<expression> or !<statement> to" & vbCrLf & _
                           "            eval or execute in the debug window." & vbCrLf
                            
    Else
        MsgBox "Error: " & Me.ScriptControl1.Error.Number & " " & Me.ScriptControl1.Error.Description & vbCrLf & vbCrLf & _
        "Use ?<expression> or !<statement> to eval or execute in the debug window."
    End If
    Err.Clear
End Sub


