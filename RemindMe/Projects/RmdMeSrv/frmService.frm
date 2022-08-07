VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmService 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RemindMe Service"
   ClientHeight    =   720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   ControlBox      =   0   'False
   Icon            =   "frmService.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   900
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
End
Attribute VB_Name = "frmService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public Sub RunProcedure(ByVal ProcedureText As String)
On Error GoTo catch
    ScriptControl1.Error.Clear
    
    ScriptControl1.ExecuteStatement ProcedureText

catch:

    If ScriptControl1.Error.Number <> 0 Or Err.Number <> 0 Then
        If (ProcessRunning(RemindMeFileName) > 0) Then
            dbSettings.Message RemindMeFileName, Replace("error:" & _
                        ProcedureText & ":proc" & vbCrLf & _
                        ScriptControl1.Error.Number & ":numb" & vbCrLf & _
                        ScriptControl1.Error.Line & ":line" & vbCrLf & _
                        ScriptControl1.Error.Column & ":colu" & vbCrLf & _
                        ScriptControl1.Error.Source & ":sour" & vbCrLf & _
                        ScriptControl1.Error.Description & ":desc" & vbCrLf, "'", "''")
        Else
            MsgBox "An error occured while trying to run a scripted procedure." & vbCrLf & vbCrLf & _
                    "Procedure: " & ProcedureText & vbCrLf & _
                    "Number: " & ScriptControl1.Error.Number & vbCrLf & _
                    IIf(ScriptControl1.Error.Source <> "", "Source: " & ScriptControl1.Error.Source & vbCrLf, "") & _
                    IIf(ScriptControl1.Error.Line > 0, "Line: " & ScriptControl1.Error.Line & vbCrLf, "") & _
                    IIf(ScriptControl1.Error.Column > 0, "Column: " & ScriptControl1.Error.Column & vbCrLf, "") & _
                    "Description: " & ScriptControl1.Error.Description & vbCrLf, vbCritical, AppName
        End If
    End If
    
If Err Then Err.Clear
On Error GoTo 0
End Sub

Public Function UpdateScript()
    EndScript
    ScriptControl1.AllowUI = True
    ScriptControl1.UseSafeSubset = False
    ScriptControl1.Language = dbSettings.GetSetting("Language")
    ScriptControl1.AddCode dbSettings.GetSetting(dbSettings.GetSetting("Language") & "Text")
End Function

Public Sub EndScript()
    On Error Resume Next

    ScriptControl1.Reset
    If Not (Err.Number = 0) Then Err.Clear

    On Error GoTo 0
End Sub

