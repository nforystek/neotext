VERSION 5.00
Begin VB.Form frmBatch 
   BorderStyle     =   0  'None
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   1305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "frmBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private xHandle As Long
Private xSource As String

Public Sub StartRun()
On Error GoTo errorit

    Dim batchText As String

    Dim verbose As Byte
    Dim LineNum As Long
    Dim inLine As String
    Dim tmp As String
        
    Do Until fMain.txtMain.FileName = ""
        DoEvents
        Sleep 1
    Loop

    batchText = fMain.txtMain.Text
    verbose = 2
    
    Do Until ((fMain.BatchSignal = 1) Or (batchText = ""))
        inLine = Trim(RemoveNextArg(batchText, vbCrLf))
        LineNum = LineNum + 1
        fMain.GotoLine LineNum
        If LineNum > 1 Then fMain.txtMain.LineDefines(LineNum - 1) = ColoringIndexes.Dream3Index
        fMain.txtMain.LineDefines(LineNum) = ColoringIndexes.Dream4Index

        DoEvents
        Sleep 1
        If Not Left(LCase(inLine), 3) = "rem" And inLine <> "" Then
        
            If Left(LCase(inLine), 5) = "@echo" Then
                Select Case Trim(LCase(Mid(inLine, 7)))
                    Case "off"
                        verbose = verbose - 1
                    Case "on"
                        verbose = verbose + 1
                End Select
            Else
                
                If verbose >= 3 Then
                    'display cmd
                    
                End If
                If (Left(LCase(inLine), 4) = "echo") And (verbose > 1) Then
                    'Debug.Print Trim(Mid(inLine, 6))
                Else
                    tmp = RemoveArg(GetFilePath(fMain.TextFileName), ":")
                    inLine = Left(GetFilePath(fMain.TextFileName), 2) & vbCrLf & _
                            "cd " & IIf(tmp = "", "\", tmp) & vbCrLf & inLine
                    xSource = App.Path & "\" & App.EXEName & ".bat"
                            
                    xHandle = FreeFile
                    
                    Open xSource For Binary As #xHandle
                    Put #xHandle, , inLine
                    Close #xHandle
                    RunProcess xSource, , vbHide, True 'fMain.BatchSignal
                    If PathExists(xSource, True) Then Kill xSource

                End If
            End If
        End If
        
    Loop
    fMain.BatchSignal = 2
errorit:

    If fMain Is Nothing Then
        TerminateProcess GetCurrentProcess, 0
    Else
        Unload Me
    End If
End Sub


Private Sub Form_Load()
    StartRun
End Sub
