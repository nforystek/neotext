Attribute VB_Name = "modRestart"
#Const modRestart = -1
Option Explicit
'TOP DOWN

Public Sub Main()

    Select Case Command
        Case "stop"
            StopIdentd
            End
        Case "start"
            StartIdentd
            End
        Case Else
            StopIdentd
            StartIdentd
        
    End Select
End Sub

Public Sub StartIdentd()
    NetStart "Identd", "Ident.exe"
End Sub

Public Sub StopIdentd()
    NetStop "Identd", "Ident.exe"
End Sub

