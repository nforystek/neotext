
Attribute VB_Name = "modMain"
#Const modMain = -1
Option Explicit
'TOP DOWN
Option Compare Text
Option Private Module
Public Sub Main()

    If App.PrevInstance Then
        MsgBox "Please finish and close any previously open ToDoster occurance to continue to another.", vbOKOnly + vbCritical, App.Title
    Else
    
        Select Case LCase(Trim(Command))
        
            Case "data"
                frmData.Show
                
            Case "edit"
                frmProds.Show
            
            Case Else
                frmMain.Show
            
        End Select
        
    End If

End Sub