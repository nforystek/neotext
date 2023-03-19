Attribute VB_Name = "Module1"
#Const Module1 = -1
Option Explicit
'TOP DOWN

Public Declare Sub CoFreeUnusedLibraries Lib "ole32" ()


Public Sub Main()
    
    Dim tmp As Object
    On Error Resume Next
    Dim lstText As String
    Dim nxtText As String
    
    Do While True
    '  CoFreeUnusedLibraries
         Set tmp = CreateObject("Project2.Class1") 'NameBasedObjectFactory.CreateObjectPrivate("Class1")
        If Err Then
            Err.Clear
            nxtText = "Error: Can't create ActiveX object."
        Else
            nxtText = tmp.test
        End If
        Set tmp = Nothing
        If nxtText <> lstText Then
            lstText = nxtText
            Debug.Print nxtText
        End If
        DoEvents
        Sleep 1
        
    Loop
    


End Sub

