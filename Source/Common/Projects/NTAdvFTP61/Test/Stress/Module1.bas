Attribute VB_Name = "Module1"
Option Explicit

Public Sub Main()

    Dim ftp As NTAdvFTP61.Socket
    
    ReadyInput
    
    Do
        Set ftp = New NTAdvFTP61.Socket
        ftp.ssl = True
        
        DoLoop
        InputLoop
        
        Set ftp = Nothing
        
    Loop Until EndOfInput
End Sub
