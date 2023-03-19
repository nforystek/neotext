Attribute VB_Name = "Module1"
Option Explicit


Public Sub Main()
    On Local Error Resume Next
    On Error Resume Next
    
    Dim ss As NTSoSweet.ScriptHost
    
    Set ss = New NTSoSweet.ScriptHost
    

    ss.Interpreter "C:\Development\Neotext\Common\Projects\NTSoSweet\Test\Sweets\Hypertext Markup Language.ssw"

    
    Debug.Print ss.Generate
    
    
   Set ss = Nothing

End Sub
