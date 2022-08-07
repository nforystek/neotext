Attribute VB_Name = "Module1"
Option Explicit

Public Sub Main()

    Form1.Show
    Exit Sub
    
    
    Dim s As New Stream
    

    s.concat Convert("testing hi")
    
    
    Dim oatom As Long
    Dim watom As Long
    

    s.offsets(oatom) = 5
    s.Widths(oatom) = 5

    s.offsets(watom) = 5
    s.Widths(watom) = 5
    
    Debug.Print oatom & " " & s.offsets(oatom) & " " & s.Widths(oatom)

    Debug.Print watom & " " & s.offsets(watom) & " " & s.Widths(watom)
    
    s.prepend Convert("hi")
    
    Debug.Print oatom & " " & s.offsets(oatom) & " " & s.Widths(oatom)

    Debug.Print watom & " " & s.offsets(watom) & " " & s.Widths(watom)

    s.Length = s.Length - 2
    
    Debug.Print oatom & " " & s.offsets(oatom) & " " & s.Widths(oatom)

    Debug.Print watom & " " & s.offsets(watom) & " " & s.Widths(watom)

    
    
    Set s = Nothing
    
End Sub
