Attribute VB_Name = "Module1"
Option Explicit

Private events As New events.Driver

Public Sub Main()
    events.AddCallBack AddressOf Testing1
    events.AddCallBack AddressOf Testing2
    Form1.Show 'keep app from closing
    
End Sub
Private Sub Testing1()
    Debug.Print "Test1"
    
End Sub
Private Sub Testing2()
    Debug.Print "Test2"
    
End Sub

