Attribute VB_Name = "Module1"
Option Explicit

Public Sub Main()
    Dim astr As String
    Dim alng As Long
    Dim aarr() As Byte
    
    Dim key As String
    Dim val As String
    
    
    astr = "Test"
    alng = 777
    ReDim aarr(0 To 2) As Byte
    aarr(0) = 7
    aarr(1) = 7
    aarr(2) = 7

    key = "Software\Testerly"
    val = "Tested"
    
    Dim reg As New Registry
    
    reg.CreateKey HKEY_CURRENT_USER, key
    
    
    reg.SetValue HKEY_CURRENT_USER, key, "TestStr", astr
    reg.SetValue HKEY_CURRENT_USER, key, "TestLng", alng
    reg.SetValue HKEY_CURRENT_USER, key, "TestArr", aarr
    
'    Debug.Print reg.GetValue(HKEY_CURRENT_USER, key, "")
    
    Erase aarr
    ReDim aarr(0 To 2) As Byte
    
    Debug.Print reg.GetValue(HKEY_CURRENT_USER, key, "TestStr")
    Debug.Print reg.GetValue(HKEY_CURRENT_USER, key, "TestLng")
    aarr = reg.GetValue(HKEY_CURRENT_USER, key, "TestArr")

    
    Debug.Print aarr(0) & " " & aarr(1) & " " & aarr(2)
    
    
    Set reg = Nothing


End Sub
