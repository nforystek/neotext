Attribute VB_Name = "Module1"
Option Explicit

'#################################################################
'#################################################################
'############# EXAMPLE: USING OBJECTS LIKE VARIABLES #############
'#################################################################
'#################################################################


Public Sub Main()
    Dim test1 As New Plane
    Dim test2 As New Plane
    Dim test3 As New Point
    
    
    test1.X = 8 'normal set of property through accessor
    test2 = "0,7,0,0" 'set the whole object by a string
    
    Debug.Print test1 & " = " & test2;
    
    'test if they are equal
    If (test1 = test2) Then
        Debug.Print " True"
    Else
        Debug.Print " False"
    End If
    
    
    test1 = test2 'more like clone, and not like 'Set'
    
    Debug.Print test1 & " = " & test2;
    
    'test if they are equal
    If (test1 = test2) Then
        Debug.Print " True"
    Else
        Debug.Print " False"
    End If
    
    
    test1.X = 6 'set some data to make it unique valued
    Set test2 = test1 'use the 'Set' statement this time

    Debug.Print test1 & " = " & test2;
    
    'test if they are equal
    If (test1 = test2) Then
        Debug.Print " True"
    Else
        Debug.Print " False"
    End If


    test1.X = 0 'clear test1 back to default all zero's
    
    Debug.Print test2 & " = " & test3;
    
    'test if they are equal
    If (test2 = test3) Then
        Debug.Print " True"
    Else
        Debug.Print " False"
    End If

End Sub
