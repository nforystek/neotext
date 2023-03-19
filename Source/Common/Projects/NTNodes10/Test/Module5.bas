Attribute VB_Name = "Module5"
#Const Module5 = -1
Option Explicit
'TOP DOWN

Type test
    var1 As Long
    var2 As Long
End Type

Public Sub Main()


    Dim col As New NTNodes10.Collection
    
    Dim tmp As test
    
    tmp.var1 = 38
    tmp.var2 = 894
    col.Add tmp
    
    
    End
    
    
    Dim cls As New Class1
    col.Add cls, "ref1"
    col.Add cls, "ref2"
    Set cls = Nothing
    col("ref1").vals = 3453
    Debug.Print col("ref2").vals


    
    
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

'    col.Add fso.getfolder("C:\")
'    col.Add fso.getfolder("C:\WINDOWS")
'    col.Add fso.getfolder("C:\WINDOWS\TEMp")
    
    Debug.Print "Add ""this"", ""K4"""
    Debug.Print "Add ""that"", ""K5"""

    col.Add "this", "4_5"
    col.Add "that", "K5"

    Debug.Print col.Exists("4_5") & " " & col.Exists("K23")
    
    

    
    Dim t As test
    t.var1 = 23478
    col.Add t, "key"
    
    
    
    Dim tmp
    For Each tmp In col
        Debug.Print TypeName(tmp) ' & " " & tmp
    Next
    Debug.Print TypeName(col("key").var1)
    

    
    Set col.Item(1) = fso

    Debug.Print TypeName(col.Item(1))

    col.Item(1) = "that"

    Debug.Print TypeName(col.Item(1))


    For Each tmp In col
        Debug.Print TypeName(tmp)
    Next

    Debug.Print col.Count

    col.Remove 1

'
'    For Each tmp In col
'        Debug.Print tmp
'    Next

    Debug.Print col.Count

    col.Add "this", "K4"
    col.Add fso, "K6"

    Dim cln As NTNodes10.Collection
    Set cln = col.Clone

    For Each tmp In cln
        Debug.Print TypeName(tmp)
    Next
    
End Sub
