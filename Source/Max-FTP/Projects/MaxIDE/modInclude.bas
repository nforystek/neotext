#Const [True] = -1
#Const [False] = 0
Attribute VB_Name = "modInclude"
#Const modInclude = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module

Public Function CheckParam(Optional Argument)

    CheckParam = IsMissing(Argument)

End Function

Public Function CheckVar(Optional Variable)
    On Error Resume Next
    Dim tmp As Object
    Set tmp = Variable
    CheckVar = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Public Function MakeVars(Optional ByVal var1 As Variant, Optional ByVal var2 As Variant, Optional ByVal var3 As Variant, Optional ByVal var4 As Variant, Optional ByVal var5 As Variant, Optional ByVal var6 As Variant, Optional ByVal var7 As Variant, Optional ByVal var8 As Variant, Optional ByVal var9 As Variant, Optional ByVal var10 As Variant) As String()
    Dim vars() As String
    If Not IsMissing(var1) Then
        ReDim Preserve vars(1 To 1) As String
        vars(1) = var1
    End If
    If Not IsMissing(var2) Then
        ReDim Preserve vars(1 To 2) As String
        vars(2) = var2
    End If
    If Not IsMissing(var3) Then
        ReDim Preserve vars(1 To 3) As String
        vars(3) = var3
    End If
    If Not IsMissing(var4) Then
        ReDim Preserve vars(1 To 4) As String
        vars(4) = var4
    End If
    If Not IsMissing(var5) Then
        ReDim Preserve vars(1 To 5) As String
        vars(5) = var5
    End If
    If Not IsMissing(var6) Then
        ReDim Preserve vars(1 To 6) As String
        vars(6) = var6
    End If
    If Not IsMissing(var7) Then
        ReDim Preserve vars(1 To 7) As String
        vars(7) = var7
    End If
    If Not IsMissing(var8) Then
        ReDim Preserve vars(1 To 8) As String
        vars(8) = var8
    End If
    If Not IsMissing(var9) Then
        ReDim Preserve vars(1 To 9) As String
        vars(9) = var9
    End If
    If Not IsMissing(var10) Then
        ReDim Preserve vars(1 To 10) As String
        vars(10) = var10
    End If
    MakeVars = vars
End Function


Attribute 