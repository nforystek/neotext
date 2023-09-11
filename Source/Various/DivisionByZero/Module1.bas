Attribute VB_Name = "Module1"
Option Explicit

'This project is a rather odd project to solve "Division by Zero" errors, as if.
'Not sure how I came across the potential resuming in these cases return values,
'that the Resume describes in which way a zero is used in division, 0/0 or #/0.

Public Sub Main()

    Debug.Print "0 / 9 = " & Divide(0, 9) ' 0
    Debug.Print "9 / 0 = " & Divide(9, 0) ' 1.#INF
    Debug.Print "9 / 9 = " & Divide(9, 9) ' 1
    Debug.Print "0 / 0 = " & Divide(0, 0) '-1.#IND
    
    Debug.Print "1 / 9 = " & Divide(1, 9) ' 0.1111111
    Debug.Print "9 / 1 = " & Divide(9, 1) ' 9
    Debug.Print "9 / 9 = " & Divide(9, 9) ' 1
    Debug.Print "1 / 1 = " & Divide(1, 1) ' 1

    Debug.Print "1 / 0 = " & Divide(1, 0) ' 1.#INF
    Debug.Print "0 / 1 = " & Divide(0, 1) ' 0
    Debug.Print "0 / 0 = " & Divide(0, 0) '-1.#IND
    Debug.Print "1 / 1 = " & Divide(1, 1) ' 1

End Sub

Public Function Divide(ByRef arg1 As Single, ByRef arg2 As Single) As Single
    On Local Error Resume Next
    On Local Error GoTo -1
    On Error Resume Next
    On Error GoTo errorlvl2
    Divide = Err.Number + (arg1 / arg2)
errorlvl1:
    On Error GoTo 0
    On Local Error GoTo 0
    Exit Function
errorlvl2:
    On Error GoTo -1
    On Local Error GoTo errorlvl1
    Err.Clear: Resume
End Function

