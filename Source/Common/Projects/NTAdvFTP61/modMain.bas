Attribute VB_Name = "modMain"
#Const [True] = -1
#Const [False] = 0






#Const modMain = -1
Option Explicit
Option Compare Binary

Public Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)


Public Sub Main()
    If KeyContainer = "" Then
        KeyContainer = "{" & modGuid.GUID & "}"
    End If
End Sub














