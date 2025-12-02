Attribute VB_Name = "modMain"



#Const modMain = -1
Option Explicit
Option Compare Binary

Public Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)


Public Sub Main()
    If KeyContainer = "" Then
        KeyContainer = "chrome@winternet.com" ' & vbNullChar '"{" & modGuid.GUID & "}"
    End If
End Sub














