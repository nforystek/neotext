#Const [True] = -1
#Const [False] = 0
Attribute VB_Name = "modMain"
#Const modMain = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public DBConn As New clsDBConnection
Public dbSettings As New clsDBSettings

Public Function GetTextObj()

    Set GetTextObj = frmMain.CodeEdit1

End Function

Public Sub Main()

    '%LICENSE%

tryit: On Error GoTo catch
    
    If Not App.PrevInstance Then

        frmMain.Show
        
    End If

GoTo final
catch: On Error GoTo 0

    'If Err Then MsgBox Err.Description, vbExclamation, App.EXEName

final: On Error Resume Next


On Error GoTo -1
End Sub

Public Function IsOnList(tList, Item) As Integer
    Dim cnt As Integer
    Dim ItemFound As Integer
    cnt = 0
    ItemFound = -1
    Do Until cnt = tList.ListCount Or ItemFound > -1
        If LCase(Trim(tList.List(cnt))) = LCase(Trim(Item)) Then ItemFound = cnt
        cnt = cnt + 1
    Loop
    IsOnList = ItemFound
End Function


