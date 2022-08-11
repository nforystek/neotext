Attribute VB_Name = "Exports"
Option Explicit
'TOP DOWN

Option Compare Binary
Option Private Module
Private Const DLL_PROCESS_DETACH = 0
Private Const DLL_PROCESS_ATTACH = 1
Private Const DLL_THREAD_ATTACH = 2
Private Const DLL_THREAD_DETACH = 3

'Define the entities you wish to export below as public
'then use the "Definition Exports menu from the "Add-Ins"
'menu after you have saved your project to select exports.

Public Function DllMain(hInst As Long, fdwReason As Long, lpvReserved As Long) As Boolean
    Select Case fdwReason
        Case DLL_PROCESS_DETACH
            DllMain = True
        Case DLL_PROCESS_ATTACH
            DllMain = True
        Case DLL_THREAD_ATTACH
            DllMain = True
        Case DLL_THREAD_DETACH
            DllMain = True
        Case Else
            DllMain = False
    End Select
End Function

Public Sub Test()
    MsgBox "Tsting..."
End Sub

Private Sub Main()

End Sub

