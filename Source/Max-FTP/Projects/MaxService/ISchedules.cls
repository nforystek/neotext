VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISchedules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Explicit
'TOP DOWN

Option Compare Binary

Private cCol As Collection

Public Property Get Count() As Long
    Count = cCol.Count
End Property

Public Property Get Schedules() As Collection
    Set AllSchedules = cCol
End Property

Public Property Get Item(Optional ByVal ID As Variant = Nothing) As Schedule
    If TypeName(ID) = "Crayon" Then
        Set Item = cCol.Item("C" & Trim(CStr(ID.ID)))
    ElseIf IsNumeric(ID) And Not TypeName(ID) = "String" Then
        Set Item = cCol.Item("C" & Trim(CStr(ID)))
    Else
        Set Item = cCol.Item(CLng(ID))
    End If
End Property


Public Sub Clear()
    ClearCollection cCol, True, True
End Sub

Private Sub Class_Initialize()
    Set cCol = New Collection
End Sub

Private Sub Class_Terminate()
    ClearCollection cCol, True, True
    Set cCol = Nothing
End Sub
