VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Helper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Holder As Object

Public Function ObjPtr(ByRef Obj) As Long
    If Not VBA.IsMissing(Obj) Then
        If VBA.IsObject(Obj) Then
            RtlMoveMemory ObjPtr, VBA.ObjPtr(Obj), 4
        End If
    End If
End Function

Public Function VarPtr(ByRef Obj) As Long
    If Not VBA.IsMissing(Obj) Then
        If Not VBA.IsObject(Obj) Then
            RtlMoveMemory VarPtr, VBA.VarPtr(Obj), 4
        End If
    End If
End Function


