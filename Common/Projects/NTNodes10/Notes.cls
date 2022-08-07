VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Notes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Public Function CreateList() As Note
    Set CreateList = New Note
End Function

Public Sub DisposeList(ByRef Note As Note)
    If Not Note Is Nothing Then
        Note.Clear
        Set Note = Nothing
    End If
End Sub

Public Sub Delete(ByRef Note As Note)
    Note.Delete
End Sub

Public Sub Insert(ByRef Note As Note)
    Note.Insert
End Sub

Public Sub Forth(ByRef Note As Note)
    Set Note = Note.Forth
End Sub

Public Sub Prior(ByRef Note As Note)
    Set Note = Note.Prior
End Sub


