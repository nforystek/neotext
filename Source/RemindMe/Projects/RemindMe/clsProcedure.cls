VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary


Public Parameters As Collection

Public ProcedureName As String

Public Sub AddParameter(ByVal ParamName As String, ByVal ParamType As String)
    Dim newParam As New clsParameter
    newParam.ParamType = ParamType
    newParam.ParamName = ParamName
    Parameters.Add newParam, ParamName
    Set newParam = Nothing
End Sub
Public Sub ClearParameters()
    Do Until Parameters.Count = 0
        Parameters.Remove 1
    Loop
End Sub

Private Sub Class_Initialize()
    Set Parameters = New Collection
End Sub
Private Sub Class_Terminate()
    ClearParameters
    Set Parameters = Nothing
End Sub
