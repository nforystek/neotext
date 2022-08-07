VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OnEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pApplyTo As NTNodes10.Collection

Private pCode As String
Public Event BeforeEvent(ByRef cancel As Boolean)


Public Property Get ApplyTo() 'As NTNodes10.Collection
    Set ApplyTo = pApplyTo
End Property

Public Property Let ApplyTo(ByVal RHS)
    If pApplyTo Is Nothing Then
        Set pApplyTo = New NTNodes10.Collection
    Else
        pApplyTo.Clear
    End If
    Dim inVal As String
    Do Until RHS = ""
        inVal = RemoveNextArg(RHS, ",")
        inVal = UCase(Left(inVal, 1)) & LCase(Mid(inVal, 2))
        pApplyTo.Add inVal, inVal
    Loop

End Property

Public Property Get Code() As String
    Code = pCode
End Property
Public Property Let Code(ByVal RHS As String)
    pCode = RHS
End Property
Public Property Let ToString(ByVal RHS As String)
    If pApplyTo Is Nothing Then
        Set pApplyTo = New NTNodes10.Collection
    Else
        pApplyTo.Clear
    End If
    Dim inName As String
    Dim inVal As String
    inName = RemoveQuotedArg(RHS, "<", ">")
    Do Until inName = ""
        inVal = RemoveNextArg(inName, ",")
        inVal = UCase(Left(inVal, 1)) & LCase(Mid(inVal, 2))
        pApplyTo.Add inVal, inVal
    Loop
    pCode = ParseWhiteSpace(RHS)
End Property
Public Property Get ToString() As String
    If pApplyTo.Count > 0 Then
        ToString = ToString & "<"
        Dim nam As Variant
        For Each nam In pApplyTo
            ToString = ToString & nam & ","
        Next
        ToString = Left(ToString, Len(ToString) - 1) & ">"
    End If
    ToString = ToString & vbCrLf & pCode & vbCrLf
End Property
Public Function RunEvent() As Boolean
    Dim cancel As Boolean
    RaiseEvent BeforeEvent(cancel)
    If Not cancel Then
        frmMain.RunEvent pCode
        RunEvent = True
    End If
End Function

Private Sub Class_Terminate()
    Set pApplyTo = Nothing
End Sub
