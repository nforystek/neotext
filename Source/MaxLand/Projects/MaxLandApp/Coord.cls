VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Coord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pU As Single
Private pV As Single

Public Property Get Serialize() As Boolean
    Serialize = True
End Property

Public Property Get U() As Single
    U = pU
End Property
Public Property Let U(ByVal RHS As Single)
    pU = RHS
End Property

Public Property Get V() As Single
    V = pV
End Property
Public Property Let V(ByVal RHS As Single)
    pV = RHS
End Property

Public Function Equals(ByRef Coord As Coord) As Boolean
    If Coord Is Nothing Then
        Equals = (pU = 0 And pV = 0)
    Else
        Equals = (pU = Coord.U And pV = Coord.V)
    End If
End Function

Public Property Get ToString() As String
    ToString = "[" & Round(pU, 6) & ", " & Round(pV, 6) & "]"
End Property
Public Property Let ToString(ByVal RHS As String)

    RHS = Trim(RHS)
    If Left(RHS, 1) = "[" And Right(RHS, 1) = "]" Then
        RHS = Mid(RHS, 2, Len(RHS) - 2) & ","
        If CountWord(RHS, ",") = 2 Then
            pU = frmMain.Evaluate(ParseNumerical(RHS))
            pV = frmMain.Evaluate(ParseNumerical(RHS))
        End If
    Else
        If InStr(RHS, ",") = 1 Then RHS = RHS & ","
        If CountWord(RHS, ",") = 2 Then
            pU = ParseNumerical(RHS)
            pV = ParseNumerical(RHS)
        End If
    End If
    
End Property


Private Sub Class_Initialize()
    Stats_Coord_Count = Stats_Coord_Count + 1
End Sub

Private Sub Class_Terminate()
    Stats_Coord_Count = Stats_Coord_Count - 1
End Sub

