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

Private pX As Single
Private pY As Single

Public Property Get Serialize() As Boolean
Attribute Serialize.VB_MemberFlags = "40"
    Serialize = True
End Property

Public Property Get X() As Single
    X = pX
End Property
Public Property Let X(ByVal RHS As Single)
    pX = RHS
End Property

Public Property Get Y() As Single
    Y = pY
End Property
Public Property Let Y(ByVal RHS As Single)
    pY = RHS
End Property

Public Function Equals(ByRef Point As Point) As Boolean
    If Point Is Nothing Then
        Equals = (pX = 0 And pY = 0)
    Else
        Equals = (Point.X = pX And Point.Y = pY)
    End If
End Function

Public Property Get ToString() As String
Attribute ToString.VB_UserMemId = 0
Attribute ToString.VB_MemberFlags = "240"
    ToString = "[" & Round(pX, 6) & ", " & Round(pY, 6) & "]"
End Property
Public Property Let ToString(ByVal RHS As String)

    RHS = Trim(RHS)
    If Left(RHS, 1) = "[" And Right(RHS, 1) = "]" Then
        RHS = Mid(RHS, 2, Len(RHS) - 2) & ","
        If CountWord(RHS, ",") = 2 Then
            pX = frmMain.Evaluate(ParseNumerical(RHS))
            pY = frmMain.Evaluate(ParseNumerical(RHS))
        End If
    Else
        If InStr(RHS, ",") = 1 Then RHS = RHS & ","
        If CountWord(RHS, ",") = 2 Then
            pX = ParseNumerical(RHS)
            pY = ParseNumerical(RHS)
        End If
    End If
    
End Property


Private Sub Class_Initialize()
    Stats_Coord_Count = Stats_Coord_Count + 1
End Sub

Private Sub Class_Terminate()
    Stats_Coord_Count = Stats_Coord_Count - 1
End Sub
