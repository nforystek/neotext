VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Point"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements Coord

Private pZ As Single

Public Property Get Serialize() As Boolean
Attribute Serialize.VB_MemberFlags = "40"
    Serialize = True
End Property

Public Property Get X() As Single
    X = Coord.X
End Property
Public Property Let X(ByVal RHS As Single)
    Coord.X = RHS
End Property

Public Property Get Y() As Single
    Y = Coord.Y
End Property
Public Property Let Y(ByVal RHS As Single)
    Coord.Y = RHS
End Property

Public Property Get z() As Single
    z = pZ
End Property
Public Property Let z(ByVal RHS As Single)
    pZ = RHS
End Property

Public Function Equals(ByRef Point As Point) As Boolean
    If Point Is Nothing Then
        Equals = (Coord.Equals(Point) And z = 0)
    Else
        Equals = (Coord.Equals(Point) And Point.z = z)
    End If
End Function

Public Property Get ToString() As String
Attribute ToString.VB_UserMemId = 0
Attribute ToString.VB_MemberFlags = "640"
    ToString = "[" & Round(Coord.X, 6) & ", " & Round(Coord.Y, 6) & ", " & Round(pZ, 6) & "]"
End Property
Public Property Let ToString(ByVal RHS As String)
Attribute ToString.VB_UserMemId = 0
Attribute ToString.VB_MemberFlags = "240"
    RHS = Trim(RHS)
    If Left(RHS, 1) = "[" And Right(RHS, 1) = "]" Then
        RHS = Mid(RHS, 2, Len(RHS) - 2) & ","
        Coord.X = ParseNumerical(RHS)
        Coord.Y = ParseNumerical(RHS)
        pZ = ParseNumerical(RHS)
    End If
End Property

Private Sub Class_Initialize()
    Stats_Point_Count = Stats_Point_Count + 1
    'Set Coord = New Coord
    vbaObjSetAddref Coord, ObjPtr(New Coord)
End Sub

Private Sub Class_Terminate()
    Stats_Point_Count = Stats_Point_Count - 1
    'Set Coord = Nothing
    vbaObjSet Coord, ObjPtr(Nothing)
End Sub

Private Function Coord_Equals(Point As Point) As Boolean
    Coord_Equals = Coord.Equals(Point)
End Function

Private Property Get Coord_Serialize() As Boolean
    Coord_Serialize = Coord.Serialize
End Property

Private Property Let Coord_ToString(ByVal RHS As String)
    Coord.ToString = RHS
End Property

Private Property Get Coord_ToString() As String
    Coord_ToString = Coord.ToString
End Property

Private Property Let Coord_X(ByVal RHS As Single)
    Coord.X = RHS
End Property

Private Property Get Coord_X() As Single
    Coord_X = Coord.X
End Property

Private Property Let Coord_Y(ByVal RHS As Single)
    Coord.Y = RHS
End Property

Private Property Get Coord_Y() As Single
    Coord_X = Coord.Y
End Property