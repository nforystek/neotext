VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Range"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements Point

Private pW As Single

Public Property Get X() As Single
    X = Point.X
End Property
Public Property Let X(ByVal RHS As Single)
    Point.X = RHS
End Property

Public Property Get Y() As Single
    Y = Point.Y
End Property
Public Property Let Y(ByVal RHS As Single)
    Point.Y = RHS
End Property

Public Property Get z() As Single
    z = Point.z
End Property
Public Property Let z(ByVal RHS As Single)
    Point.z = RHS
End Property

Public Property Get W() As Single
    W = pW
End Property
Public Property Let W(ByVal RHS As Single)
    pW = RHS
End Property

Public Function Equals(ByRef Range As Range) As Boolean
    Equals = (Point.Equals(Range) And Range.W = W)
End Function

Public Property Get Serialize() As String
Attribute Serialize.VB_UserMemId = 0
Attribute Serialize.VB_MemberFlags = "200"
    Serialize = "[" & Round(X, 6) & ", " & Round(Y, 6) & ", " & Round(z, 6) & ", " & Round(W, 6) & "]"
End Property
Public Property Let Serialize(ByVal RHS As String)
    RHS = RemoveQuotedArg(RHS, "[", "]")
    If IsNumeric(NextArg(RHS, ",")) Then
        X = CSng(NextArg(RHS, ","))
    End If
    If IsNumeric(NextArg(RemoveArg(RHS, ","), ",")) Then
        Y = CSng(NextArg(RemoveArg(RHS, ","), ","))
    End If
    If IsNumeric(NextArg(RemoveArg(RemoveArg(RHS, ","), ","), ",")) Then
        z = CSng(NextArg(RemoveArg(RemoveArg(RHS, ","), ","), ","))
    End If
    If IsNumeric(NextArg(RemoveArg(RemoveArg(RemoveArg(RHS, ","), ","), ","), ",")) Then
        W = CSng(NextArg(RemoveArg(RemoveArg(RemoveArg(RHS, ","), ","), ","), ","))
    End If
End Property

Private Sub Class_Initialize()
    vbaObjSetAddref Point, ObjPtr(New Point)

End Sub

Private Function Point_Equals(Point As Point) As Boolean
    Point_Equals = Me.Equals(Point)
End Function

Private Property Let Point_Serialize(ByVal RHS As String)
    Me.Serialize = RHS
End Property

Private Property Get Point_Serialize() As String
    Point_Serialize = Me.Serialize
End Property

Private Property Let Point_X(ByVal RHS As Single)
    Me.X = RHS
End Property

Private Property Get Point_X() As Single
    Point_X = Me.X
End Property

Private Property Let Point_Y(ByVal RHS As Single)
    Me.Y = RHS
End Property

Private Property Get Point_Y() As Single
    Point_X = Me.Y
End Property

Private Property Let Point_Z(ByVal RHS As Single)
    Me.z = RHS
End Property

Private Property Get Point_Z() As Single
    Point_Z = Me.Y
End Property
