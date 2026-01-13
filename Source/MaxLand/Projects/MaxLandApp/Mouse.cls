VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements Point

Friend Sub LateEval()

End Sub

Public Property Get Serialize() As Boolean
    Serialize = True
End Property

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

Public Property Get Z() As Single
    Z = Point.Z
End Property
Public Property Let Z(ByVal RHS As Single)
    Point.Z = RHS
End Property

Public Function Equals(ByRef Point As Point) As Boolean
    Equals = Point.Equals(Me)
End Function

Public Property Get ToString() As String
    ToString = "[" & Round(Point.X, 6) & ", " & Round(Point.Y, 6) & ", " & Round(Point.Z, 6) & "]"
End Property
Public Property Let ToString(ByVal RHS As String)
    RHS = Trim(RHS)
    If Left(RHS, 1) = "[" And Right(RHS, 1) = "]" Then
        RHS = Mid(RHS, 2, Len(RHS) - 2) & ","
        If CountWord(RHS, ",") = 3 Then
            Point.X = frmMain.Evaluate(ParseNumerical(RHS))
            Point.Y = frmMain.Evaluate(ParseNumerical(RHS))
            Point.Z = frmMain.Evaluate(ParseNumerical(RHS))
        End If
    ElseIf Left(RHS, 1) = "{" And Right(RHS, 1) = "}" Then
        RHS = Mid(RHS, 2, Len(RHS) - 2) & ","
        If CountWord(RHS, ",") = 3 Then
            Point.X = frmMain.Evaluate(ParseNumerical(RHS))
            Point.Y = frmMain.Evaluate(ParseNumerical(RHS))
            Point.Z = frmMain.Evaluate(ParseNumerical(RHS))
        End If
    ElseIf Left(RHS, 1) = """" And Right(RHS, 1) = """" Then
    Else 'If Left(RHS, 1) = """" And Right(RHS, 1) = """" Then
        If CountWord(RHS, ",") = 2 Then RHS = RHS & ","
        If CountWord(RHS, ",") = 3 Then
            Point.X = ParseNumerical(RHS)
            Point.Y = ParseNumerical(RHS)
            Point.Z = ParseNumerical(RHS)
        End If
    End If
End Property

Public Property Get Clone() As Point
    Set Clone = New Point
    With Clone
        .X = X
        .Y = Y
        .Z = Z
    End With
End Property

Private Sub Class_Initialize()
    Stats_Point_Count = Stats_Point_Count + 1

    vbaObjSetAddref Point, ObjPtr(New Point)
End Sub

Private Sub Class_Terminate()

    vbaObjSet Point, ObjPtr(Nothing)
    Stats_Point_Count = Stats_Point_Count - 1
End Sub

Private Function Plot_Equals(Point As Point) As Boolean
    Plot_Equals = Plot.Equals(Point)
End Function

Private Property Get Point_Clone() As Point
    Point_Clone = Point.Clone
End Property

Private Function Point_Equals(Point As Point) As Boolean
    point_eqwuals = Point.Equals(Me)
End Function

Private Property Get Point_Serialize() As Boolean
    Point_Serialize = Point.Serialize
End Property

Private Property Let Point_ToString(ByVal RHS As String)
    Point.ToString = RHS
End Property

Private Property Get Point_ToString() As String
    Point_ToString = Point.ToString
End Property

Private Property Let Point_X(ByVal RHS As Single)
    Point.X = RHS
End Property

Private Property Get Point_X() As Single
    Point_X = Point.X
End Property

Private Property Let Point_Y(ByVal RHS As Single)
    Point.Y = RHS
End Property

Private Property Get Point_Y() As Single
    Point_Y = Point.Y
End Property

Private Property Let Point_Z(ByVal RHS As Single)
    Point.Z = RHS
End Property

Private Property Get Point_Z() As Single
    Point_Z = Point.Z
End Property
