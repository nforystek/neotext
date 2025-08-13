VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Plane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements Point

Private pW As Double

Public Property Get X() As Double
    X = Point.X
End Property
Public Property Let X(ByVal RHS As Double)
    Point.X = RHS
End Property

Public Property Get Y() As Double
    Y = Point.Y
End Property
Public Property Let Y(ByVal RHS As Double)
    Point.Y = RHS
End Property

Public Property Get Z() As Double
    Z = Point.Z
End Property
Public Property Let Z(ByVal RHS As Double)
    Point.Z = RHS
End Property

Public Property Get W() As Double
    W = pW
End Property
Public Property Let W(ByVal RHS As Double)
    pW = RHS
End Property

'The example debug when running this project havily depends
'on setting the "Procedure Attributes" under the Tools menu
'for the ToString() property setting to be "Procedure ID:"
'as "(Default)," checking "Hide this member," and checking
'"User Interface Default," for classes Plot, Point and Plane.

Public Property Get ToString()
Attribute ToString.VB_UserMemId = 0
Attribute ToString.VB_MemberFlags = "240"
    ToString = "[" & Round(pW, 6) & ", " & Round(Point.X, 6) & ", " & Round(Point.Y, 6) & ", " & Round(Point.Z, 6) & "]"
End Property
Public Property Let ToString(ByVal RHS)
    Select Case TypeName(RHS)
        Case "String"
            RHS = Trim(RHS)
            If Left(RHS, 1) = "[" And Right(RHS, 1) = "]" Then
                RHS = Mid(RHS, 2, Len(RHS) - 2) & ","
            Else
                If CountWord(RHS, ",") = 3 Then RHS = RHS & ","
            End If
            If CountWord(RHS, ",") = 4 Then
                pW = ParseNumerical(RHS)
                Point.X = ParseNumerical(RHS)
                Point.Y = ParseNumerical(RHS)
                Point.Z = ParseNumerical(RHS)
            End If
        Case "Plane"
            pW = RHS.W
    End Select
    Select Case TypeName(RHS)
        Case "Point", "Plane"
            Point.Z = RHS.Z
    End Select
    Select Case TypeName(RHS)
        Case "Plot", "Point", "Plane"
            Point.X = RHS.X
            Point.Y = RHS.Y
    End Select
End Property

Private Sub Class_Initialize()
    vbaObjSetAddref Point, ObjPtr(New Point)
End Sub

Private Sub Class_Terminate()
    vbaObjSet Point, ObjPtr(Nothing)
End Sub

Private Property Get Point_ToString()
    Point_ToString = Point.ToString
End Property

Private Property Let Point_ToString(ByVal RHS)
    Point.ToString = RHS
End Property

Private Property Let Point_X(ByVal RHS As Double)
    Point.X = RHS
End Property

Private Property Get Point_X() As Double
    Point_X = Point.X
End Property

Private Property Let Point_Y(ByVal RHS As Double)
    Point.Y = RHS
End Property

Private Property Get Point_Y() As Double
    Point_X = Point.Y
End Property

Private Property Let Point_Z(ByVal RHS As Double)
    Point.Z = RHS
End Property

Private Property Get Point_Z() As Double
    Point_Z = Point.Y
End Property

