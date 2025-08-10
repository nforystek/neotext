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

Implements Plot

Private pZ As Single

Public Property Get Serialize() As Boolean
Attribute Serialize.VB_MemberFlags = "40"
    Serialize = True
End Property

Public Property Get X() As Single
    X = Plot.X
End Property
Public Property Let X(ByVal RHS As Single)
    Plot.X = RHS
End Property

Public Property Get Y() As Single
    Y = Plot.Y
End Property
Public Property Let Y(ByVal RHS As Single)
    Plot.Y = RHS
End Property

Public Property Get Z() As Single
    Z = pZ
End Property
Public Property Let Z(ByVal RHS As Single)
    pZ = RHS
End Property

Public Function Equals(ByRef Point As Point) As Boolean
    If Point Is Nothing Then
        Equals = (Plot.Equals(Point) And Z = 0)
    Else
        Equals = (Plot.Equals(Point) And Point.Z = Z)
    End If
End Function

Public Property Get ToString() As String
Attribute ToString.VB_UserMemId = 0
Attribute ToString.VB_MemberFlags = "240"
    ToString = "[" & Round(Plot.X, 6) & ", " & Round(Plot.Y, 6) & ", " & Round(pZ, 6) & "]"
End Property
Public Property Let ToString(ByVal RHS As String)
    RHS = Trim(RHS)
    If Left(RHS, 1) = "[" And Right(RHS, 1) = "]" Then
        RHS = Mid(RHS, 2, Len(RHS) - 2) & ","
        If CountWord(RHS, ",") = 3 Then
            Plot.X = frmMain.Evaluate(ParseNumerical(RHS))
            Plot.Y = frmMain.Evaluate(ParseNumerical(RHS))
            pZ = frmMain.Evaluate(ParseNumerical(RHS))
        End If
    Else
        If CountWord(RHS, ",") = 2 Then RHS = RHS & ","
        If CountWord(RHS, ",") = 3 Then
            Plot.X = ParseNumerical(RHS)
            Plot.Y = ParseNumerical(RHS)
            pZ = ParseNumerical(RHS)
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
    'Set Plot = New Plot
    vbaObjSetAddref Plot, ObjPtr(New Plot)
End Sub

Private Sub Class_Terminate()
    'Set Plot = Nothing
    vbaObjSet Plot, ObjPtr(Nothing)
    Stats_Point_Count = Stats_Point_Count - 1
End Sub

Private Function Plot_Equals(Point As Point) As Boolean
    Plot_Equals = Plot.Equals(Point)
End Function

Private Property Get Plot_Serialize() As Boolean
    Plot_Serialize = Plot.Serialize
End Property

Private Property Let Plot_ToString(ByVal RHS As String)
    Plot.ToString = RHS
End Property

Private Property Get Plot_ToString() As String
    Plot_ToString = Plot.ToString
End Property

Private Property Let Plot_X(ByVal RHS As Single)
    Plot.X = RHS
End Property

Private Property Get Plot_X() As Single
    Plot_X = Plot.X
End Property

Private Property Let Plot_Y(ByVal RHS As Single)
    Plot.Y = RHS
End Property

Private Property Get Plot_Y() As Single
    Plot_Y = Plot.Y
End Property

