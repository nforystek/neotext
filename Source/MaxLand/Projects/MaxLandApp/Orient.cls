VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Orient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements Point

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

Public Function Equals(ByRef Orient As Orient) As Boolean
    Equals = Point.Equals(Orient)
End Function

Public Property Get ToString() As String
    ToString = "[" & Round(Point.X, 6) & ", " & Round(Point.Y, 6) & ", " & Round(Point.Z, 6) & "]"
End Property
Public Property Let ToString(ByVal RHS As String)
    RHS = Trim(RHS)
    If Left(RHS, 1) = "[" And Right(RHS, 1) = "]" Then
        RHS = Mid(RHS, 2, Len(RHS) - 2) & ","
        Point.X = ParseNumerical(RHS)
        Point.Y = ParseNumerical(RHS)
        Point.Z = ParseNumerical(RHS)
    End If
End Property


Public Property Get Clone() As Orient
    Set Clone = New Orient
    With Clone
        .X = X
        .Y = Y
        .Z = Z
    End With
End Property

Private Sub Class_Initialize()
    Stats_Orient_Count = Stats_Orient_Count + 1
    vbaObjSetAddref Point, ObjPtr(New Point)
End Sub

Private Sub Class_Terminate()
    vbaObjSet Point, ObjPtr(Nothing)
    Stats_Orient_Count = Stats_Orient_Count - 1
End Sub

Private Property Get Point_Clone() As Point
    Set Point_Clone = Point.Clone
End Property

Private Function Point_Equals(Point As Point) As Boolean
    Point_Equals = Point.Equals(Point)
End Function


Private Property Get Point_Serialize() As Boolean
    Point_Serialize = Point.Serialize
End Property

Private Property Get Point_ToString() As String
    Point_ToString = Point.ToString
End Property

Private Property Let Point_ToString(ByVal RHS As String)
    Point.ToString = RHS
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
    Point_X = Point.Y
End Property

Private Property Let Point_Z(ByVal RHS As Single)
    Point.Z = RHS
End Property

Private Property Get Point_Z() As Single
    Point_Z = Point.Y
End Property

'Option Explicit
'
'Private pParentRotation As Point
'Private pSelfRotation As Point
'
'Public Property Get Serialize() As Boolean
'    Serialize = True
'End Property
'
'
'
'
'Public Property Get ParentPose() As Point
'
'    Set ParentPose = New Point
'
'    With ParentPose
'        .X = CDec(NextArg(CStr(pParentRotation.X), "."))
'        .Y = CDec(NextArg(CStr(pParentRotation.Y), "."))
'        .Z = CDec(NextArg(CStr(pParentRotation.Z), "."))
'    End With
'
'End Property
'
'Public Property Let ParentPose(ByRef RHS As Point)
'
'    Dim pTemp As Point
'    Set pTemp = ParentTranspose
'
'    AngleAxisRestrictDegree RHS
'
'    With pParentRotation
'        .X = CDec(CStr(RHS.X) & "." & PaddingLeft(CStr(pTemp.X), 3))
'        .Y = CDec(CStr(RHS.Y) & "." & PaddingLeft(CStr(pTemp.Y), 3))
'        .Z = CDec(CStr(RHS.Z) & "." & PaddingLeft(CStr(pTemp.Z), 3))
'    End With
'
'    Set pTemp = Nothing
'
'End Property
'Public Property Get ParentTranspose() As Point
'
'    Set ParentTranspose = New Point
'
'    With ParentTranspose
'        .X = CDec(PaddingRight(RemoveArg(CStr(pParentRotation.X), "."), 3))
'        .Y = CDec(PaddingRight(RemoveArg(CStr(pParentRotation.Y), "."), 3))
'        .Z = CDec(PaddingRight(RemoveArg(CStr(pParentRotation.Z), "."), 3))
'    End With
'
'End Property
'
'Public Property Let ParentTranspose(ByRef RHS As Point)
'
'    Dim pTemp As Point
'    Set pTemp = ParentPose
'
'    AngleAxisRestrictDegree RHS
'
'    With pParentRotation
'        .X = CDec(CStr(pTemp.X) & "." & PaddingLeft(CStr(RHS.X), 3))
'        .Y = CDec(CStr(pTemp.Y) & "." & PaddingLeft(CStr(RHS.Y), 3))
'        .Z = CDec(CStr(pTemp.Z) & "." & PaddingLeft(CStr(RHS.Z), 3))
'    End With
'
'    Set pTemp = Nothing
'
'End Property
'
'
'
'
'Public Property Get SelfPose() As Point
'
'    Set SelfPose = New Point
'
'    With SelfPose
'        .X = CDec(NextArg(CStr(pSelfRotation.X), "."))
'        .Y = CDec(NextArg(CStr(pSelfRotation.Y), "."))
'        .Z = CDec(NextArg(CStr(pSelfRotation.Z), "."))
'    End With
'
'End Property
'Public Property Let SelfPose(ByRef RHS As Point)
'
'    Dim pTemp As Point
'    Set pTemp = SelfTranspose
'
'    AngleAxisRestrictDegree RHS
'
'    With pSelfRotation
'        .X = CDec(CStr(RHS.X) & "." & PaddingLeft(CStr(pTemp.X), 3))
'        .Y = CDec(CStr(RHS.Y) & "." & PaddingLeft(CStr(pTemp.Y), 3))
'        .Z = CDec(CStr(RHS.Z) & "." & PaddingLeft(CStr(pTemp.Z), 3))
'    End With
'
'    Set pTemp = Nothing
'
'End Property
'
'Public Property Get SelfTranspose() As Point
'
'    Set SelfTranspose = New Point
'
'    With SelfTranspose
'        .X = CDec(PaddingRight(RemoveArg(CStr(pSelfRotation.X), "."), 3))
'        .Y = CDec(PaddingRight(RemoveArg(CStr(pSelfRotation.Y), "."), 3))
'        .Z = CDec(PaddingRight(RemoveArg(CStr(pSelfRotation.Z), "."), 3))
'    End With
'
'End Property
'Public Property Let SelfTranspose(ByRef RHS As Point)
'
'    Dim pTemp As Point
'    Set pTemp = SelfPose
'
'    AngleAxisRestrictDegree RHS
'
'    With pSelfRotation
'        .X = CDec(CStr(pTemp.X) & "." & PaddingLeft(CStr(RHS.X), 3))
'        .Y = CDec(CStr(pTemp.Y) & "." & PaddingLeft(CStr(RHS.Y), 3))
'        .Z = CDec(CStr(pTemp.Z) & "." & PaddingLeft(CStr(RHS.Z), 3))
'    End With
'
'    Set pTemp = Nothing
'
'End Property
'
'
'
'
'
'
'
'
'
'
'Public Property Get ParentRotation() As Point
'    Set ParentRotation = pParentRotation
'End Property
'Public Property Let ParentRotation(ByRef RHS As Point)
'    Set pParentRotation = RHS
'End Property
'
'
'Public Property Get SelfRotation() As Point
'    Set SelfRotation = pSelfRotation
'End Property
'Public Property Let SelfRotation(ByRef RHS As Point)
'    Set pSelfRotation = RHS
'End Property
'
'
'Public Function Equals(ByRef Orient As Orient) As Boolean
'    Equals = (Orient.ToString = ToString)
'End Function
'
'Public Property Get ToString() As String
'    ToString = "[" & pParentRotation.X & ", " & pParentRotation.Y & ", " & pParentRotation.Z & ", " & pSelfRotation.X & ", " & pSelfRotation.Y & ", " & pSelfRotation.Z & "]"
'End Property
'Public Property Let ToString(ByVal RHS As String)
'    RHS = Trim(RHS)
'    If Left(RHS, 1) = "[" And Right(RHS, 1) = "]" Then
'        RHS = Mid(RHS, 2, Len(RHS) - 2) & ","
'        If CountWord(RHS, ",") = 3 Then
'            pParentRotation.X = frmMain.Evaluate(ParseNumerical(RHS))
'            pParentRotation.Y = frmMain.Evaluate(ParseNumerical(RHS))
'            pParentRotation.Z = frmMain.Evaluate(ParseNumerical(RHS))
'            pSelfRotation.X = frmMain.Evaluate(ParseNumerical(RHS))
'            pSelfRotation.Y = frmMain.Evaluate(ParseNumerical(RHS))
'            pSelfRotation.Z = frmMain.Evaluate(ParseNumerical(RHS))
'        End If
'    Else
'        If CountWord(RHS, ",") = 2 Then RHS = RHS & ","
'        If CountWord(RHS, ",") = 3 Then
'            pParentRotation.X = ParseNumerical(RHS)
'            pParentRotation.Y = ParseNumerical(RHS)
'            pParentRotation.Z = ParseNumerical(RHS)
'            pSelfRotation.X = ParseNumerical(RHS)
'            pSelfRotation.Y = ParseNumerical(RHS)
'            pSelfRotation.Z = ParseNumerical(RHS)
'        End If
'    End If
'End Property
'
''Public Property Get Clone() As Orient
''    Set Clone = New Orient
''    With Clone
''        .Rotation = Rotation
''    End With
''End Property
'
'Private Sub Class_Initialize()
'    Stats_Orient_Count = Stats_Orient_Count + 1
''    'Set Plot = New Plot
''    vbaObjSetAddref Plot, ObjPtr(New Plot)
'End Sub
'
'Private Sub Class_Terminate()
''    'Set Plot = Nothing
''    vbaObjSet Plot, ObjPtr(Nothing)
'    Stats_Orient_Count = Stats_Orient_Count - 1
'End Sub



