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

Private pParentRotation As Point
Private pSelfRotation As Point

Public Property Get Serialize() As Boolean
Attribute Serialize.VB_MemberFlags = "40"
    Serialize = True
End Property




Public Property Get ParentPose() As Point

End Property
Public Property Let ParentPose(ByRef RHS As Point)

End Property
Public Property Get ParentTranspose() As Point

End Property
Public Property Let ParentTranspose(ByRef RHS As Point)

End Property




Public Property Get SelfPose() As Point

End Property
Public Property Let SelfPose(ByRef RHS As Point)

End Property
Public Property Get SelfTranspose() As Point

End Property
Public Property Let SelfTranspose(ByRef RHS As Point)

End Property










Public Property Get ParentRotation() As Point

End Property
Public Property Let ParentRotation(ByRef RHS As Point)

End Property


Public Property Get SelfRotation() As Point

End Property
Public Property Let SelfRotation(ByRef RHS As Point)

End Property


Public Function Equals(ByRef Orient As Orient) As Boolean
    Equals = (Orient.ToString = ToString)
End Function

Public Property Get ToString() As String
Attribute ToString.VB_UserMemId = 0
Attribute ToString.VB_MemberFlags = "240"
    ToString = "[" & pParentRotation.X & ", " & pParentRotation.Y & ", " & pParentRotation.Z & ", " & pSelfRotation.X & ", " & pSelfRotation.Y & ", " & pSelfRotation.Z & "]"
End Property
Public Property Let ToString(ByVal RHS As String)
    RHS = Trim(RHS)
    If Left(RHS, 1) = "[" And Right(RHS, 1) = "]" Then
        RHS = Mid(RHS, 2, Len(RHS) - 2) & ","
        If CountWord(RHS, ",") = 3 Then
            pParentRotation.X = frmMain.Evaluate(ParseNumerical(RHS))
            pParentRotation.Y = frmMain.Evaluate(ParseNumerical(RHS))
            pParentRotation.Z = frmMain.Evaluate(ParseNumerical(RHS))
            pSelfRotation.X = frmMain.Evaluate(ParseNumerical(RHS))
            pSelfRotation.Y = frmMain.Evaluate(ParseNumerical(RHS))
            pSelfRotation.Z = frmMain.Evaluate(ParseNumerical(RHS))
        End If
    Else
        If CountWord(RHS, ",") = 2 Then RHS = RHS & ","
        If CountWord(RHS, ",") = 3 Then
            pParentRotation.X = ParseNumerical(RHS)
            pParentRotation.Y = ParseNumerical(RHS)
            pParentRotation.Z = ParseNumerical(RHS)
            pSelfRotation.X = ParseNumerical(RHS)
            pSelfRotation.Y = ParseNumerical(RHS)
            pSelfRotation.Z = ParseNumerical(RHS)
        End If
    End If
End Property

Public Property Get Clone() As Orient
    Set Clone = New Orient
    With Clone
        .Rotation = Rotation
    End With
End Property

Private Sub Class_Initialize()
    Stats_Orient_Count = Stats_Orient_Count + 1
'    'Set Plot = New Plot
'    vbaObjSetAddref Plot, ObjPtr(New Plot)
End Sub

Private Sub Class_Terminate()
'    'Set Plot = Nothing
'    vbaObjSet Plot, ObjPtr(Nothing)
    Stats_Orient_Count = Stats_Orient_Count - 1
End Sub

