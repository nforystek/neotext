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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pOrigin As Point
Private pOffset As Point
Private pScaled As Point
Private pRotate As Point

Private pDirect As Point
Private pTwists As Point
Private pScalar As Point

Private pKey As String
Private pAttachments As NTNodes10.Collection
Private pAttachedTo As String

Public Property Get Origin() As Point
    Set Origin = pOrigin
End Property
Public Property Set Origin(ByRef RHS As Point)
    Set pOrigin = RHS
End Property

Public Property Get Offset() As Point
    Set Offset = pOffset
End Property
Public Property Set Offset(ByRef RHS As Point)
    Set pOffset = RHS
End Property

Public Property Get Scaled() As Point
    Set Scaled = pScaled
End Property
Public Property Set Scaled(ByRef RHS As Point)
    Set pScaled = RHS
End Property

Public Property Get Rotate() As Point
    Set Rotate = pRotate
End Property
Public Property Set Rotate(ByRef RHS As Point)
    Set pRotate = RHS
End Property


Friend Property Get Direct() As Point
    Set Direct = pDirect
End Property
Friend Property Set Direct(ByRef RHS As Point)
    Set pDirect = RHS
End Property

Friend Property Get Twists() As Point
    Set Twists = pTwists
End Property
Friend Property Set Twists(ByRef RHS As Point)
    Set pTwists = RHS
End Property

Friend Property Get Scalar() As Point
    Set Scalar = pScalar
End Property
Friend Property Set Scalar(ByRef RHS As Point)
    Set pScalar = RHS
End Property

Public Property Get Attachments() As NTNodes10.Collection
    Set Attachments = pAttachments
End Property
Public Property Set Attachments(ByRef RHS As NTNodes10.Collection)
    Set pAttachments = RHS
End Property

Public Property Get Key() As String
    Key = pKey
End Property
Public Property Let Key(ByVal RHS As String)
    If RHS <> "" Then
        If ParseReservedWord(RHS) <> 0 Then
            Err.Raise 457, App.Title, "This key is already associated with an element of this collection, or a global declaration restriction."
        Else
            If pKey <> RHS And pKey <> "" Then If All.Exists(pKey) Then All.Remove pKey
            If All.Exists(RHS) Then
                All.Remove RHS
                All.Add Me, RHS
            End If
            pKey = RHS
        End If
    End If
End Property

Public Property Get AttachedTo() As String
    AttachedTo = pAttachedTo
End Property
Public Property Let AttachedTo(ByVal RHS As String)
    If pAttachedTo <> "" Then
        If All.Exists(pAttachedTo) Then
            If All(pAttachedTo).Attachments.Exists(Me.Key) Then
                All(pAttachedTo).Attachments.Remove Me.Key
            End If
        End If
    End If
    pAttachedTo = RHS
    If RHS <> "" Then
        If All(pAttachedTo).Attachments Is Nothing Then
            Set All(pAttachedTo).Attachments = New NTNodes10.Collection
        End If
        All(pAttachedTo).Attachments.Add Me, RHS
    End If
End Property

Public Property Get Serialize() As Boolean
    Serialize = True
End Property

'Public Function Equals(ByRef Orient As Orient) As Boolean
'    Equals = Point.Equals(Orient)
'End Function


'Public Property Get ToString() As String
'    ToString = "[" & Round(Point.X, 6) & ", " & Round(Point.Y, 6) & ", " & Round(Point.Z, 6) & "]"
'End Property
'Public Property Let ToString(ByVal RHS As String)
'    RHS = Trim(RHS)
'    If Left(RHS, 1) = "[" And Right(RHS, 1) = "]" Then
'        RHS = Mid(RHS, 2, Len(RHS) - 2) & ","
'        Point.X = ParseNumerical(RHS)
'        Point.Y = ParseNumerical(RHS)
'        Point.Z = ParseNumerical(RHS)
'    End If
'End Property


Public Property Get Clone() As Orient
'    Set Clone = New Orient
'    With Clone
'        .X = X
'        .Y = Y
'        .Z = Z
'    End With
End Property

Private Sub Class_Initialize()
    Stats_Orient_Count = Stats_Orient_Count + 1
    Set pOrigin = New Point
    Set pScaled = MakePoint(1, 1, 1)
    Set pRotate = New Point
    Set pOffset = New Point
    
    Set pDirect = New Point
    Set pTwists = New Point
    Set pScalar = New Point
End Sub

Private Sub Class_Terminate()
    Stats_Orient_Count = Stats_Orient_Count - 1
    
    AttachedTo = "" 'relocalizes
    
    If Not pAttachments Is Nothing Then
        Dim e As Element
        Do Until pAttachments.Count = 0
            Set e = pAttachments(1)
            pAttachments.Remove 1
            Set e = Nothing
        Loop
        Set pAttachments = Nothing
    End If
    
    Set pOrigin = Nothing
    Set pScaled = Nothing
    Set pRotate = Nothing
    Set pOffset = Nothing
    Set pDirect = Nothing
    Set pTwists = Nothing
    Set pScalar = Nothing
End Sub



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



