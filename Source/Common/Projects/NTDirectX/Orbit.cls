VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Orbit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pOrigin As Point
Private pOffset As Point
Private pRotate As Point
Private pScaled As Point
Private pRanges As Range

Public Visible As Boolean

Private pParent As Orbit
Private pRelative As Orbit
Private pAbsolute As Orbit


Friend Property Get Parent() As Orbit
    Set Parent = pParent
End Property
Friend Property Set Parent(ByRef RHS As Orbit)
    Set pParent = RHS
End Property

Friend Property Get Relative() As Orbit
    If pRelative Is Nothing Then Set pRelative = New Orbit
    Set Relative = pRelative
End Property

Friend Property Set Relative(ByRef RHS As Orbit)
    Set pRelative = RHS
End Property

Friend Property Get Absolute() As Orbit
    If pAbsolute Is Nothing Then Set pAbsolute = New Orbit
    Set Absolute = pAbsolute
End Property

Friend Property Set Absolute(ByRef RHS As Orbit)
    Set pAbsolute = RHS
End Property

Public Property Get Origin() As Point
    If pOrigin Is Nothing Then Set pOrigin = New Point
    Set Origin = pOrigin
End Property
Public Property Set Origin(ByRef RHS As Point)
    Set pOrigin = RHS
End Property

Public Property Get Offset() As Point
    If pOffset Is Nothing Then Set pOffset = New Point
    Set Offset = pOffset
End Property
Public Property Set Offset(ByRef RHS As Point)
    Set pOffset = RHS
End Property

Public Property Get Rotate() As Point
    If pRotate Is Nothing Then Set pRotate = New Point
    Set Rotate = pRotate
End Property
Public Property Set Rotate(ByRef RHS As Point)
    Set pRotate = RHS
End Property


Public Property Get Scaled() As Point
    If pScaled Is Nothing Then
        Set pScaled = New Point
        pScaled.X = 1
        pScaled.Y = 1
        pScaled.z = 1
    End If
    Set Scaled = pScaled
End Property
Public Property Set Scaled(ByRef RHS As Point)
    Set pScaled = RHS
End Property

Public Property Get Ranges() As Range
    If pRanges Is Nothing Then Set pRanges = New Range
    Set Ranges = pRanges
End Property
Public Property Set Ranges(ByRef RHS As Range)
    Set pRanges = RHS
End Property

Public Property Get Serialize() As String
    Serialize = "[" & Origin.Serialize & Offset.Serialize & Rotate.Serialize & Scaled.Serialize & Ranges.Serialize & "]"
End Property
Public Property Let Serialize(ByVal RHS As String)
    RHS = Trim(RHS)
    If Left(RHS, 1) = "[" And Right(RHS, 1) = "]" Then
        RHS = Mid(RHS, 2, Len(RHS) - 2)
        If Left(RHS, 1) = "[" And Right(RHS, 1) = "]" Then
            Origin.Serialize = Left(RHS, InStr(RHS, "]"))
            RHS = Mid(RHS, InStr(RHS, "]") + 1)
        End If
        If Left(RHS, 1) = "[" And Right(RHS, 1) = "]" Then
            Offset.Serialize = Left(RHS, InStr(RHS, "]"))
            RHS = Mid(RHS, InStr(RHS, "]") + 1)
        End If
        If Left(RHS, 1) = "[" And Right(RHS, 1) = "]" Then
            Rotate.Serialize = Left(RHS, InStr(RHS, "]"))
            RHS = Mid(RHS, InStr(RHS, "]") + 1)
        End If
        If Left(RHS, 1) = "[" And Right(RHS, 1) = "]" Then
            Scaled.Serialize = Left(RHS, InStr(RHS, "]"))
            RHS = Mid(RHS, InStr(RHS, "]") + 1)
        End If
        If Left(RHS, 1) = "[" And Right(RHS, 1) = "]" Then
            Ranges.Serialize = Left(RHS, InStr(RHS, "]"))
            RHS = Mid(RHS, InStr(RHS, "]") + 1)
        End If
    End If
End Property



Private Sub Class_Terminate()
    Set pOrigin = Nothing
    Set pOffset = Nothing
    Set pRotate = Nothing
    Set pScaled = Nothing
    Set pRanges = Nothing
    Set pParent = Nothing
    Set pRelative = Nothing
    Set pAbsolute = Nothing
End Sub
