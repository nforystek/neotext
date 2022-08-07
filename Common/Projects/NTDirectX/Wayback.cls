VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Wayback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IUnknown

Public Default As Molecule

Private List As Wayback

Private pOrigin As Point
Private pScaled As Point
Private pOffset As Point
Private pRotate As Point

Public Property Get Origin() As Point
    If pOrigin Is Nothing Then Set pOrigin = New Point
    Set Origin = pOrigin
End Property
Public Property Set Origin(ByRef RHS As Point)
    Set pOrigin = RHS
End Property

Public Property Get Scaled() As Point
    If pScaled Is Nothing Then Set pScaled = New Point
    Set Scaled = pScaled
End Property
Public Property Set Scaled(ByRef RHS As Point)
    Set pScaled = RHS
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

Public Property Get NotEmpty() As Boolean
    NotEmpty = Not (List Is Nothing)
End Property

Public Sub Stack(ByRef RHS As Wayback)
    If List Is Nothing Then
        Set List = RHS
    Else
        List.Stack RHS
    End If
End Sub

Public Function Remove() As Wayback
    If List Is Nothing Then
        Set Remove = Nothing
    Else
        Set Remove = List.Remove
        If Remove Is Nothing Then
            Set Remove = List
            Set List = Nothing
        End If
    End If
End Function

