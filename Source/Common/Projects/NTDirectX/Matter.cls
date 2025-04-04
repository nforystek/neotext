VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Matter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Point1 As Point
Public Point2 As Point
Public Point3 As Point
Public Normal As Point

Public U1 As Single
Public V1 As Single

Public U2 As Single
Public V2 As Single

Public U3 As Single
Public V3 As Single

Public Transparent As Boolean
Public Translucent As Boolean

Public TextureIndex As Long
Public NormalIndex As Long
Public TriangleIndex As Long
Public ObjectIndex As Long
Public FaceIndex As Long
Public Serialize As Boolean

Friend Property Get Clone() As Matter
    Set Clone = New Matter
    
    Set Clone.Point1 = New Point
    Clone.Point1 = Point1
    
    Set Clone.Point2 = New Point
    Clone.Point2 = Point2
    
    Set Clone.Point3 = New Point
    Clone.Point3 = Point3
    
    Set Clone.Normal = New Point
    Clone.Normal = Normal
    
    Clone.U1 = U1
    Clone.V1 = V1
    Clone.U2 = U2
    Clone.V2 = V2
    Clone.U3 = U3
    Clone.V3 = V3
    Clone.Transparent = Transparent
    Clone.Translucent = Translucent
    Clone.TextureIndex = TextureIndex
    Clone.NormalIndex = NormalIndex
    Clone.TriangleIndex = TriangleIndex
    Clone.ObjectIndex = ObjectIndex
    Clone.FaceIndex = FaceIndex
    Clone.Serialize = Serialize
    
End Property

Private Sub Class_Initialize()
    Stats_Matter_Count = Stats_Matter_Count + 1
    TriangleIndex = -1
End Sub

Private Sub Class_Terminate()
    Stats_Matter_Count = Stats_Matter_Count - 1
'    If TriangleIndex > -1 Then
'        RemoveTriangleArray TriangleIndex
'        TriangleIndex = -1
'    End If
    Set Point1 = Nothing
    Set Point2 = Nothing
    Set Point3 = Nothing
    Set Normal = Nothing
End Sub
