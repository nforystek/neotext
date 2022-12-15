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

Public Index1 As Long
Public Index2 As Long
Public Index3 As Long

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
