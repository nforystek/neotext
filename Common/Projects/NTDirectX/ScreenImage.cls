VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Visible As Boolean
    
Public Image As Direct3DTexture8
Public Verticies(0 To 4) As MyScreen
    
Public Translucent As Boolean
Public BlackAlpha As Boolean
Public Dimension As ImgDimType
    
Public Padding As Long

