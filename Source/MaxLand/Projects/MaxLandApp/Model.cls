VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Model"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Public Type MyMesh
'    Mesh As D3DXMesh
'
'    Materials() As D3DMATERIAL8
'    Textures() As Direct3DTexture8
'
'    Verticies() As D3DVERTEX
'    Indicies() As Integer
'    VBuffer As Direct3DVertexBuffer8
'
'    MaterialBuffer As D3DXBuffer
'    MaterialCount As Long
'    Filename As String
'End Type



Private pKey As String
Private pSerialize As Boolean

Public Property Get Serialize() As Boolean
    Serialize = True
End Property
Public Property Let Serialize(ByVal RHS As Boolean)
    pSerialize = RHS
End Property

Public Property Get Key() As String
    Key = pKey
End Property
Public Property Let Key(ByVal RHS As String)
    If RHS <> "" Then
        If pKey <> RHS And pKey <> "" Then If All.Exists(pKey) Then All.Remove pKey
        If All.Exists(RHS) Then
            All.Remove RHS
            All.Add Me, RHS
        End If
        pKey = RHS
    End If
End Property

Private Sub Class_Initialize()
    Key = "K" & ObjPtr(Me)
End Sub

