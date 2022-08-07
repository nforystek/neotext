Attribute VB_Name = "mod3DCaching"
Option Explicit

'Public Materials As New NTNodes10.Collection
Public Fields As New NTNodes10.Collection
Public Visions As New NTNodes10.Collection
Public Points As New NTNodes10.Collection

Public ObjectCount As Long
Public TriangleCount As Long
Public TriangleFace() As Single
't=triangle index in TriangleFace, VertexXAxis, VertexYAxis and VertexZAxis
'TriangleFace dimension (n,t) where n=0 is x of the face normal
'TriangleFace dimension (n,t) where n=1 is y of the face normal
'TriangleFace dimension (n,t) where n=2 is z of the face normal
'TriangleFace dimension (n,t) where n=3 flag for 1st arg of collision(), culling
'TriangleFace dimension (n,t) where n=4 is the object index
'TriangleFace dimension (n,t) where n=5 is the face index

Public VertexXAxis() As Single
Public VertexYAxis() As Single
Public VertexZAxis() As Single
't=triangle index in TriangleFace, VertexXAxis, VertexYAxis and VertexZAxis
'VertexXAxis dimension (n,t) where n=0 is X of the first vertex
'VertexXAxis dimension (n,t) where n=1 is X of the second vertex
'VertexXAxis dimension (n,t) where n=2 is X of the third vertex
'VertexYAxis dimension (n,t) where n=0 is Y of the first vertex
'VertexYAxis dimension (n,t) where n=1 is Y of the second vertex
'VertexYAxis dimension (n,t) where n=2 is Y of the third vertex
'VertexZAxis dimension (n,t) where n=0 is Z of the first vertex
'VertexZAxis dimension (n,t) where n=1 is Z of the second vertex
'VertexZAxis dimension (n,t) where n=2 is Z of the third vertex


Public Function PointCache(ByRef p As Point) As Long
    If Points.Count > 0 Then
        Dim i As Long
        For i = 1 To Points.Count
            If Points(i).Serialize = p.Serialize Then
                PointCache = i
                Set p = Points(i)
                Exit Function
            End If
        Next
    End If
    Points.Add p, p.Serialize
    PointCache = Points.Count
End Function
