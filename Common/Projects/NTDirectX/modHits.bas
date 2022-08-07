Attribute VB_Name = "modHits"

Option Explicit

'############################################################################################################
'Derived Exports ############################################################################################
'############################################################################################################
                                    
Private Declare Function Collision Lib "MaxLandLib.dll" _
                                    (ByVal lngStreamFlagValue As _
                                    Long, ByVal lngTotalTriangles As _
                                    Long, sngTriangleFaceData() As _
                                    Single, sngVertexXAxisData() As _
                                    Single, sngVertexYAxisData() As _
                                    Single, sngVertexZAxisData() As _
                                    Single, ByVal lngTriangleToCheck As _
                                    Long, ByRef lngReturnHitObject As _
                                    Long, ByRef lngReturnHitTriangle As Long) As Boolean

'############################################################################################################
'Variable Declare ###########################################################################################
'############################################################################################################

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

Public VertexDirectX() As MyVertex







