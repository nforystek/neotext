VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bound"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'main typelibary that brings the basis
'for rendering/functions and objects
'that are build through this library
Private Origins As Stars
Private Scalars As Stars
Private Rotates As Stars

Public Spaces As Space



'Public Property Get FVF_VERTEX_SIZE() As Long
'    FVF_VERTEX_SIZE = 12
'End Property
'= 12
'Public Property Get FVF_RENDER_SIZE() As Long
'
'End Property
'= 32
'Public Const FVF_SCREEN = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX1
'Public Const FVF_RENDER = D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1

Public Function Serialize(ByRef Obj) As String
    Serialize = NTNodes10.Serialize(Obj)
End Function

Public Function Deserialize(ByVal txt As String)
    Set Deserialize = NTNodes10.Deserialize(txt)
End Function

Public Function NewVision(ByVal hWnd As Long) As Field
    Dim f As New Field
    If Not Visions.Exists("H" & hWnd) Then
        Set f.Vision = New Vision
        f.hWnd = hWnd
        Visions.Add f.Vision, "H" & hWnd
    Else
        Set f.Vision = Visions("H" & hWnd)
    End If
    Spaces.Add New Stars, "H" & hWnd
    Set f.Stars = Spaces.Current
    Fields.Add f
    Set NewVision = f
End Function

Private Function NewMatter(Optional ByRef Relative As Volume) As Matter
    If Not TypeName(Relative) = "Volume" Then
        Set Relative = New Volume
    End If
    Set NewMatter = New Matter
    Relative.Add NewMatter
    Set NewMatter.Shape1 = New Shape
    Set NewMatter.Shape1.Angle = New Angle
    Set NewMatter.Shape1.Angle.Interest = New Line
    Set NewMatter.Shape2 = New Shape
    Set NewMatter.Shape2.Angle = New Angle
    Set NewMatter.Shape2.Angle.Interest = New Line
    Set NewMatter.Shape2.Angle.Interest.Interest = New Point
    Set NewMatter.Shape2.Angle.Interest.Waypoint = New Point
    Set NewMatter.Shape2.Angle.Waypoint = New Point
    Set NewMatter.Shape2.Waypoint = New Point
End Function

Public Function NewSpirit(Optional ByRef Relative As Axis) As Spirit
    Dim p As New Spirit
    If Not TypeName(Relative) = "Axis" Then
        Set Relative = New Axis
    End If
    Set p.Axis = Relative
    Set NewSpirit = p
End Function

Private Function NewAxis(Optional ByRef Relative As Stars) As Axis
    If Not TypeName(Relative) = "Stars" Then
        Set Relative = Spaces.Current
    End If
    Dim p As Spirit
    Set p = NewSpirit()
    Set NewAxis = p.Axis
    Relative.Add p
End Function

Public Function NewOrbit(Optional ByRef Relative As Spirit) As Orbit
    Set NewOrbit = New Orbit
    If Not TypeName(Relative) = "Spirit" Then
        Set Relative = NewSpirit
    End If
    Set Relative.Axis.Orbit = NewOrbit
    Spaces.Current.Add Relative
End Function

'Public Sub Reference(ByRef Relative As Spirit)
'    Set AtSpirit = Relative
'    Dim p As Spirit
'    For Each p In Origins
'        With p.Point
'            .X = .X + AtSpirit.Point.X
'            .Y = .Y + AtSpirit.Point.Y
'            .Z = .Z + AtSpirit.Point.Z
'        End With
'    Next
'    AtSpirit.Point.X = 0
'    AtSpirit.Point.Y = 0
'    AtSpirit.Point.Z = 0
'End Sub

Public Sub Instance(ByRef Relative As Volume)
    NewOrbit
    Spaces.Current.Current.Axis.Orbit.Add Relative
End Sub

Private Function BuildArrays() As Long
    ReDim Preserve TriangleFace(0 To 5, 0 To TriangleCount) As Single
    ReDim Preserve VertexXAxis(0 To 2, 0 To TriangleCount) As Single
    ReDim Preserve VertexYAxis(0 To 2, 0 To TriangleCount) As Single
    ReDim Preserve VertexZAxis(0 To 2, 0 To TriangleCount) As Single
    BuildArrays = (((TriangleCount + 1) * 3) - 1)
'    ReDim Preserve VertexDirectX(0 To BuildArrays) As MyVertex
End Function

Public Function CreateVolumeFace(ByRef TextureFileName As String, ByRef TopRight As Point, ByRef BottomRight As Point, ByRef BottomLeft As Point, ByRef TopLeft As Point) As Volume
    If (((Not (TopRight.Equals(BottomRight) Or TopRight.Equals(BottomLeft) Or TopRight.Equals(TopLeft))) And _
        (Not (BottomLeft.Equals(BottomRight) Or BottomLeft.Equals(TopRight) Or BottomLeft.Equals(TopLeft))) And _
        (Not (BottomRight.Equals(TopRight) Or BottomRight.Equals(BottomLeft) Or BottomRight.Equals(TopLeft))) And _
        (Not (TopLeft.Equals(BottomRight) Or TopLeft.Equals(BottomLeft) Or TopLeft.Equals(TopRight)))) And _
        PathExists(TextureFileName, True)) Then

        With NewMatter(CreateVolumeFace)
            .TriangleIndex = BuildArrays

            .Index1 = PointCache(TopRight)
            .Index2 = PointCache(BottomRight)
            .Index3 = PointCache(BottomLeft)

            Set .Point1 = Points(.Index1)
            Set .Point2 = Points(.Index2)
            Set .Point3 = Points(.Index3)

            .v1 = 1
            .v2 = 1
            .U2 = 1
            .U3 = 1

            Set .Normal = TriangleNormal(.Point1, .Point2, .Point3)

            VertexXAxis(0, TriangleCount) = .Point1.X
            VertexXAxis(1, TriangleCount) = .Point2.X
            VertexXAxis(2, TriangleCount) = .Point3.X

            VertexYAxis(0, TriangleCount) = .Point1.Y
            VertexYAxis(1, TriangleCount) = .Point2.Y
            VertexYAxis(2, TriangleCount) = .Point3.Y

            VertexZAxis(0, TriangleCount) = .Point1.Z
            VertexZAxis(1, TriangleCount) = .Point2.Z
            VertexZAxis(2, TriangleCount) = .Point3.Z

            TriangleFace(0, TriangleCount) = .Normal.X
            TriangleFace(1, TriangleCount) = .Normal.Y
            TriangleFace(2, TriangleCount) = .Normal.Z
            TriangleFace(4, TriangleCount) = ObjectCount
            TriangleFace(5, TriangleCount) = 0

            .TriangleIndex = TriangleCount
            .ObjectIndex = ObjectCount
            .FaceIndex = 0

'            VertexDirectX(TriangleCount + 0).X = .Point1.X
'            VertexDirectX(TriangleCount + 0).Y = .Point1.Y
'            VertexDirectX(TriangleCount + 0).Z = .Point1.Z
'
'            VertexDirectX(TriangleCount + 1).X = .Point2.X
'            VertexDirectX(TriangleCount + 1).Y = .Point2.Y
'            VertexDirectX(TriangleCount + 1).Z = .Point2.Z
'
'            VertexDirectX(TriangleCount + 2).X = .Point3.X
'            VertexDirectX(TriangleCount + 2).Y = .Point3.Y
'            VertexDirectX(TriangleCount + 2).Z = .Point3.Z
'
'            VertexDirectX(TriangleCount + 0).NX = .Normal.X
'            VertexDirectX(TriangleCount + 0).NY = .Normal.Y
'            VertexDirectX(TriangleCount + 0).Nz = .Normal.Z
'
'            VertexDirectX(TriangleCount + 1).NX = .Normal.X
'            VertexDirectX(TriangleCount + 1).NY = .Normal.Y
'            VertexDirectX(TriangleCount + 1).Nz = .Normal.Z
'
'            VertexDirectX(TriangleCount + 2).NX = .Normal.X
'            VertexDirectX(TriangleCount + 2).NY = .Normal.Y
'            VertexDirectX(TriangleCount + 2).Nz = .Normal.Z
'
'            VertexDirectX(TriangleCount + 0).tv = .v1
'            VertexDirectX(TriangleCount + 1).tv = .v2
'            VertexDirectX(TriangleCount + 1).tu = .U2
'            VertexDirectX(TriangleCount + 2).tu = .U3

        End With
        TriangleCount = TriangleCount + 1
        
        With NewMatter(CreateVolumeFace)

            .TriangleIndex = BuildArrays

            .Index1 = PointCache(BottomLeft)
            .Index2 = PointCache(TopLeft)
            .Index3 = PointCache(TopRight)

            Set .Point1 = Points(.Index1)
            Set .Point2 = Points(.Index2)
            Set .Point3 = Points(.Index3)

            .v1 = 1
            .U2 = 1

            Set .Normal = TriangleNormal(.Point1, .Point2, .Point3)

            VertexXAxis(0, TriangleCount) = .Point1.X
            VertexXAxis(1, TriangleCount) = .Point2.X
            VertexXAxis(2, TriangleCount) = .Point3.X

            VertexYAxis(0, TriangleCount) = .Point1.Y
            VertexYAxis(1, TriangleCount) = .Point2.Y
            VertexYAxis(2, TriangleCount) = .Point3.Y

            VertexZAxis(0, TriangleCount) = .Point1.Z
            VertexZAxis(1, TriangleCount) = .Point2.Z
            VertexZAxis(2, TriangleCount) = .Point3.Z

            TriangleFace(0, TriangleCount) = .Normal.X
            TriangleFace(1, TriangleCount) = .Normal.Y
            TriangleFace(2, TriangleCount) = .Normal.Z
            TriangleFace(4, TriangleCount) = ObjectCount
            TriangleFace(5, TriangleCount) = 1

            .TriangleIndex = TriangleCount
            .ObjectIndex = ObjectCount
            .FaceIndex = 1

'            VertexDirectX(TriangleCount + 0).X = .Point1.X
'            VertexDirectX(TriangleCount + 0).Y = .Point1.Y
'            VertexDirectX(TriangleCount + 0).Z = .Point1.Z
'
'            VertexDirectX(TriangleCount + 1).X = .Point2.X
'            VertexDirectX(TriangleCount + 1).Y = .Point2.Y
'            VertexDirectX(TriangleCount + 1).Z = .Point2.Z
'
'            VertexDirectX(TriangleCount + 2).X = .Point3.X
'            VertexDirectX(TriangleCount + 2).Y = .Point3.Y
'            VertexDirectX(TriangleCount + 2).Z = .Point3.Z
'
'            VertexDirectX(TriangleCount + 0).NX = .Normal.X
'            VertexDirectX(TriangleCount + 0).NY = .Normal.Y
'            VertexDirectX(TriangleCount + 0).Nz = .Normal.Z
'
'            VertexDirectX(TriangleCount + 1).NX = .Normal.X
'            VertexDirectX(TriangleCount + 1).NY = .Normal.Y
'            VertexDirectX(TriangleCount + 1).Nz = .Normal.Z
'
'            VertexDirectX(TriangleCount + 2).NX = .Normal.X
'            VertexDirectX(TriangleCount + 2).NY = .Normal.Y
'            VertexDirectX(TriangleCount + 2).Nz = .Normal.Z
'
'            VertexDirectX(TriangleCount + 0).tv = .v1
'            VertexDirectX(TriangleCount + 1).tv = .U2
        End With
        TriangleCount = TriangleCount + 1

        ObjectCount = ObjectCount + 1
    End If
End Function

Public Function MakePoint(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As Point
    Set MakePoint = New Point
    MakePoint.X = X
    MakePoint.Y = Y
    MakePoint.Z = Z
    Set MakePoint = Points(PointCache(MakePoint))
End Function

Public Function CreateVolumeLanding(ByRef TextureFileName As String, ByVal OuterRadii As Single, ByVal RadiiSegments As Single, Optional ByVal InnerRadii As Single = 0) As Volume
    If OuterRadii <= 0 Then
        Err.Raise 8, , "OuterRadii must be above 0."
    ElseIf InnerRadii < 0 Then
        Err.Raise 8, , "InnerRadii must be 0 or above."
    ElseIf OuterRadii < InnerRadii Then
        Err.Raise 8, , "OuterRadii must be below InnerRadii."
    ElseIf RadiiSegments < 3 Then
        Err.Raise 8, , "RadiiSegments must be three or more"
    End If
    
    If PathExists(TextureFileName, True) Then
        Dim i As Long
        Dim g As Single
        Dim a As Double
        Dim l1 As Single
        Dim l2 As Single
        Dim l3 As Single
    
        Dim intX1 As Single
        Dim intX2 As Single
        Dim intX3 As Single
        Dim intX4 As Single
    
        Dim intY1 As Single
        Dim intY2 As Single
        Dim intY3 As Single
        Dim intY4 As Single
        Dim dist1 As Single
        Dim dist2 As Single
        Dim dist3 As Single
        Dim dist4 As Single
                    
        For i = -IIf(InnerRadii > 0, 6, 3) To ((IIf(InnerRadii > 0, 6, 3) * RadiiSegments) - 1) + (IIf(InnerRadii > 0, 6, 3) * 2) Step IIf(InnerRadii > 0, 6, 3)
    
            g = (((360 / RadiiSegments) * (((i + 1) / IIf(InnerRadii > 0, 6, 3)) - 1)) * (PI / 180))
    
            intX2 = (OuterRadii * Sin(g))
            intY2 = (-OuterRadii * Cos(g))
            intX3 = (InnerRadii * Sin(g))
            intY3 = (-InnerRadii * Cos(g))
    
            If i >= 0 Then
            
                If (InnerRadii > 0) Then
                    If (i Mod 12) = 0 Then
    
                        dist1 = Distance(intX2, 0, intY2, intX1, 0, intY1) * (1 / 100) * IIf(i Mod 4 = 0, 1, -Sin(g))
                        dist2 = Distance(intX1, 0, intY1, intX4, 0, intY4) * (1 / 100) * IIf(i Mod 4 = 0, 1, -Cos(g))
                        dist3 = Distance(intX3, 0, intY3, intX4, 0, intY4) * (1 / 100) * IIf(i Mod 4 = 0, 1, -Sin(g))
                        dist4 = Distance(intX3, 0, intY3, intX1, 0, intY1) * (1 / 100) * IIf(i Mod 4 = 0, 1, -Cos(g))
    
                    End If
       
                    With NewMatter(CreateVolumeLanding)
                        .TriangleIndex = BuildArrays
                        
                        .Index1 = PointCache(MakePoint(intX2, 0, intY2))
                        .Index2 = PointCache(MakePoint(intX1, 0, intY1))
                        .Index3 = PointCache(MakePoint(intX4, 0, intY4))
                        
                        Set .Point1 = Points(.Index1)
                        Set .Point2 = Points(.Index2)
                        Set .Point3 = Points(.Index3)
                        .v1 = dist2
                        .U2 = dist1
                        .v2 = dist4
                        .U3 = dist3
                        Set .Normal = TriangleNormal(.Point1, .Point2, .Point3)
    
                        l1 = Distance(.Point1.X, .Point1.Y, .Point1.Z, .Point2.X, .Point2.Y, .Point2.Z)
                        l2 = Distance(.Point2.X, .Point2.Y, .Point2.Z, .Point3.X, .Point3.Y, .Point3.Z)
                        l3 = Distance(.Point3.X, .Point3.Y, .Point3.Z, .Point1.X, .Point1.Y, .Point1.Z)
    
                        VertexXAxis(0, TriangleCount) = .Point1.X
                        VertexXAxis(1, TriangleCount) = .Point2.X
                        VertexXAxis(2, TriangleCount) = .Point3.X
                        
                        VertexYAxis(0, TriangleCount) = .Point1.Y
                        VertexYAxis(1, TriangleCount) = .Point2.Y
                        VertexYAxis(2, TriangleCount) = .Point3.Y
                        
                        VertexZAxis(0, TriangleCount) = .Point1.Z
                        VertexZAxis(1, TriangleCount) = .Point2.Z
                        VertexZAxis(2, TriangleCount) = .Point3.Z

                        TriangleFace(0, TriangleCount) = .Normal.X
                        TriangleFace(1, TriangleCount) = .Normal.Y
                        TriangleFace(2, TriangleCount) = .Normal.Z
                        TriangleFace(4, TriangleCount) = ObjectCount
                        TriangleFace(5, TriangleCount) = ((IIf(InnerRadii > 0, 6, 3) + i) \ IIf(InnerRadii > 0, 6, 3))

                        .TriangleIndex = TriangleCount
                        .ObjectIndex = ObjectCount
                        .FaceIndex = ((IIf(InnerRadii > 0, 6, 3) + i) \ IIf(InnerRadii > 0, 6, 3))
            
'                        VertexDirectX(TriangleCount + 0).X = .Point1.X
'                        VertexDirectX(TriangleCount + 0).Y = .Point1.Y
'                        VertexDirectX(TriangleCount + 0).Z = .Point1.Z
'
'                        VertexDirectX(TriangleCount + 1).X = .Point2.X
'                        VertexDirectX(TriangleCount + 1).Y = .Point2.Y
'                        VertexDirectX(TriangleCount + 1).Z = .Point2.Z
'
'                        VertexDirectX(TriangleCount + 2).X = .Point3.X
'                        VertexDirectX(TriangleCount + 2).Y = .Point3.Y
'                        VertexDirectX(TriangleCount + 2).Z = .Point3.Z
'
'                        VertexDirectX(TriangleCount + 0).NX = .Normal.X
'                        VertexDirectX(TriangleCount + 0).NY = .Normal.Y
'                        VertexDirectX(TriangleCount + 0).Nz = .Normal.Z
'
'                        VertexDirectX(TriangleCount + 1).NX = .Normal.X
'                        VertexDirectX(TriangleCount + 1).NY = .Normal.Y
'                        VertexDirectX(TriangleCount + 1).Nz = .Normal.Z
'
'                        VertexDirectX(TriangleCount + 2).NX = .Normal.X
'                        VertexDirectX(TriangleCount + 2).NY = .Normal.Y
'                        VertexDirectX(TriangleCount + 2).Nz = .Normal.Z
'
'                        VertexDirectX(TriangleCount + 0).tv = dist2
'                        VertexDirectX(TriangleCount + 1).tv = dist1
'                        VertexDirectX(TriangleCount + 1).tu = dist4
'                        VertexDirectX(TriangleCount + 2).tu = dist3
                    End With
                    
                    TriangleCount = TriangleCount + 1
        
                    With NewMatter(CreateVolumeLanding)
                        .TriangleIndex = BuildArrays
                        
                        .Index1 = PointCache(MakePoint(intX2, 0, intY2))
                        .Index2 = PointCache(MakePoint(intX4, 0, intY4))
                        .Index3 = PointCache(MakePoint(intX3, 0, intY3))
                        
                        Set .Point1 = Points(.Index1)
                        Set .Point2 = Points(.Index2)
                        Set .Point3 = Points(.Index3)
                        .v1 = dist2
                        .U2 = dist3
                        Set .Normal = TriangleNormal(.Point1, .Point2, .Point3)
    
                        l1 = Distance(.Point1.X, .Point1.Y, .Point1.Z, .Point2.X, .Point2.Y, .Point2.Z)
                        l2 = Distance(.Point2.X, .Point2.Y, .Point2.Z, .Point3.X, .Point3.Y, .Point3.Z)
                        l3 = Distance(.Point3.X, .Point3.Y, .Point3.Z, .Point1.X, .Point1.Y, .Point1.Z)
    
                        VertexXAxis(0, TriangleCount) = .Point1.X
                        VertexXAxis(1, TriangleCount) = .Point2.X
                        VertexXAxis(2, TriangleCount) = .Point3.X
                        
                        VertexYAxis(0, TriangleCount) = .Point1.Y
                        VertexYAxis(1, TriangleCount) = .Point2.Y
                        VertexYAxis(2, TriangleCount) = .Point3.Y
                        
                        VertexZAxis(0, TriangleCount) = .Point1.Z
                        VertexZAxis(1, TriangleCount) = .Point2.Z
                        VertexZAxis(2, TriangleCount) = .Point3.Z

                        TriangleFace(0, TriangleCount) = .Normal.X
                        TriangleFace(1, TriangleCount) = .Normal.Y
                        TriangleFace(2, TriangleCount) = .Normal.Z
                        TriangleFace(4, TriangleCount) = ObjectCount
                        TriangleFace(5, TriangleCount) = ((IIf(InnerRadii > 0, 6, 3) + i) \ IIf(InnerRadii > 0, 6, 3))

                        .TriangleIndex = TriangleCount
                        .ObjectIndex = ObjectCount
                        .FaceIndex = ((IIf(InnerRadii > 0, 6, 3) + i) \ IIf(InnerRadii > 0, 6, 3))

'                        VertexDirectX(TriangleCount + 0).X = .Point1.X
'                        VertexDirectX(TriangleCount + 0).Y = .Point1.Y
'                        VertexDirectX(TriangleCount + 0).Z = .Point1.Z
'
'                        VertexDirectX(TriangleCount + 1).X = .Point2.X
'                        VertexDirectX(TriangleCount + 1).Y = .Point2.Y
'                        VertexDirectX(TriangleCount + 1).Z = .Point2.Z
'
'                        VertexDirectX(TriangleCount + 2).X = .Point3.X
'                        VertexDirectX(TriangleCount + 2).Y = .Point3.Y
'                        VertexDirectX(TriangleCount + 2).Z = .Point3.Z
'
'                        VertexDirectX(TriangleCount + 0).NX = .Normal.X
'                        VertexDirectX(TriangleCount + 0).NY = .Normal.Y
'                        VertexDirectX(TriangleCount + 0).Nz = .Normal.Z
'
'                        VertexDirectX(TriangleCount + 1).NX = .Normal.X
'                        VertexDirectX(TriangleCount + 1).NY = .Normal.Y
'                        VertexDirectX(TriangleCount + 1).Nz = .Normal.Z
'
'                        VertexDirectX(TriangleCount + 2).NX = .Normal.X
'                        VertexDirectX(TriangleCount + 2).NY = .Normal.Y
'                        VertexDirectX(TriangleCount + 2).Nz = .Normal.Z
'
'                        VertexDirectX(TriangleCount + 0).tv = dist2
'                        VertexDirectX(TriangleCount + 1).tv = dist3
                    End With
                    
                    TriangleCount = TriangleCount + 1
                Else
    
                    dist1 = Distance(intX2, 0, intY2, intX1, 0, intY1) * IIf(i Mod 4 = 0, 1, Sin(g))
                    dist2 = Distance(intX1, 0, intY1, intX4, 0, intY4) * IIf(i Mod 4 = 0, 1, Cos(g))

                    With NewMatter(CreateVolumeLanding)
                        .TriangleIndex = BuildArrays
                        
                        .Index1 = PointCache(MakePoint(intX2, 0, intY2))
                        .Index2 = PointCache(MakePoint(intX1, 0, intY1))
                        .Index3 = PointCache(MakePoint(intX4, 0, intY4))
                        
                        Set .Point1 = Points(.Index1)
                        Set .Point2 = Points(.Index2)
                        Set .Point3 = Points(.Index3)
                        .U1 = ((1 / dist1) * (dist1 / 100))
                        .v2 = ((1 / dist2) * (dist2 / 100))
                        Set .Normal = TriangleNormal(.Point1, .Point2, .Point3)
    
                        l1 = Distance(.Point1.X, .Point1.Y, .Point1.Z, .Point2.X, .Point2.Y, .Point2.Z)
                        l2 = Distance(.Point2.X, .Point2.Y, .Point2.Z, .Point3.X, .Point3.Y, .Point3.Z)
                        l3 = Distance(.Point3.X, .Point3.Y, .Point3.Z, .Point1.X, .Point1.Y, .Point1.Z)
    
                        VertexXAxis(0, TriangleCount) = .Point1.X
                        VertexXAxis(1, TriangleCount) = .Point2.X
                        VertexXAxis(2, TriangleCount) = .Point3.X
                        
                        VertexYAxis(0, TriangleCount) = .Point1.Y
                        VertexYAxis(1, TriangleCount) = .Point2.Y
                        VertexYAxis(2, TriangleCount) = .Point3.Y
                        
                        VertexZAxis(0, TriangleCount) = .Point1.Z
                        VertexZAxis(1, TriangleCount) = .Point2.Z
                        VertexZAxis(2, TriangleCount) = .Point3.Z

                        TriangleFace(0, TriangleCount) = .Normal.X
                        TriangleFace(1, TriangleCount) = .Normal.Y
                        TriangleFace(2, TriangleCount) = .Normal.Z
                        TriangleFace(4, TriangleCount) = ObjectCount
                        TriangleFace(5, TriangleCount) = ((IIf(InnerRadii > 0, 6, 3) + i) \ IIf(InnerRadii > 0, 6, 3))

                        .TriangleIndex = TriangleCount
                        .ObjectIndex = ObjectCount
                        .FaceIndex = ((IIf(InnerRadii > 0, 6, 3) + i) \ IIf(InnerRadii > 0, 6, 3))
            
'                        VertexDirectX(TriangleCount + 0).X = .Point1.X
'                        VertexDirectX(TriangleCount + 0).Y = .Point1.Y
'                        VertexDirectX(TriangleCount + 0).Z = .Point1.Z
'
'                        VertexDirectX(TriangleCount + 1).X = .Point2.X
'                        VertexDirectX(TriangleCount + 1).Y = .Point2.Y
'                        VertexDirectX(TriangleCount + 1).Z = .Point2.Z
'
'                        VertexDirectX(TriangleCount + 2).X = .Point3.X
'                        VertexDirectX(TriangleCount + 2).Y = .Point3.Y
'                        VertexDirectX(TriangleCount + 2).Z = .Point3.Z
'
'                        VertexDirectX(TriangleCount + 0).NX = .Normal.X
'                        VertexDirectX(TriangleCount + 0).NY = .Normal.Y
'                        VertexDirectX(TriangleCount + 0).Nz = .Normal.Z
'
'                        VertexDirectX(TriangleCount + 1).NX = .Normal.X
'                        VertexDirectX(TriangleCount + 1).NY = .Normal.Y
'                        VertexDirectX(TriangleCount + 1).Nz = .Normal.Z
'
'                        VertexDirectX(TriangleCount + 2).NX = .Normal.X
'                        VertexDirectX(TriangleCount + 2).NY = .Normal.Y
'                        VertexDirectX(TriangleCount + 2).Nz = .Normal.Z
'
'                        VertexDirectX(TriangleCount + 0).tv = ((1 / dist1) * (dist1 / 100))
'                        VertexDirectX(TriangleCount + 1).tv = ((1 / dist2) * (dist2 / 100))
                    End With
                    
                    TriangleCount = TriangleCount + 1
                End If
    
            End If
    
            intX1 = intX2
            intY1 = intY2
            intX4 = intX3
            intY4 = intY3
    
        Next
        
        ObjectCount = ObjectCount + 1
    End If
    
End Function

'Public Function CreateVolumeMesh(ByVal DirectXFileName As String) As Volume
'    If PathExists(DirectXFileName, True) Then
'
'        Dim MeshVerticies() As D3DVERTEX
'        Dim MeshIndicies() As Integer
'
'        Dim Mesh As D3DXMesh
'        Set Mesh = D3DX.LoadMeshFromX(DirectXFileName, D3DXMESH_DYNAMIC, DDevice, Nothing, Nothing, 0&)
'
'        Dim VD As D3DVERTEXBUFFER_DESC
'        Mesh.GetVertexBuffer.GetDesc VD
'        ReDim MeshVerticies(0 To 0) As D3DVERTEX
'        ReDim MeshVerticies(0 To ((VD.Size \ Len(MeshVerticies(0))) - 1)) As D3DVERTEX
'        D3DVertexBuffer8GetData Mesh.GetVertexBuffer, 0, VD.Size, 0, MeshVerticies(0)
'
'        Dim ID As D3DINDEXBUFFER_DESC
'        Mesh.GetIndexBuffer.GetDesc ID
'        ReDim MeshIndicies(0 To 0) As Integer
'        ReDim MeshIndicies(0 To ((ID.Size \ Len(MeshIndicies(0))) - 1)) As Integer
'        D3DIndexBuffer8GetData Mesh.GetIndexBuffer, 0, ID.Size, 0, MeshIndicies(0)
'
'        Set CreateVolumeMesh = New Volume
'
'        Dim Verts() As D3DVERTEX
'        Dim chk1 As Integer
'        Dim chk2 As Integer
'        Dim chk3 As Integer
'        Dim chk4 As Integer
'        Dim chk5 As Integer
'        Dim chk6 As Integer
'        chk1 = -1
'        Dim Index As Long
'
'        Const TrianglePerFace As Long = 2
'
'        For Index = 0 To UBound(MeshIndicies)
'            Debug.Print MeshIndicies(Index);
'        Next
'        Debug.Print
'
'        Index = 0
'
'        Do While Index <= UBound(MeshIndicies)
'
'            With NewMatter(CreateVolumeMesh)
'                .TriangleIndex = BuildArrays
'
'                .Index1 = PointCache(MakePoint( _
'                    MeshVerticies(MeshIndicies(Index + 0)).X, _
'                    MeshVerticies(MeshIndicies(Index + 0)).Y, _
'                    MeshVerticies(MeshIndicies(Index + 0)).Z))
'
'                .Index2 = PointCache(MakePoint( _
'                    MeshVerticies(MeshIndicies(Index + 1)).X, _
'                    MeshVerticies(MeshIndicies(Index + 1)).Y, _
'                    MeshVerticies(MeshIndicies(Index + 1)).Z))
'
'                .Index3 = PointCache(MakePoint( _
'                    MeshVerticies(MeshIndicies(Index + 2)).X, _
'                    MeshVerticies(MeshIndicies(Index + 2)).Y, _
'                    MeshVerticies(MeshIndicies(Index + 2)).Z))
'
'                Set .Point1 = Points(.Index1)
'                Set .Point2 = Points(.Index2)
'                Set .Point3 = Points(.Index3)
'
'                Set .Normal = TriangleNormal(.Point1, .Point2, .Point3)
'
'                VertexXAxis(0, TriangleCount) = .Point1.X
'                VertexXAxis(1, TriangleCount) = .Point2.X
'                VertexXAxis(2, TriangleCount) = .Point3.X
'
'                VertexYAxis(0, TriangleCount) = .Point1.Y
'                VertexYAxis(1, TriangleCount) = .Point2.Y
'                VertexYAxis(2, TriangleCount) = .Point3.Y
'
'                VertexZAxis(0, TriangleCount) = .Point1.Z
'                VertexZAxis(1, TriangleCount) = .Point2.Z
'                VertexZAxis(2, TriangleCount) = .Point3.Z
'
'                TriangleFace(0, TriangleCount) = .Normal.X
'                TriangleFace(1, TriangleCount) = .Normal.Y
'                TriangleFace(2, TriangleCount) = .Normal.Z
'                TriangleFace(4, TriangleCount) = ObjectCount
'                TriangleFace(5, TriangleCount) = (Index \ 4) + 0
'
'                .TriangleIndex = TriangleCount
'                .ObjectIndex = ObjectCount
'                .FaceIndex = (Index \ 4) + 0
'
''                VertexDirectX(TriangleCount + 0).X = .Point1.X
''                VertexDirectX(TriangleCount + 0).Y = .Point1.Y
''                VertexDirectX(TriangleCount + 0).Z = .Point1.Z
''                VertexDirectX(TriangleCount + 0).NX = .Normal.X
''                VertexDirectX(TriangleCount + 0).NY = .Normal.Y
''                VertexDirectX(TriangleCount + 0).Nz = .Normal.Z
''                VertexDirectX(TriangleCount + 0).tu = MeshVerticies(MeshIndicies(Index + 0)).tu
''                VertexDirectX(TriangleCount + 0).tv = MeshVerticies(MeshIndicies(Index + 0)).tv
''
''                VertexDirectX(TriangleCount + 1).X = .Point2.X
''                VertexDirectX(TriangleCount + 1).Y = .Point2.Y
''                VertexDirectX(TriangleCount + 1).Z = .Point2.Z
''                VertexDirectX(TriangleCount + 1).NX = .Normal.X
''                VertexDirectX(TriangleCount + 1).NY = .Normal.Y
''                VertexDirectX(TriangleCount + 1).Nz = .Normal.Z
''                VertexDirectX(TriangleCount + 1).tu = MeshVerticies(MeshIndicies(Index + 1)).tu
''                VertexDirectX(TriangleCount + 1).tv = MeshVerticies(MeshIndicies(Index + 1)).tv
''
''                VertexDirectX(TriangleCount + 2).X = .Point3.X
''                VertexDirectX(TriangleCount + 2).Y = .Point3.Y
''                VertexDirectX(TriangleCount + 2).Z = .Point3.Z
''                VertexDirectX(TriangleCount + 2).NX = .Normal.X
''                VertexDirectX(TriangleCount + 2).NY = .Normal.Y
''                VertexDirectX(TriangleCount + 2).Nz = .Normal.Z
''                VertexDirectX(TriangleCount + 2).tu = MeshVerticies(MeshIndicies(Index + 2)).tu
''                VertexDirectX(TriangleCount + 2).tv = MeshVerticies(MeshIndicies(Index + 2)).tv
'
'            End With
'
'            TriangleCount = TriangleCount + 1
'
'            With NewMatter(CreateVolumeMesh)
'                .TriangleIndex = BuildArrays
'
'                .Index1 = PointCache(MakePoint( _
'                    MeshVerticies(MeshIndicies(Index + 3)).X, _
'                    MeshVerticies(MeshIndicies(Index + 3)).Y, _
'                    MeshVerticies(MeshIndicies(Index + 3)).Z))
'
'                .Index2 = PointCache(MakePoint( _
'                    MeshVerticies(MeshIndicies(Index + 4)).X, _
'                    MeshVerticies(MeshIndicies(Index + 4)).Y, _
'                    MeshVerticies(MeshIndicies(Index + 4)).Z))
'
'                .Index3 = PointCache(MakePoint( _
'                    MeshVerticies(MeshIndicies(Index + 5)).X, _
'                    MeshVerticies(MeshIndicies(Index + 5)).Y, _
'                    MeshVerticies(MeshIndicies(Index + 5)).Z))
'
'                Set .Point1 = Points(.Index1)
'                Set .Point2 = Points(.Index2)
'                Set .Point3 = Points(.Index3)
'
'                Set .Normal = TriangleNormal(.Point1, .Point2, .Point3)
'
'                VertexXAxis(0, TriangleCount) = .Point1.X
'                VertexXAxis(1, TriangleCount) = .Point2.X
'                VertexXAxis(2, TriangleCount) = .Point3.X
'
'                VertexYAxis(0, TriangleCount) = .Point1.Y
'                VertexYAxis(1, TriangleCount) = .Point2.Y
'                VertexYAxis(2, TriangleCount) = .Point3.Y
'
'                VertexZAxis(0, TriangleCount) = .Point1.Z
'                VertexZAxis(1, TriangleCount) = .Point2.Z
'                VertexZAxis(2, TriangleCount) = .Point3.Z
'
'                TriangleFace(0, TriangleCount) = .Normal.X
'                TriangleFace(1, TriangleCount) = .Normal.Y
'                TriangleFace(2, TriangleCount) = .Normal.Z
'                TriangleFace(4, TriangleCount) = ObjectCount
'                TriangleFace(5, TriangleCount) = (Index \ 4) + 1
'
'                .TriangleIndex = TriangleCount
'                .ObjectIndex = ObjectCount
'                .FaceIndex = (Index \ 4) + 1
'
''                VertexDirectX(TriangleCount + 0).X = .Point1.X
''                VertexDirectX(TriangleCount + 0).Y = .Point1.Y
''                VertexDirectX(TriangleCount + 0).Z = .Point1.Z
''                VertexDirectX(TriangleCount + 0).NX = .Normal.X
''                VertexDirectX(TriangleCount + 0).NY = .Normal.Y
''                VertexDirectX(TriangleCount + 0).Nz = .Normal.Z
''                VertexDirectX(TriangleCount + 0).tu = MeshVerticies(MeshIndicies(Index + 0)).tu
''                VertexDirectX(TriangleCount + 0).tv = MeshVerticies(MeshIndicies(Index + 0)).tv
''
''                VertexDirectX(TriangleCount + 1).X = .Point2.X
''                VertexDirectX(TriangleCount + 1).Y = .Point2.Y
''                VertexDirectX(TriangleCount + 1).Z = .Point2.Z
''                VertexDirectX(TriangleCount + 1).NX = .Normal.X
''                VertexDirectX(TriangleCount + 1).NY = .Normal.Y
''                VertexDirectX(TriangleCount + 1).Nz = .Normal.Z
''                VertexDirectX(TriangleCount + 1).tu = MeshVerticies(MeshIndicies(Index + 1)).tu
''                VertexDirectX(TriangleCount + 1).tv = MeshVerticies(MeshIndicies(Index + 1)).tv
''
''                VertexDirectX(TriangleCount + 2).X = .Point3.X
''                VertexDirectX(TriangleCount + 2).Y = .Point3.Y
''                VertexDirectX(TriangleCount + 2).Z = .Point3.Z
''                VertexDirectX(TriangleCount + 2).NX = .Normal.X
''                VertexDirectX(TriangleCount + 2).NY = .Normal.Y
''                VertexDirectX(TriangleCount + 2).Nz = .Normal.Z
''                VertexDirectX(TriangleCount + 2).tu = MeshVerticies(MeshIndicies(Index + 2)).tu
''                VertexDirectX(TriangleCount + 2).tv = MeshVerticies(MeshIndicies(Index + 2)).tv
'
'            End With
'
'            TriangleCount = TriangleCount + 1
'            Index = Index + 6
'        Loop
'
'        ObjectCount = ObjectCount + 1
'
'    End If
'End Function
'
'Public Function CreateVolumeCube(ByRef TextureFileName As String, ByRef LengthDimension As Single, ByRef WidthDimension As Single, ByRef HeightDimension As Single) As Volume
'    If (((LengthDimension > 0) And (WidthDimension > 0) And (HeightDimension > 0)) And PathExists(TextureFileName, True)) Then
'
'    End If
'End Function
'

'Public Sub DuplicateSpace(ByRef Space As Space)
'    Set AtSpace.Forth = AtSpace.Clone
'    ForthSpace
'End Sub
'
'Public Sub DuplicateOrbit(ByRef Orbit As Orbit)
'    Set AtSpace.Forth = New Space
'    ForthSpace
'    Dim p As New Spirit
'    Set p.Axis = New Axis
'    Set p.Axis.Orbit = Orbit
'    Set AtSpace.Forth.Stars = New Stars
'    Set AtSpace.Forth.Stars.Spirit = New Spirit
'    Set AtSpirit = AtSpace.Forth.Stars.Spirit
'    Set AtSpace = AtSpace.Forth
'End Sub
'
'Public Sub DuplicateVolume(ByRef Volume As Volume)
'    Dim o As New Orbit
'    Dim p As New Spirit
'    Set p.Axis = New Axis
'    Set p.Axis.Orbit = o
'    Set o.Volume = Volume
'End Sub

Public Function Localize(Optional ByRef Center As Axis)
    Dim s As Stars
    Dim p As Spirit
    For Each s In Spaces
        For Each p In s
            With p.Axis.Range
                If TypeName(Center) = "Axis" Then
                    .X = (.X - (Center.Range.X + Spaces.Current.Current.Axis.Range.X))
                    .Y = (.Y - (Center.Range.Y + Spaces.Current.Current.Axis.Range.Y))
                    .Z = (.Z - (Center.Range.Z + Spaces.Current.Current.Axis.Range.Z))
                Else
                    .X = (.X - Spaces.Current.Current.Axis.Range.X)
                    .Y = (.Y - Spaces.Current.Current.Axis.Range.Y)
                    .Z = (.Z - Spaces.Current.Current.Axis.Range.Z)
                End If
            End With
        Next
    Next

End Function

Public Sub Location(ByRef Origin As Point)
    Dim p As Spirit
    Set p = NewSpirit(Spaces.Current.Current.Axis)
    With p.Point

        .X = Origin.X
        .Y = Origin.Y
        .Z = Origin.Z

    End With
    Origins.Add p
End Sub

Public Sub Position(ByRef Origin As Point)
    Dim v As Volume
    Dim m As Matter
    Dim vin As D3DVECTOR
    Dim vout As D3DVECTOR
    Dim matMesh2 As D3DMATRIX
    Dim matMesh As D3DMATRIX
    
    For Each v In Spaces.Current.Current.Axis.Orbit
        For Each m In v
            With m

               D3DXMatrixIdentity matMesh
               D3DXMatrixIdentity matMesh2

               D3DXMatrixTranslation matMesh2, Origin.X, Origin.Y, Origin.Z
               D3DXMatrixMultiply matMesh, matMesh, matMesh2

               vin.X = .Point1.X
               vin.Y = .Point1.Y
               vin.Z = .Point1.Z
               D3DXVec3TransformCoord vout, vin, matMesh
                .Index1 = PointCache(MakePoint(vout.X, vout.Y, vout.Z))
                Set .Point1 = Points(.Index1)
'               .Point1.X = vout.X
'               .Point1.Y = vout.Y
'               .Point1.Z = vout.Z

               vin.X = .Point2.X
               vin.Y = .Point2.Y
               vin.Z = .Point2.Z
               D3DXVec3TransformCoord vout, vin, matMesh
                .Index2 = PointCache(MakePoint(vout.X, vout.Y, vout.Z))
                Set .Point2 = Points(.Index1)
'               .Point2.X = vout.X
'               .Point2.Y = vout.Y
'               .Point2.Z = vout.Z
                    
            End With

        Next
    Next

End Sub

Public Sub Rotation(ByRef Radians As Point)
    Dim p As Spirit
    Set p = NewSpirit(Spaces.Current.Current.Axis)
    With p.Point
        
        .X = Radians.X
        .Y = Radians.Y
        .Z = Radians.Z

    End With
    Rotates.Add p
End Sub

Public Sub Orientate(ByRef Radians As Point)
    Dim v As Volume
    Dim m As Matter
    Dim vin As D3DVECTOR
    Dim vout As D3DVECTOR
    Dim matMesh2 As D3DMATRIX
    Dim matMesh As D3DMATRIX
    
    For Each v In Spaces.Current.Current.Axis.Orbit
        For Each m In v
            With m

                D3DXMatrixIdentity matMesh
                D3DXMatrixIdentity matMesh2

                D3DXMatrixRotationX matMesh2, Radians.X
                D3DXMatrixMultiply matMesh, matMesh2, matMesh

                D3DXMatrixRotationY matMesh2, Radians.Y
                D3DXMatrixMultiply matMesh, matMesh2, matMesh

                D3DXMatrixRotationZ matMesh2, Radians.Z
                D3DXMatrixMultiply matMesh, matMesh2, matMesh

                vin.X = .Point1.X
                vin.Y = .Point1.Y
                vin.Z = .Point1.Z
                D3DXVec3TransformCoord vout, vin, matMesh
                .Index1 = PointCache(MakePoint(vout.X, vout.Y, vout.Z))
                Set .Point1 = Points(.Index1)
'                .Point1.X = vout.X
'                .Point1.Y = vout.Y
'                .Point1.Z = vout.Z

                vin.X = .Point2.X
                vin.Y = .Point2.Y
                vin.Z = .Point2.Z
                D3DXVec3TransformCoord vout, vin, matMesh
                .Index2 = PointCache(MakePoint(vout.X, vout.Y, vout.Z))
                Set .Point2 = Points(.Index1)
'                .Point2.X = vout.X
'                .Point2.Y = vout.Y
'                .Point2.Z = vout.Z

            End With

        Next
    Next

End Sub

Public Sub Explode(ByVal Ratios As Point)
    Dim v As Volume
    Dim m As Matter
    Dim vin As D3DVECTOR
    Dim vout As D3DVECTOR
    Dim matMesh2 As D3DMATRIX
    Dim matMesh As D3DMATRIX
    
    For Each v In Spaces.Current.Current.Axis.Orbit
        For Each m In v
            With m

               D3DXMatrixIdentity matMesh
               D3DXMatrixIdentity matMesh2

               D3DXMatrixScaling matMesh2, Ratios.X, Ratios.Y, Ratios.Z
               D3DXMatrixMultiply matMesh, matMesh, matMesh2

               vin.X = .Point1.X
               vin.Y = .Point1.Y
               vin.Z = .Point1.Z
               D3DXVec3TransformCoord vout, vin, matMesh
                .Index1 = PointCache(MakePoint(vout.X, vout.Y, vout.Z))
                Set .Point1 = Points(.Index1)
'               .Point1.X = vout.X
'               .Point1.Y = vout.Y
'               .Point1.Z = vout.Z

               vin.X = .Point2.X
               vin.Y = .Point2.Y
               vin.Z = .Point2.Z
               D3DXVec3TransformCoord vout, vin, matMesh
                .Index2 = PointCache(MakePoint(vout.X, vout.Y, vout.Z))
                Set .Point2 = Points(.Index1)
'               .Point2.X = vout.X
'               .Point2.Y = vout.Y
'               .Point2.Z = vout.Z
                    
            End With

        Next
    Next

End Sub

Public Sub Scaling(ByVal Ratios As Point)
    Dim p As Spirit
    Set p = NewSpirit(Spaces.Current.Current.Axis)
    With p.Point

        .X = Ratios.X
        .Y = Ratios.Y
        .Z = Ratios.Z
        
    End With
    Scalars.Add p
End Sub

Public Sub Begin()

    Dim p As Spirit
    Dim vin As D3DVECTOR
    Dim vout As D3DVECTOR

    Dim matMesh2 As D3DMATRIX
    Dim matMesh As D3DMATRIX

    Dim v As Volume
    Dim m As Matter

    For Each p In Origins
        For Each v In p.Axis.Orbit
            For Each m In v
                With m
    
                   D3DXMatrixIdentity matMesh
                   D3DXMatrixIdentity matMesh2
    
                   D3DXMatrixTranslation matMesh2, -p.Point.X, -p.Point.Y, -p.Point.Z
                   D3DXMatrixMultiply matMesh, matMesh, matMesh2
    
                   vin.X = .Point1.X
                   vin.Y = .Point1.Y
                   vin.Z = .Point1.Z
                   D3DXVec3TransformCoord vout, vin, matMesh
                   .Index1 = PointCache(MakePoint(vout.X, vout.Y, vout.Z))
                   Set .Point1 = Points(.Index1)
'                   .Point1.X = vout.X
'                   .Point1.Y = vout.Y
'                   .Point1.Z = vout.Z
    
                   vin.X = .Point2.X
                   vin.Y = .Point2.Y
                   vin.Z = .Point2.Z
                   D3DXVec3TransformCoord vout, vin, matMesh
                   .Index2 = PointCache(MakePoint(vout.X, vout.Y, vout.Z))
                   Set .Point2 = Points(.Index1)
'                   .Point2.X = vout.X
'                   .Point2.Y = vout.Y
'                   .Point2.Z = vout.Z

                End With
            Next
        Next
    Next
    Origins.Clear
    
    For Each p In Rotates
        For Each v In p.Axis.Orbit
            For Each m In v
                With m

                    D3DXMatrixIdentity matMesh
                    D3DXMatrixIdentity matMesh2

                    D3DXMatrixRotationX matMesh2, -p.Point.X
                    D3DXMatrixMultiply matMesh, matMesh2, matMesh

                    D3DXMatrixRotationY matMesh2, -p.Point.Y
                    D3DXMatrixMultiply matMesh, matMesh2, matMesh

                    D3DXMatrixRotationZ matMesh2, -p.Point.Z
                    D3DXMatrixMultiply matMesh, matMesh2, matMesh

                    vin.X = .Point1.X
                    vin.Y = .Point1.Y
                    vin.Z = .Point1.Z
                    D3DXVec3TransformCoord vout, vin, matMesh
                   .Index1 = PointCache(MakePoint(vout.X, vout.Y, vout.Z))
                   Set .Point1 = Points(.Index1)
'                    .Point1.X = vout.X
'                    .Point1.Y = vout.Y
'                    .Point1.Z = vout.Z

                    vin.X = .Point2.X
                    vin.Y = .Point2.Y
                    vin.Z = .Point2.Z
                    D3DXVec3TransformCoord vout, vin, matMesh
                   .Index2 = PointCache(MakePoint(vout.X, vout.Y, vout.Z))
                   Set .Point2 = Points(.Index1)
'                    .Point2.X = vout.X
'                    .Point2.Y = vout.Y
'                    .Point2.Z = vout.Z

                End With

            Next
        Next
    Next
    Rotates.Clear

    For Each p In Scalars
        For Each v In p.Axis.Orbit
            For Each m In v
                With m

                    D3DXMatrixIdentity matMesh
                    D3DXMatrixIdentity matMesh2

                    D3DXMatrixScaling matMesh2, ((1 / p.Point.X) * p.Point.X), ((1 / p.Point.Y) * p.Point.Y), ((1 / p.Point.Z) * p.Point.Z)
                    D3DXMatrixMultiply matMesh, matMesh, matMesh2
     
                    vin.X = .Point1.X
                    vin.Y = .Point1.Y
                    vin.Z = .Point1.Z
                    D3DXVec3TransformCoord vout, vin, matMesh
                    .Index1 = PointCache(MakePoint(vout.X, vout.Y, vout.Z))
                    Set .Point1 = Points(.Index1)
'                    .Point1.X = vout.X
'                    .Point1.Y = vout.Y
'                    .Point1.Z = vout.Z

                    vin.X = .Point2.X
                    vin.Y = .Point2.Y
                    vin.Z = .Point2.Z
                    D3DXVec3TransformCoord vout, vin, matMesh
                    .Index2 = PointCache(MakePoint(vout.X, vout.Y, vout.Z))
                    Set .Point2 = Points(.Index1)
'                    .Point2.X = vout.X
'                    .Point2.Y = vout.Y
'                    .Point2.Z = vout.Z

                End With
            Next
        Next
    Next
    Scalars.Clear

End Sub

Public Sub Finish()

    Dim f As Field
    Dim s As Spirit
    Dim v As Volume
    Dim m As Matter

    Dim p As Spirit

    Dim vin As D3DVECTOR
    Dim vout As D3DVECTOR

    Dim matMesh2 As D3DMATRIX
    Dim matMesh As D3DMATRIX

    For Each p In Origins
        For Each v In p.Axis.Orbit
            For Each m In v
                With m

                   D3DXMatrixIdentity matMesh
                   D3DXMatrixIdentity matMesh2
    
                   D3DXMatrixTranslation matMesh2, p.Point.X, p.Point.Y, p.Point.Z
                   D3DXMatrixMultiply matMesh, matMesh, matMesh2
    
                   vin.X = .Point1.X
                   vin.Y = .Point1.Y
                   vin.Z = .Point1.Z
                   D3DXVec3TransformCoord vout, vin, matMesh
                   .Index1 = PointCache(MakePoint(vout.X, vout.Y, vout.Z))
                   Set .Point1 = Points(.Index1)
'                   .Point1.X = vout.X
'                   .Point1.Y = vout.Y
'                   .Point1.Z = vout.Z
    
                   vin.X = .Point2.X
                   vin.Y = .Point2.Y
                   vin.Z = .Point2.Z
                   D3DXVec3TransformCoord vout, vin, matMesh
                   .Index2 = PointCache(MakePoint(vout.X, vout.Y, vout.Z))
                   Set .Point2 = Points(.Index1)
'                   .Point2.X = vout.X
'                   .Point2.Y = vout.Y
'                   .Point2.Z = vout.Z
                        
                End With
            Next
        Next
    Next
    
    For Each p In Rotates
        For Each v In p.Axis.Orbit
            For Each m In v
                With m

                    D3DXMatrixIdentity matMesh
                    D3DXMatrixIdentity matMesh2

                    D3DXMatrixRotationX matMesh2, p.Point.X
                    D3DXMatrixMultiply matMesh, matMesh2, matMesh

                    D3DXMatrixRotationY matMesh2, p.Point.Y
                    D3DXMatrixMultiply matMesh, matMesh2, matMesh

                    D3DXMatrixRotationZ matMesh2, p.Point.Z
                    D3DXMatrixMultiply matMesh, matMesh2, matMesh

                    vin.X = .Point1.X
                    vin.Y = .Point1.Y
                    vin.Z = .Point1.Z
                    D3DXVec3TransformCoord vout, vin, matMesh
                    .Index1 = PointCache(MakePoint(vout.X, vout.Y, vout.Z))
                    Set .Point1 = Points(.Index1)
'                    .Point1.X = vout.X
'                    .Point1.Y = vout.Y
'                    .Point1.Z = vout.Z

                    vin.X = .Point2.X
                    vin.Y = .Point2.Y
                    vin.Z = .Point2.Z
                    D3DXVec3TransformCoord vout, vin, matMesh
                    .Index2 = PointCache(MakePoint(vout.X, vout.Y, vout.Z))
                    Set .Point2 = Points(.Index1)
'                    .Point2.X = vout.X
'                    .Point2.Y = vout.Y
'                    .Point2.Z = vout.Z

                End With

            Next
        Next
    Next

    For Each p In Scalars
        For Each v In p.Axis.Orbit
            For Each m In v
                With m

                    D3DXMatrixIdentity matMesh
                    D3DXMatrixIdentity matMesh2

                    D3DXMatrixScaling matMesh2, p.Point.X, p.Point.Y, p.Point.Z
                    D3DXMatrixMultiply matMesh, matMesh, matMesh2
     
                    vin.X = .Point1.X
                    vin.Y = .Point1.Y
                    vin.Z = .Point1.Z
                    D3DXVec3TransformCoord vout, vin, matMesh
                    .Index1 = PointCache(MakePoint(vout.X, vout.Y, vout.Z))
                    Set .Point1 = Points(.Index1)
'                    .Point1.X = vout.X
'                    .Point1.Y = vout.Y
'                    .Point1.Z = vout.Z

                    vin.X = .Point2.X
                    vin.Y = .Point2.Y
                    vin.Z = .Point2.Z
                    D3DXVec3TransformCoord vout, vin, matMesh
                    .Index2 = PointCache(MakePoint(vout.X, vout.Y, vout.Z))
                    Set .Point2 = Points(.Index1)
'                    .Point2.X = vout.X
'                    .Point2.Y = vout.Y
'                    .Point2.Z = vout.Z

                End With
            Next
        Next
    Next

    For Each f In Fields
        For Each s In f.Stars

            If Not s.Axis.Orbit Is Nothing Then
            
            For Each v In s.Axis.Orbit
                
                For Each m In v

                    With m

'                        VertexDirectX(.TriangleIndex + 0).X = (.Point1.X + s.Axis.Range.X)
'                        VertexDirectX(.TriangleIndex + 0).Y = (.Point1.Y + s.Axis.Range.Y)
'                        VertexDirectX(.TriangleIndex + 0).Z = (.Point1.Z + s.Axis.Range.Z)
'
'                        VertexDirectX(.TriangleIndex + 1).X = (.Point2.X + s.Axis.Range.X)
'                        VertexDirectX(.TriangleIndex + 1).Y = (.Point2.Y + s.Axis.Range.Y)
'                        VertexDirectX(.TriangleIndex + 1).Z = (.Point2.Z + s.Axis.Range.Z)
'
'                        VertexDirectX(.TriangleIndex + 2).X = (.Point3.X + s.Axis.Range.X)
'                        VertexDirectX(.TriangleIndex + 2).Y = (.Point3.Y + s.Axis.Range.Y)
'                        VertexDirectX(.TriangleIndex + 2).Z = (.Point3.Z + s.Axis.Range.Z)

                        VertexXAxis(0, .TriangleIndex) = (.Point1.X + s.Axis.Range.X)
                        VertexXAxis(1, .TriangleIndex) = (.Point2.X + s.Axis.Range.X)
                        VertexXAxis(2, .TriangleIndex) = (.Point3.X + s.Axis.Range.X)

                        VertexYAxis(0, .TriangleIndex) = (.Point1.Y + s.Axis.Range.Y)
                        VertexYAxis(1, .TriangleIndex) = (.Point2.Y + s.Axis.Range.Y)
                        VertexYAxis(2, .TriangleIndex) = (.Point3.Y + s.Axis.Range.Y)

                        VertexZAxis(0, .TriangleIndex) = (.Point1.Z + s.Axis.Range.Z)
                        VertexZAxis(1, .TriangleIndex) = (.Point2.Z + s.Axis.Range.Z)
                        VertexZAxis(2, .TriangleIndex) = (.Point3.Z + s.Axis.Range.Z)
                        
'                        DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 1, VertexDirectX(.TriangleIndex), Len(VertexDirectX(0))

                    End With
                    
                Next
                
            Next
            End If

        Next
    Next

End Sub

Public Sub Clear(Optional ByRef AtObj)

End Sub

Public Sub Frame()
    Dim f As Field
    For Each f In Fields
       f.Frame
    Next
End Sub

Public Sub Unload()
    Spaces.Clear
    Origins.Clear
    Rotates.Clear
    Scalars.Clear

    Fields.Clear
    Visions.Clear
    Points.Clear

    ObjectCount = 0
    TriangleCount = 0
    Erase TriangleFace
    
    Erase VertexXAxis
    Erase VertexYAxis
    Erase VertexZAxis

    Erase VertexDirectX

End Sub

Private Sub Class_Initialize()

    Set Spaces = New Space
    Set Origins = New Stars
    Set Rotates = New Stars
    Set Scalars = New Stars
End Sub

