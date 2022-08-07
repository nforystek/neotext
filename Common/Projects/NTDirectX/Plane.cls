VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Plane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'##################################################################################################################
'## Planular Description: Defines flat planes in the 3D world, i.e. a 3D canvas, or ever stretching ground floor ##
'##################################################################################################################

'    Level = 0 'default sky reference, can be circulated
'    Inland = 1 'aligned screen images on the 2d plane for swapping screen panels of electron
'    Island = 2 'a single axis value stretch all ways textured, can have a center polygon hole cut out of it (or square)
'    Canvas = 3 'a grid like 3d face plate just infront the camera displacement that mvoes along it from direction change, for maps and editors

Private pForm As PlaneTypes

Private pElectron As Electron
Private pElectronKey As String

Private pMolecule As Molecule
Attribute pMolecule.VB_VarHelpID = -1
Private pMoleculeKey As String

Private Area As Double

Private pLength As Single
Private pWidth As Single
Private pHeight As Single

Private pRows As Long
Private pColumns As Long
Private pPalettes As Long

Private pRadii As Single
Private pSegments As Single

Public Translucent As Boolean
Public Transparent As Boolean

Private pSerialize As Boolean

Public Property Get Collision() As Long
    Collision = Molecule.Collision
End Property
Public Property Let Collision(ByVal RHS As Long)
    Molecule.Collision = RHS
End Property
Public Property Get CollideIndex() As Long
    CollideIndex = Molecule.CollideIndex
End Property
Public Property Let CollideIndex(ByVal RHS As Long)
    Molecule.CollideIndex = RHS
End Property
Public Property Get CollideObject() As Long
    CollideObject = Molecule.CollideObject
End Property
Public Property Let CollideObject(ByVal RHS As Long)
    Molecule.CollideObject = RHS
End Property
Public Property Get CollideFaces() As Long
    CollideFaces = Molecule.CollideFaces
End Property
Public Property Let CollideFaces(ByVal RHS As Long)
    Molecule.CollideFaces = RHS
End Property

Public Property Get CulledFaces() As Long
    CulledFaces = Molecule.CulledFaces
End Property
Public Property Let CulledFaces(ByVal RHS As Long)
    Molecule.CulledFaces = RHS
End Property

Public Property Get OnInRange() As OnEvent
    Set OnInRange = Molecule.OnInRange
End Property
Public Property Set OnInRange(ByRef RHS As OnEvent)
    Set Molecule.OnInRange = RHS
End Property
Public Property Get OnOutRange() As OnEvent
    Set OnOutRange = Molecule.OnOutRange
End Property
Public Property Set OnOutRange(ByRef RHS As OnEvent)
    Set Molecule.OnOutRange = RHS
End Property
Public Property Get OnCollide() As OnEvent
    Set OnCollide = Molecule.OnCollide
End Property
Public Property Set OnCollide(ByRef RHS As OnEvent)
    Set Molecule.OnCollide = RHS
End Property

Public Property Get Serialize() As Boolean
    Serialize = pSerialize
End Property
Public Property Let Serialize(ByVal RHS As Boolean)
    pSerialize = RHS
    Molecule.Serialize = RHS
    Electron.Serialize = RHS
End Property

Public Property Get ToString(Optional ByVal Indent As String = "") As String

    ToString = ToString & Indent & "<Form>" & pForm & "</Form>" & vbCrLf
    
    If pLength <> 0 Then ToString = ToString & Indent & "<Length>" & pLength & "</Length>" & vbCrLf
    If pWidth <> 0 Then ToString = ToString & Indent & "<Width>" & pWidth & "</Width>" & vbCrLf
    If pHeight <> 0 Then ToString = ToString & Indent & "<Height>" & pHeight & "</Height>" & vbCrLf
    ToString = ToString & Indent & "<Translucent>" & Translucent & "</Translucent>" & vbCrLf
    ToString = ToString & Indent & "<Transparent>" & Transparent & "</Transparent>" & vbCrLf
    
    If pSegments <> 0 Then ToString = ToString & Indent & "<Segments>" & pSegments & "</Segments>" & vbCrLf
    If pRadii <> 0 Then ToString = ToString & Indent & "<Radii>" & pRadii & "</Radii>" & vbCrLf
    
    If pRows <> 0 Then ToString = ToString & Indent & "<Rows>" & pRows & "</Rows>" & vbCrLf
    If pColumns <> 0 Then ToString = ToString & Indent & "<Columns>" & pColumns & "</Columns>" & vbCrLf

    If pPalettes <> 0 Then ToString = ToString & Indent & "<Palettes>" & pPalettes & "</Palettes>" & vbCrLf

    ToString = ToString & Indent & "<ElectronKey>" & pElectronKey & "</ElectronKey>" & vbCrLf
    ToString = ToString & Indent & "<MoleculeKey>" & pMoleculeKey & "</MoleculeKey>" & vbCrLf

End Property
Public Property Get SurfaceArea() As Double
    SurfaceArea = Molecule.SurfaceArea
End Property
Private Sub Reset(ByVal CircleOrsquare As Boolean)
    If CircleOrsquare Then
        pColumns = 0
        pRows = 0
        pPalettes = 0
        If pSegments = 0 Then
            pSegments = 8
        End If
    Else
        pRadii = 0
        pSegments = 0
    End If
    Molecule.SurfaceArea = 0
End Sub
Public Property Get Radii() As Single
    Radii = pRadii
End Property
Public Property Let Radii(ByVal RHS As Single)
    pRadii = RHS
    Reset True
    Reworld
    Reform
End Property

Public Property Get Segments() As Single
    Segments = pSegments
End Property
Public Property Let Segments(ByVal RHS As Single)
    pSegments = RHS
    Reset True
    Reworld
    Reform
End Property

'Public Property Get Distance() As Single
'    Distance = FAR
'End Property
'Public Property Let Distance(ByVal RHS As Single)
'    FAR = RHS
'    Reworld
'    Reform
'End Property
Public Property Get Color() As Color
    Set Color = PlaneColor
End Property

Public Property Get Rows() As Long
    Rows = pRows
End Property
Public Property Let Rows(ByVal RHS As Long)
    pRows = RHS
    Reset False
    Reworld
    Reform
End Property
Public Property Get Columns() As Long
    Columns = pColumns
End Property
Public Property Let Columns(ByVal RHS As Long)
    pColumns = RHS
    Reset False
    Reworld
    Reform
End Property
Public Property Get Palettes() As Long
    Palettes = pPalettes
End Property
Public Property Let Palettes(ByVal RHS As Long)
    pPalettes = RHS
    Reset False
    Reworld
    Reform
End Property
Public Property Get Length() As Single
    Length = pLength
End Property
Public Property Let Length(ByVal RHS As Single)
    pLength = RHS
    Reset False
    Reworld
    Reform
End Property
Public Property Get Width() As Single
    Width = pWidth
End Property
Public Property Let Width(ByVal RHS As Single)
    pWidth = RHS
    Reset False
    Reworld
    Reform
End Property
Public Property Get Height() As Single
    Height = pHeight
End Property
Public Property Let Height(ByVal RHS As Single)
    pHeight = RHS
    Reset False
    Reworld
    Reform
End Property
Public Property Get Visible() As Boolean
    If ((pForm And World) = World) Then
        Visible = (Faces(Electron.FaceIndex).ImageCount > 5) And Molecule.Visible
    Else
        Visible = (Faces(Electron.FaceIndex).ImageCount > 0) And Molecule.Visible
    End If
End Property
Public Property Let Visible(ByVal RHS As Boolean)
    Molecule.Visible = RHS
End Property
Public Property Get Form() As Long
    If ((pForm And Island) = Island) Then
        Form = Island Or Plateau
    ElseIf ((pForm And Inland) = Inland) Then
        Form = Plateau Or Inland
    Else
        Form = pForm
    End If
End Property

Public Property Let Form(ByVal RHS As Long)
    If pForm <> RHS Then
        If pForm <> 0 Then
            If Faces(Electron.FaceIndex).ImageCount > 1 Then
                ReDim Preserve Faces(Electron.FaceIndex).Images(1 To 1) As Long
                Faces(Electron.FaceIndex).ImageCount = 1
            End If
        End If

        If ((pForm And Island) = Island) Then
            pForm = Island Or Plateau
        ElseIf ((pForm And Inland) = Inland) Then
            pForm = Plateau Or Inland
        Else
            pForm = RHS
        End If

        Reworld
        Reform

    End If
End Property
Private Sub Reworld()
    With Faces(Electron.FaceIndex)
        Molecule.FaceIndex = Electron.FaceIndex
        
        If ((pForm And World) = World) Then
            If ((.Dimension.Width <> 0) Or (pWidth <> 0) Or (pLength <> 0)) And ((.Dimension.Height <> 0) Or (pHeight <> 0)) Then
            
                If Not .VBuffer Is Nothing And CollideIndex > -1 Then
                    DelCollision Me, ((UBound(Faces(Electron.FaceIndex).Verticies) + 1) / 3)
                End If
                
                ReDim .Verticies(0 To 35) As MyVertex
            
                Dim pX As Single
                Dim pY As Single
                Dim pZ As Single
                Dim pU As Single
                Dim pV As Single
                pU = 1
                pV = 1
                pX = pWidth
                If pX = 0 Then pX = ((.Dimension.Width * 2) + (.Dimension.Height * 2))
                pY = pHeight
                If pY = 0 Then pY = ((.Dimension.Width * 2) + (.Dimension.Height * 2))
                pZ = pLength
                If pZ = 0 Then pZ = ((.Dimension.Width * 2) + (.Dimension.Height * 2))
                
                Molecule.SurfaceArea = CreateSquare(.Verticies, 0, _
                                        MakeVector(-(pX \ 2), -(pY \ 2), (pZ \ 2)), _
                                        MakeVector(-(pX \ 2), -(pY \ 2), -(pZ \ 2)), _
                                        MakeVector(-(pX \ 2), (pY \ 2), -(pZ \ 2)), _
                                        MakeVector(-(pX \ 2), (pY \ 2), (pZ \ 2)), pU, pV)
                Molecule.SurfaceArea = Molecule.SurfaceArea + CreateSquare(.Verticies, 6, _
                                        MakeVector(-(pX \ 2), -(pY \ 2), -(pZ \ 2)), _
                                        MakeVector((pX \ 2), -(pY \ 2), -(pZ \ 2)), _
                                        MakeVector((pX \ 2), (pY \ 2), -(pZ \ 2)), _
                                        MakeVector(-(pX \ 2), (pY \ 2), -(pZ \ 2)), pU, pV)
                Molecule.SurfaceArea = Molecule.SurfaceArea + CreateSquare(.Verticies, 12, _
                                        MakeVector((pX \ 2), -(pY \ 2), -(pZ \ 2)), _
                                        MakeVector((pX \ 2), -(pY \ 2), (pZ \ 2)), _
                                        MakeVector((pX \ 2), (pY \ 2), (pZ \ 2)), _
                                        MakeVector((pX \ 2), (pY \ 2), -(pZ \ 2)), pU, pV)
                Molecule.SurfaceArea = Molecule.SurfaceArea + CreateSquare(.Verticies, 18, _
                                        MakeVector((pX \ 2), -(pY \ 2), -(pZ \ 2)), _
                                        MakeVector(-(pX \ 2), -(pY \ 2), -(pZ \ 2)), _
                                        MakeVector(-(pX \ 2), -(pY \ 2), (pZ \ 2)), _
                                        MakeVector((pX \ 2), -(pY \ 2), (pZ \ 2)), pU, pV)
                Molecule.SurfaceArea = Molecule.SurfaceArea + CreateSquare(.Verticies, 24, _
                                        MakeVector((pX \ 2), -(pY \ 2), (pZ \ 2)), _
                                        MakeVector(-(pX \ 2), -(pY \ 2), (pZ \ 2)), _
                                        MakeVector(-(pX \ 2), (pY \ 2), (pZ \ 2)), _
                                        MakeVector((pX \ 2), (pY \ 2), (pZ \ 2)), pU, pV)
                Molecule.SurfaceArea = Molecule.SurfaceArea + CreateSquare(.Verticies, 30, _
                                        MakeVector((pX \ 2), (pY \ 2), (pZ \ 2)), _
                                        MakeVector(-(pX \ 2), (pY \ 2), (pZ \ 2)), _
                                        MakeVector(-(pX \ 2), (pY \ 2), -(pZ \ 2)), _
                                        MakeVector((pX \ 2), (pY \ 2), -(pZ \ 2)), pU, pV)
                
                Set .VBuffer = DDevice.CreateVertexBuffer(Len(.Verticies(0)) * (UBound(.Verticies) + 1), 0, FVF_RENDER, D3DPOOL_DEFAULT)
                D3DVertexBuffer8SetData .VBuffer, 0, Len(.Verticies(0)) * (UBound(.Verticies) + 1), 0, .Verticies(0)
                
                AddPlaneCollision Me, ((UBound(Faces(Electron.FaceIndex).Verticies) + 1) / 3), Faces(Me.Electron.FaceIndex).Verticies
            End If
        End If
    End With
End Sub
Private Sub Reform(Optional ByVal ResetFor As Boolean = False)
    With Faces(Electron.FaceIndex)
        Molecule.FaceIndex = Electron.FaceIndex
        
        If ((Form And PlaneTypes.Plateau) = Plateau) Then
            If Not .VBuffer Is Nothing And CollideIndex > -1 Then
                DelCollision Me, ((UBound(Faces(pElectron.FaceIndex).Verticies) + 1) / 3)
            End If
            
            Erase .Verticies
            Set .VBuffer = Nothing
            
            If (pSegments > 0) Then
            
                ReDim .Verticies(0 To 5) As MyVertex
                
                Molecule.SurfaceArea = CreateSquare(.Verticies, 0, _
                                            MakeVector((FAR / 2), 0, -(FAR / 2)), _
                                            MakeVector(-(FAR / 2), 0, -(FAR / 2)), _
                                            MakeVector(-(FAR / 2), 0, (FAR / 2)), _
                                            MakeVector((FAR / 2), 0, (FAR / 2)), _
                                            (FAR * ((.Dimension.Width / 2) / 100)), (FAR * ((.Dimension.Height / 2) / 100)))
            
                
                If ((pForm And Inland) = Inland) Then
                    
                    Molecule.SurfaceArea = Molecule.SurfaceArea + CreateCircle(.Verticies, (FAR / 2), pSegments, pRadii, .Dimension.Width, .Dimension.Height)
                ElseIf ((pForm And Island) = Island) Then
                    Molecule.SurfaceArea = Molecule.SurfaceArea + CreateCircle(.Verticies, pRadii, pSegments, 0, .Dimension.Width, .Dimension.Height)

                End If
            Else

                Dim i As Single
                Dim X As Single
                Dim Z As Single
                    
                If ((pLength <> 0) And (pWidth <> 0)) And ((Form And PlaneTypes.Island) = PlaneTypes.Island) Then
                    If ((pRows > 0) And (pColumns > 0)) Then
    
                        If .ImageCount < pRows * pColumns Then
                            .ImageCount = pRows * pColumns
                            ReDim Preserve .Images(1 To .ImageCount) As Long
                        End If
                        
                        For i = 2 To pRows * pColumns
                            If .Images(i) = 0 Then .Images(i) = .Images(1)
                        Next
                        
                        X = -((pRows \ 2) * pWidth) - pWidth
                        Z = ((pColumns \ 2) * pLength)
                            
                        ReDim .Verticies(0 To (((pRows * pColumns) * 6) - 1)) As MyVertex
                        For i = 0 To UBound(.Verticies) Step 6
                        
                            If (i \ 6) Mod pColumns = 0 Then
                                Z = Z - pLength
                                X = -((pRows \ 2) * pWidth)
                            Else
                                X = X + pWidth
                            End If
            
                            Molecule.SurfaceArea = CreateSquare(.Verticies, i, _
                                                        MakeVector(((pWidth / 2) + X), 0, (-(pLength / 2) + Z)), _
                                                        MakeVector((-(pWidth / 2) + X), 0, (-(pLength / 2) + Z)), _
                                                        MakeVector((-(pWidth / 2) + X), 0, ((pLength / 2) + Z)), _
                                                        MakeVector(((pWidth / 2) + X), 0, ((pLength / 2) + Z)), _
                                                        (pWidth * ((.Dimension.Width / 2) / 100)), (pLength * ((.Dimension.Height / 2) / 100)))
                        Next
                    Else
                    
                        ReDim .Verticies(0 To 5) As MyVertex
            
                        Molecule.SurfaceArea = CreateSquare(.Verticies, 0, _
                                                    MakeVector((pWidth / 2), 0, -(pLength / 2)), _
                                                    MakeVector(-(pWidth / 2), 0, -(pLength / 2)), _
                                                    MakeVector(-(pWidth / 2), 0, (pLength / 2)), _
                                                    MakeVector((pWidth / 2), 0, (pLength / 2)), _
                                                    (pWidth * ((.Dimension.Width / 2) / 100)), (pLength * ((.Dimension.Height / 2) / 100)))
                    
                    End If
                
                Else
                    
                    ReDim .Verticies(0 To 5) As MyVertex
                    
                    Molecule.SurfaceArea = CreateSquare(.Verticies, 0, _
                                                MakeVector((FAR / 2), 0, -(FAR / 2)), _
                                                MakeVector(-(FAR / 2), 0, -(FAR / 2)), _
                                                MakeVector(-(FAR / 2), 0, (FAR / 2)), _
                                                MakeVector((FAR / 2), 0, (FAR / 2)), _
                                                (FAR * ((.Dimension.Width / 2) / 100)), (FAR * ((.Dimension.Height / 2) / 100)))
                                                    
                    If ((pLength <> 0) And (pWidth <> 0)) And ((Form And Inland) = Inland) Then
                    
                        ReDim Preserve .Verticies(0 To 29) As MyVertex
                        
                        Molecule.SurfaceArea = Molecule.SurfaceArea + CreateSquare(.Verticies, 6, _
                                                    MakeVector(-(pWidth / 2), 0, -(pLength / 2)), _
                                                    MakeVector(-((FAR / 2) - (pWidth / 2)), 0, -(pLength / 2)), _
                                                    MakeVector(-((FAR / 2) - (pWidth / 2)), 0, ((FAR / 2) - (pLength / 2))), _
                                                    MakeVector(-(pWidth / 2), 0, ((FAR / 2) - (pLength / 2))), _
                                                    (FAR * ((.Dimension.Width / 2) / 100)), (FAR * ((.Dimension.Height / 2) / 100)))
                                                    
                        Molecule.SurfaceArea = Molecule.SurfaceArea + CreateSquare(.Verticies, 12, _
                                                    MakeVector(-(pWidth / 2), 0, (pLength / 2)), _
                                                    MakeVector(-(pWidth / 2), 0, ((FAR / 2) - (pLength / 2))), _
                                                    MakeVector(((FAR / 2) - (pWidth / 2)), 0, ((FAR / 2) - (pLength / 2))), _
                                                    MakeVector(((FAR / 2) - (pWidth / 2)), 0, -(pLength / 2)), _
                                                    (FAR * ((.Dimension.Width / 2) / 100)), (FAR * ((.Dimension.Height / 2) / 100)))
                                                    
                        Molecule.SurfaceArea = Molecule.SurfaceArea + CreateSquare(.Verticies, 18, _
                                                    MakeVector((pWidth / 2), 0, (pLength / 2)), _
                                                    MakeVector(((FAR / 2) - (pWidth / 2)), 0, (pLength / 2)), _
                                                    MakeVector(((FAR / 2) - (pWidth / 2)), 0, -((FAR / 2) - (pLength / 2))), _
                                                    MakeVector((pWidth / 2), 0, -((FAR / 2) - (pLength / 2))), _
                                                    (FAR * ((.Dimension.Width / 2) / 100)), (FAR * ((.Dimension.Height / 2) / 100)))
                                                    
                        Molecule.SurfaceArea = Molecule.SurfaceArea + CreateSquare(.Verticies, 24, _
                                                    MakeVector((pWidth / 2), 0, -(pLength / 2)), _
                                                    MakeVector((pWidth / 2), 0, -((FAR / 2) - (pLength / 2))), _
                                                    MakeVector(-((FAR / 2) - (pWidth / 2)), 0, -((FAR / 2) - (pLength / 2))), _
                                                    MakeVector(-((FAR / 2) - (pWidth / 2)), 0, -(pLength / 2)), _
                                                    (FAR * ((.Dimension.Width / 2) / 100)), (FAR * ((.Dimension.Height / 2) / 100)))
                    End If
                    
                End If
            End If
            Set .VBuffer = DDevice.CreateVertexBuffer(Len(.Verticies(0)) * (UBound(.Verticies) + 1), 0, FVF_RENDER, D3DPOOL_DEFAULT)
            D3DVertexBuffer8SetData .VBuffer, 0, Len(.Verticies(0)) * (UBound(.Verticies) + 1), 0, .Verticies(0)
                    
            AddPlaneCollision Me, ((UBound(Faces(Electron.FaceIndex).Verticies) + 1) / 3), Faces(Me.Electron.FaceIndex).Verticies
        End If

    End With
End Sub

Public Property Get Electron() As Electron
    If pElectronKey = "" Then
        Set pElectron = New Electron
        pElectronKey = Include.Unnamed(All)
        All.Add pElectron, pElectronKey
        Electrons.Add pElectron, pElectronKey
    End If
    Set Electron = All(pElectronKey)
End Property
Public Property Get ElectronKey() As String
    ElectronKey = pElectronKey
End Property
Public Property Let ElectronKey(ByVal RHS As String)
    pElectronKey = RHS
End Property
Public Property Get Molecule() As Molecule
    If pMoleculeKey = "" Then
        Set pMolecule = New Molecule
        pMoleculeKey = Include.Unnamed(All)
        pMolecule.MeshIndex = -1
        All.Add pMolecule, pMoleculeKey
        Molecules.Add pMolecule, pMoleculeKey
    End If
    Set Molecule = All(pMoleculeKey)
End Property
Public Property Get MoleculeKey() As String
    MoleculeKey = pMoleculeKey
End Property
Public Property Let MoleculeKey(ByVal RHS As String)
    pMoleculeKey = RHS
End Property
Public Property Get Motions() As NTNodes10.Collection
    Set Motions = Molecule.Motions
End Property

Public Property Get File(Optional ByVal Number As Long = 0) As Variant
    File = Electron.File(Number)
End Property
Public Property Let File(Optional ByVal Number As Long = 0, ByVal RHS As Variant)
    Electron.File(Number) = RHS
    Reworld
    Reform
End Property

Public Property Get Origin() As Vector
    Set Origin = Molecule.Origin
End Property
Public Property Get Rotate() As Vector
    Set Rotate = Molecule.Rotate
End Property
Public Property Get Relative() As Boolean
    Relative = Molecule.Relative
End Property
Public Property Let Relative(ByVal RHS As Boolean)
    Molecule.Relative = RHS
End Property
Public Property Get Absolute() As Boolean
    Absolute = Molecule.Absolute
End Property
Public Property Let Absolute(ByVal RHS As Boolean)
    Molecule.Absolute = RHS
End Property

Private Sub Class_Terminate()
    If Not pElectron Is Nothing Then
        If pElectron.FaceIndex > 0 And pElectron.FaceIndex <= UBound(Faces) Then
            If Not Faces(pElectron.FaceIndex).VBuffer Is Nothing Then
                DelCollision Me, ((UBound(Faces(pElectron.FaceIndex).Verticies) + 1) / 3)
            End If
        End If
    End If
            
    If All.Exists(ElectronKey) Then All.Remove ElectronKey
    If Electrons.Exists(ElectronKey) Then Electrons.Remove ElectronKey
    Set pElectron = Nothing

    If All.Exists(MoleculeKey) Then All.Remove MoleculeKey
    If Molecules.Exists(MoleculeKey) Then Molecules.Remove MoleculeKey
    Set pMolecule = Nothing
End Sub

