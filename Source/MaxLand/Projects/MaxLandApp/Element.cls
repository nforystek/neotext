VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Element"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum Collides
    None = 0
    InDoor = 1
    Ground = 2
    Liquid = 3
    Ladder = 4
End Enum


Public Enum Moving
    None = 0
    Level = 1
    Flying = 2
    Falling = 4
    Stepping = 8
End Enum

Private pOrigin As Point
Private pOffset As Point
Private pScaled As Point
Private pRotate As Point

Private pKey As String
Private pVisible As Boolean
Private pGravitational As Boolean
Private pAlphablend As Boolean
Private pWireFrame As Boolean
Private pEffect As Long
Private pCollision As Boolean
Private pMotions As ntnodes10.Collection


Private pSerialize As Boolean

Private pVisualIndex As Long
Private pBoundsIndex As Long

Private pIsMoving As Long
Private pInLiquid As Boolean
Private pOnLadder As Boolean

Private pDirect As Point
Private pTwists As Point
Private pScalar As Point


Private pCollideIndex As Long
Private pCollideObject As Long
Private pCollideFaces As Long
Private pCulledFaces As Long

Private pReplacerVals As ntnodes10.Collection
Private pReplacerKeys As ntnodes10.Collection
Attribute pReplacerKeys.VB_VarMemberFlags = "440"

Private pAttachments As ntnodes10.Collection
Private pAttachedTo As String

Private Matrix As D3DMATRIX

Public Property Get Attachments() As ntnodes10.Collection
    Set Attachments = pAttachments
End Property
Public Property Set Attachments(ByRef RHS As ntnodes10.Collection)
    Set pAttachments = RHS
End Property

Public Property Get AttachedTo() As String
    AttachedTo = pAttachedTo
End Property
Public Property Let AttachedTo(ByVal RHS As String)
    If pAttachedTo <> "" Then
        All(pAttachedTo).Attachments.Remove Me.Key
    End If
    pAttachedTo = RHS
    If RHS <> "" Then
        All(pAttachedTo).Attachments.Add Me, RHS
    End If
    
End Property

Friend Property Get ReplacerVals() As ntnodes10.Collection
    Set ReplacerVals = pReplacerVals
End Property
Friend Property Set ReplacerVals(ByRef RHS As ntnodes10.Collection)
 Set pReplacerVals = RHS
End Property
Friend Property Get ReplacerKeys() As ntnodes10.Collection
    Set ReplacerKeys = pReplacerKeys
End Property
Friend Property Set ReplacerKeys(ByRef RHS As ntnodes10.Collection)
    Set pReplacerKeys = RHS
End Property

Public Sub Replacer(ByVal TextFilename As String, ByVal BoardKey As String)
    If Key = "" Then
        Err.Raise 8, App.EXEName, "Key must be set before using Replacer"
    Else
        ReplacerVals.Add BoardKey, Key & "_" & Replace(TextFilename, ".", "")
        ReplacerKeys.Add Key & "_" & Replace(TextFilename, ".", "")
    End If
End Sub

Public Function AddMotion(ByRef Action As Long, ByVal aGUID As String, ByRef Data As Point, Optional ByRef Emphasis As Single = 0, Optional ByVal Friction As Single = 0, Optional ByVal Reactive As Single = -1, Optional ByVal Recount As Single = -1, Optional Script As String = "") As String

    Dim act As Motion
    Set act = New Motion
    With act
        If aGUID = "" Then
            .Identity = Replace(modGuid.GUID, "-", "K")
        Else
            .Identity = aGUID
        End If
        .Action = Action
        Set .Data = Data
        .Emphasis = Emphasis
        .Initials = Emphasis
        .Friction = Friction
        .Reactive = Reactive
        .latency = Timer
        .Recount = Recount
        .Script = Script
        AddMotion = .Identity
    End With
    If pMotions.Count > 0 Then
        pMotions.Add act, , 1
    Else
        pMotions.Add act
    End If
    
'    If Not pAttachments Is Nothing Then
'        If pAttachments.Count > 0 Then
'            Dim e2 As Element
'            For Each e2 In pAttachments
'                If e2.Motions.Count > 0 Then
'                    e2.Motions.Add act, , 1
'                Else
'                    e2.Motions.Add act
'                End If
'            Next
'
'        End If
'    End If

    
    
    Set act = Nothing
    
    

End Function

Public Function DeleteMotion(ByVal MGUID As String) As Boolean

    Dim A As Long
    Dim act As Motion
    A = 1
    Do While A <= Motions.Count
        Set act = Motions(A)
        If act.Identity = MGUID Or (act.Identity = "") Then
            Motions.Remove A
            DeleteMotion = True
        Else
            A = A + 1
        End If
        Set act = Nothing
    Loop

'a prior way that only removed one occurance of the idenity in the collection
'only it started with checking the very last item,and removed it if was idenity
'otherwise it begane a forward iteration, as well it did not check for blank id's
'
'this one is slightly different, it only delete's one occurance and starts from
'the back of the list, iterating all, it also considers blank ID match for removal
'
'    A = Motions.count
'    Do While A > 0
'        Set act = Motions(A)
'        If act.Identity = MGUID Or (act.Identity = "") Then
'            Motions.Remove A
'            DeleteMotion = True
'            Set act = Nothing
'            Exit Function
'        End If
'        A = A - 1
'    Loop
'    Set act = Nothing
'
'
'because of array differences the above was working, yet no longer in objects
'so the choice is a full iteration of the collection from the start and blanks
'are considered matches for removal as well as ID matching, and all matches are
'removed if satisfying those conditions, now it acts like before, with out
'stacking too many motions that do or don't remove bogging down the system


End Function

Public Sub ClearMotions()
    Dim act As Motion
    Do While pMotions.Count > 0
        Set act = pMotions(1)
        pMotions.Remove 1
        Set act = Nothing
    Loop
End Sub

Public Property Get Motions() As ntnodes10.Collection
    Set Motions = pMotions
End Property
Public Property Set Motions(ByRef RHS As ntnodes10.Collection)
    Set pMotions = RHS
End Property

Public Property Get Visible() As Boolean
    Visible = pVisible
End Property
Public Property Let Visible(ByVal RHS As Boolean)
    pVisible = RHS
End Property

Public Property Get VisualIndex() As Long
Attribute VisualIndex.VB_MemberFlags = "440"
    VisualIndex = pVisualIndex
End Property
Public Property Let VisualIndex(ByVal RHS As Long)
    pVisualIndex = RHS
End Property
Public Property Get VisualModel() As String
    If pVisualIndex > 0 Then
        VisualModel = Replace(Meshes(pVisualIndex).FileName, ScriptRoot & "\Models\", "")
    End If
End Property
Public Property Let VisualModel(ByVal RHS As String)

    Dim i As Long
    If MeshCount > 0 Then
        For i = LBound(Meshes) To UBound(Meshes)
            If Meshes(i).FileName = LCase(RHS) Then
                pVisualIndex = i
                Exit For
            End If
        Next
    End If
    
    If pVisualIndex = 0 Then
        MeshCount = MeshCount + 1
        ReDim Preserve Meshes(1 To MeshCount) As MyMesh
        pVisualIndex = MeshCount
        Meshes(pVisualIndex).FileName = LCase(RHS)
        If PathExists(ScriptRoot & "\Models\" & Meshes(pVisualIndex).FileName, True) Then
            CreateMesh ScriptRoot & "\Models\" & Meshes(pVisualIndex).FileName, Meshes(pVisualIndex).Mesh, Meshes(pVisualIndex).MaterialBuffer, _
                    ToVector(pOrigin), ToVector(pScaled), Meshes(pVisualIndex).Materials, Meshes(pVisualIndex).Textures, _
                    Meshes(pVisualIndex).Verticies, Meshes(pVisualIndex).Indicies, Meshes(pVisualIndex).MaterialCount
        Else
            ReDim Meshes(pVisualIndex).Textures(0 To 0) As Direct3DTexture8
            ReDim Meshes(pVisualIndex).Materials(0 To 0) As D3DMATERIAL8
            pVisualIndex = 0
        End If
    End If

End Property

Public Property Get FileName() As String
    FileName = BoundsModel
End Property
Public Property Let FileName(ByVal RHS As String)
    BoundsModel = RHS
End Property

Public Property Get BoundsIndex() As Long
Attribute BoundsIndex.VB_MemberFlags = "440"
    BoundsIndex = pBoundsIndex
End Property
Public Property Let BoundsIndex(ByVal RHS As Long)
    pBoundsIndex = RHS
End Property
Public Property Get BoundsModel() As String
    If pBoundsIndex > 0 Then
        BoundsModel = Replace(Meshes(pBoundsIndex).FileName, ScriptRoot & "\Models\", "")
    End If
End Property
Public Property Let BoundsModel(ByVal RHS As String)

    If (Not pCollision) Then
        If (pCollideIndex > -1) And (pBoundsIndex > 0) Then
            DelCollision Me
        End If
    End If
    DefaultMatrix
  
    Dim i As Long
    If MeshCount > 0 Then
        For i = LBound(Meshes) To UBound(Meshes)
            If Meshes(i).FileName = LCase(RHS) Then
                pBoundsIndex = i
                Exit For
            End If
        Next
    End If

    If pBoundsIndex = 0 Then
        MeshCount = MeshCount + 1
        ReDim Preserve Meshes(1 To MeshCount) As MyMesh
        pBoundsIndex = MeshCount
        Meshes(pBoundsIndex).FileName = LCase(RHS)
        If PathExists(ScriptRoot & "\Models\" & Meshes(pBoundsIndex).FileName, True) Then
            CreateMesh ScriptRoot & "\Models\" & Meshes(pBoundsIndex).FileName, Meshes(pBoundsIndex).Mesh, Meshes(pBoundsIndex).MaterialBuffer, _
                    ToVector(pOrigin), ToVector(pScaled), Meshes(pBoundsIndex).Materials, Meshes(pBoundsIndex).Textures, _
                    Meshes(pBoundsIndex).Verticies, Meshes(pBoundsIndex).Indicies, Meshes(pBoundsIndex).MaterialCount
        Else
            ReDim Meshes(pBoundsIndex).Textures(0 To 0) As Direct3DTexture8
            ReDim Meshes(pBoundsIndex).Materials(0 To 0) As D3DMATERIAL8
            pBoundsIndex = 0
        End If
    End If

    If pCollision Or (pCollideIndex = 0) Then
        If (pCollideIndex <= 0) And (pBoundsIndex > 0) Then
            PrepairMatrix
            AddCollision Me
        ElseIf (pBoundsIndex = 0) Then
            pCollideIndex = -1
        End If
    Else
        pCollideIndex = -1
    End If
                        
End Property


Public Property Get CollideIndex() As Long
Attribute CollideIndex.VB_MemberFlags = "440"
    CollideIndex = pCollideIndex
End Property
Public Property Let CollideIndex(ByVal RHS As Long)
    pCollideIndex = RHS
End Property
Public Property Get CollideObject() As Long
Attribute CollideObject.VB_MemberFlags = "440"
    CollideObject = pCollideObject
End Property
Public Property Let CollideObject(ByVal RHS As Long)
    pCollideObject = RHS
End Property
Public Property Get CollideFaces() As Long
Attribute CollideFaces.VB_MemberFlags = "440"
    CollideFaces = pCollideFaces
End Property
Public Property Let CollideFaces(ByVal RHS As Long)
    pCollideFaces = RHS
End Property

Public Property Get Collision() As Boolean
    Collision = pCollision And ((pCollideIndex > -1) And (pBoundsIndex > 0))
End Property
Public Property Let Collision(ByVal RHS As Boolean)
    pCollision = RHS
    If pCollision Or (pCollideIndex = 0) Then
        If (pBoundsIndex > 0) And ((pCollideIndex <= 0)) Then
            PrepairMatrix
            AddCollision Me
        ElseIf (pBoundsIndex = 0) Then
            pCollideIndex = -1
        End If
    Else
        If (pBoundsIndex > 0) And (pCollideIndex > -1) Then
            DelCollision Me
        End If
        pCollideIndex = -1
    End If
End Property

Public Property Get CulledFaces() As Long
Attribute CulledFaces.VB_MemberFlags = "440"
    CulledFaces = pCulledFaces
End Property
Public Property Let CulledFaces(ByVal RHS As Long)
    pCulledFaces = RHS
End Property

Public Property Get Alphablend() As Boolean
    Alphablend = pAlphablend
End Property
Public Property Let Alphablend(ByVal RHS As Boolean)
    pAlphablend = RHS
End Property
Public Property Get WireFrame() As Boolean
    WireFrame = pWireFrame
End Property
Public Property Let WireFrame(ByVal RHS As Boolean)
    pWireFrame = RHS
End Property


Public Property Get Effect() As Long
    Effect = pEffect
End Property
Public Property Let Effect(ByVal RHS As Long)
    pEffect = RHS
End Property

Public Property Get IsMoving() As Long
Attribute IsMoving.VB_MemberFlags = "440"
    IsMoving = pIsMoving
End Property
Public Property Let IsMoving(ByVal RHS As Long)
    pIsMoving = RHS
End Property
Public Property Get InLiquid() As Boolean
Attribute InLiquid.VB_MemberFlags = "440"
    InLiquid = pInLiquid
End Property
Public Property Let InLiquid(ByVal RHS As Boolean)
    pInLiquid = RHS
End Property
Public Property Get OnLadder() As Boolean
Attribute OnLadder.VB_MemberFlags = "440"
    OnLadder = pOnLadder
End Property
Public Property Let OnLadder(ByVal RHS As Boolean)
    pOnLadder = RHS
End Property


Public Property Get Origin() As Point
    Set Origin = pOrigin

End Property
Public Property Set Origin(ByRef RHS As Point)
    Set pOrigin = RHS
End Property

Public Property Get Offset() As Point
    Set Offset = pOffset
End Property
Public Property Set Offset(ByRef RHS As Point)
    Set pOffset = RHS
End Property

Public Property Get Scaled() As Point
    Set Scaled = pScaled
End Property
Public Property Set Scaled(ByRef RHS As Point)
    Set pScaled = RHS
End Property

Public Property Get Rotate() As Point
    Set Rotate = pRotate
End Property
Public Property Set Rotate(ByRef RHS As Point)
    Set pRotate = RHS
End Property


Public Property Get Direct() As Point
Attribute Direct.VB_MemberFlags = "440"
    Set Direct = pDirect
End Property
Public Property Set Direct(ByRef RHS As Point)
    Set pDirect = RHS
End Property

Public Property Get Twists() As Point
Attribute Twists.VB_MemberFlags = "440"
    Set Twists = pTwists
End Property
Public Property Set Twists(ByRef RHS As Point)
    Set pTwists = RHS
End Property

Public Property Get Scalar() As Point
Attribute Scalar.VB_MemberFlags = "440"
    Set Scalar = pScalar
End Property
Public Property Set Scalar(ByRef RHS As Point)
    Set pScalar = RHS
End Property

Public Property Get Serialize() As Boolean
    Serialize = True
End Property
Public Property Let Serialize(ByVal RHS As Boolean)
    pSerialize = RHS
End Property

Public Property Get Value() As Long
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "40"
    Value = ObjPtr(Me)
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

Public Property Get Gravitational() As Boolean
    Gravitational = pGravitational
End Property
Public Property Let Gravitational(ByVal RHS As Boolean)
    pGravitational = RHS
End Property

Friend Sub PrepairMatrix()
                        
    D3DXMatrixIdentity Matrix
    D3DXMatrixTranslation Matrix, pOffset.X, pOffset.Y, pOffset.Z
    D3DXMatrixRotationX Matrix, pRotate.X * RADIAN
    D3DXMatrixRotationY Matrix, pRotate.Y * RADIAN
    D3DXMatrixRotationZ Matrix, pRotate.Z * RADIAN
    D3DXMatrixScaling Matrix, pScaled.X, pScaled.Y, pScaled.Z

    D3DXMatrixIdentity matWorld

End Sub
Friend Sub ApplyMatrix()

'    Dim matMesh As D3DMATRIX
'
'
'    D3DXMatrixIdentity matMesh
'    D3DXMatrixRotationX Matrix, Rotate.X * (PI / 180)
'    D3DXMatrixMultiply Matrix, matMesh, Matrix
'    D3DXMatrixMultiply matMesh, Matrix, matMesh
'    D3DXMatrixRotationY Matrix, Rotate.Y * (PI / 180)
'    D3DXMatrixMultiply Matrix, matMesh, Matrix
'    D3DXMatrixMultiply matMesh, Matrix, matMesh
'    D3DXMatrixRotationZ Matrix, Rotate.Z * (PI / 180)
'    D3DXMatrixMultiply Matrix, matMesh, Matrix
'    D3DXMatrixMultiply matMesh, Matrix, matMesh
'    D3DXMatrixTranslation Matrix, Origin.X, Origin.Y, Origin.Z
'    D3DXMatrixMultiply Matrix, matMesh, Matrix
'    DDevice.SetTransform D3DTS_WORLD, Matrix
'    D3DXMatrixScaling matMesh, Scaled.X, Scaled.Y, Scaled.Z
'    D3DXMatrixMultiply matMesh, matMesh, Matrix
'    DDevice.SetTransform D3DTS_WORLD, matMesh
'
            
    Dim matMesh As D3DMATRIX
    

    D3DXMatrixIdentity matWorld
    D3DXMatrixIdentity matMesh

    D3DXMatrixTranslation Matrix, Origin.X, Origin.Y, Origin.Z
    D3DXMatrixScaling matMesh, Scaled.X, Scaled.Y, Scaled.Z
    D3DXMatrixMultiply Matrix, matMesh, Matrix
    DDevice.SetTransform D3DTS_WORLD, Matrix
    
    
    D3DXMatrixRotationX matMesh, Rotate.X * (PI / 180)
    D3DXMatrixMultiply Matrix, matMesh, Matrix
    D3DXMatrixRotationY matMesh, Rotate.Y * (PI / 180)
    D3DXMatrixMultiply matMesh, Matrix, matMesh
    D3DXMatrixRotationZ Matrix, Rotate.Z * (PI / 180)
    D3DXMatrixMultiply Matrix, matMesh, Matrix
    DDevice.SetTransform D3DTS_WORLD, Matrix
    
    
    D3DXMatrixTranslation matMesh, Offset.X, Offset.Y, Offset.Z
    D3DXMatrixMultiply Matrix, matMesh, Matrix
    DDevice.SetTransform D3DTS_WORLD, Matrix



End Sub

Friend Sub PlayerMatrix()
    
    D3DXMatrixIdentity matWorld
    D3DXMatrixIdentity Matrix
    
    D3DXMatrixScaling Matrix, pScaled.X, pScaled.Y, pScaled.Z
    D3DXMatrixMultiply matWorld, matWorld, Matrix
    
    D3DXMatrixRotationY Matrix, -Player.Angle
    D3DXMatrixMultiply matWorld, matWorld, Matrix
    
    D3DXMatrixTranslation Matrix, pOrigin.X, pOrigin.Y, pOrigin.Z
    D3DXMatrixMultiply matWorld, matWorld, Matrix
    
    DDevice.SetTransform D3DTS_WORLD, matWorld

End Sub

Friend Sub SetWorldMatrix()
    DDevice.SetTransform D3DTS_WORLD, Matrix
End Sub


Friend Function PointMatrix(ByRef p As Point) As Point
    Dim v As D3DVECTOR
    v = ToVector(p)
    D3DXVec3TransformCoord v, v, Matrix
    Set PointMatrix = ToPoint(v)
End Function

Friend Sub DefaultMatrix()
    D3DXMatrixIdentity Matrix
    D3DXMatrixTranslation Matrix, 0, 0, 0
    D3DXMatrixRotationX Matrix, 0
    D3DXMatrixRotationY Matrix, 0
    D3DXMatrixRotationZ Matrix, 0
    D3DXMatrixScaling Matrix, 1, 1, 1

    D3DXMatrixIdentity matWorld
    D3DXMatrixTranslation matWorld, 0, 0, 0
    D3DXMatrixRotationX matWorld, 0
    D3DXMatrixRotationY matWorld, 0
    D3DXMatrixRotationZ matWorld, 0
    D3DXMatrixScaling matWorld, 1, 1, 1
    
    DDevice.SetTransform D3DTS_WORLD, Matrix
    DDevice.SetTransform D3DTS_WORLD, matWorld

End Sub
    
Private Sub Class_Initialize()
    Stats_Element_Count = Stats_Element_Count + 1
    Key = "K" & ObjPtr(Me)

    pBoundsIndex = 0
    pVisualIndex = 0
    pEffect = Collides.None
    pGravitational = False
    pAlphablend = False
    pCollideIndex = 0
    pCollideObject = -1
    Set pOrigin = New Point
    Set pScaled = MakePoint(1, 1, 1)
    Set pRotate = New Point
    Set pOffset = New Point
    
    Set pDirect = New Point
    Set pTwists = New Point
    Set pScalar = New Point
    
    Set pReplacerVals = New ntnodes10.Collection
    Set pReplacerKeys = New ntnodes10.Collection
   
    Set pAttachments = New ntnodes10.Collection
  
    Set pMotions = New ntnodes10.Collection
    
    'DefaultMatrix
    
    pWireFrame = False
                        
End Sub

Private Sub Class_Terminate()

    If pCollision Or (pCollideIndex > -1) And (pBoundsIndex > 0) Then
        DelCollision Me
    End If
    

    
    If Not pReplacerVals Is Nothing Then
        Do Until pReplacerVals.Count = 0
            pReplacerVals.Remove 1
        Loop
        Set pReplacerVals = Nothing
    End If
    
    If Not pReplacerKeys Is Nothing Then
        Do Until pReplacerKeys.Count = 0
            pReplacerKeys.Remove 1
        Loop
        Set pReplacerKeys = Nothing
    End If
    If Not pMotions Is Nothing Then
        ClearMotions
        Set pMotions = Nothing
    End If
    
    If Not pAttachments Is Nothing Then
        Do Until pAttachments.Count = 0
            pAttachments.Remove 1
        Loop
        Set pAttachments = Nothing
    End If

    Set pOrigin = Nothing
    Set pScaled = Nothing
    Set pRotate = Nothing
    Set pOffset = Nothing
    Set pDirect = Nothing
    Set pTwists = Nothing
    Set pScalar = Nothing
    Stats_Element_Count = Stats_Element_Count - 1
End Sub

