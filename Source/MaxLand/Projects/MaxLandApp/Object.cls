VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Object"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pOrigin As Point

Private pMeshIndex As Long
Private pVisualIndex As Long

    
Private pKey As String
Private pSerialize As Boolean
Private pVisible As Boolean
Private pGravitational As Boolean
    
Private pBlackAlpha As Boolean
Private pWireFrame As Boolean
    
Private pIsMoving As Moving
Private pInLiquid As Boolean
Private pOnLadder As Boolean


Private pOffset As D3DVECTOR
Private pScaled As D3DVECTOR
Private pRotate As D3DVECTOR


Private pDirect As D3DVECTOR
Private pTwists As D3DVECTOR
Private pScalar As D3DVECTOR
    
    
Private pCollideIndex As Long
Private pCollideObject As Long
Private pCollideFaces As Long
Private pCulledFaces As Long




'Public Type MyObject
'
'    Effect As Collides
'
'    Folcrum() As D3DVECTOR
'    FolcrumCount As Long
'
'    ReplacerVals As New VBA.Collection
'    ReplacerKeys As New VBA.Collection
'
'
'    ActivityCount As Long
'    Activities() As MyActivity
'
'    Matrix As D3DMATRIX
'
'End Type


Public Property Get Origin() As Point
    Set Origin = pOrigin
End Property
Public Property Set Origin(ByRef RHS As Point)
    Set pOrigin = RHS
End Property
Public Property Let Origin(ByRef RHS)
    pOrigin.ToString = RHS
End Property

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
        If pKey <> RHS And pKey <> "" Then
            If All.Exists(pKey) Then All.Remove pKey
            If elements.Exists(pKey) Then elements.Remove pKey
        End If
        All.Add Me, RHS
        elements.Add Me, RHS
        pKey = RHS
    End If
End Property

Public Property Get Mesh() As D3DXMesh

End Property
'
Public Property Get Materials() As D3DMATERIAL8

End Property
Public Property Get Textures() As Direct3DTexture8

End Property
'
Public Property Get Verticies() As D3DVERTEX

End Property
Public Property Get Indicies() As Integer

End Property
Public Property Get VBuffer() As Direct3DVertexBuffer8

End Property
'
Public Property Get MaterialBuffer() As D3DXBuffer

End Property

'
'                        If (inItem = "player") Then
'                            Player.CameraAngle = 0
'                            Player.CameraPitch = 0
'                            Player.MoveSpeed = 0.3
'                            Space.Boundary = 90
'                        End If
'

'                                        Case "visible"
'                                            If inArg(1) = "" Then
'                                                NewObj.Visible = True
'                                            Else
'                                                NewObj.Visible = CBool(inArg(1))
'                                            End If
'                                        Case "hidden"
'                                            NewObj.Visible = False
'                                        Case "wireframe"
'                                            NewObj.WireFrame = True
'                                        Case "indoorcollide"
'                                            NewObj.Effect = Collides.InDoor
'                                        Case "groundcollide"
'                                            NewObj.Effect = Collides.Ground
'                                        Case "liquidcollide"
'                                            NewObj.Effect = Collides.Liquid
'                                        Case "laddercollide"
'                                            NewObj.Effect = Collides.Ladder
'                                        Case "gravitational"
'                                            NewObj.Gravitational = True
'                                        Case "key"
'                                            NewObj.Key = inArg(1)
'                                        Case "blackalpha"
'                                            NewObj.BlackAlpha = True
'                                        Case "nocollision"
'                                            NewObj.CollideIndex = -1
'                                        Case "origin", "location", "position"
'                                            NewObj.Origin = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
'                                        Case "scale"
'                                            NewObj.Scaled = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
'                                        Case "rotate"
'                                            NewObj.Rotate = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
'                                        Case "offset"
'                                            NewObj.Offset = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
'                                        Case "folcrum"
'                                            NewObj.FolcrumCount = NewObj.FolcrumCount + 1
'                                            ReDim Preserve NewObj.Folcrum(1 To NewObj.FolcrumCount) As D3DVECTOR
'                                            NewObj.Folcrum(NewObj.FolcrumCount) = MakeVector(CSng(inArg(1)), CSng(inArg(2)), CSng(inArg(3)))
'                                        Case "visualobj"
'                                            If MeshCount = 0 Then
'                                                MeshCount = MeshCount + 1
'                                                ReDim Meshes(1 To MeshCount) As MyMesh
'                                                NewObj.VisualIndex = MeshCount
'                                                Meshes(NewObj.VisualIndex).Filename = LCase(inArg(1))
'                                                If PathExists(AppPath & "Models\" & Meshes(NewObj.VisualIndex).Filename, True) Then
'                                                    CreateMesh AppPath & "Models\" & Meshes(NewObj.VisualIndex).Filename, Meshes(NewObj.VisualIndex).Mesh, Meshes(NewObj.VisualIndex).MaterialBuffer, _
'                                                            NewObj.Origin, NewObj.Scaled, Meshes(NewObj.VisualIndex).Materials, Meshes(NewObj.VisualIndex).Textures, _
'                                                            Meshes(NewObj.VisualIndex).Verticies, Meshes(NewObj.VisualIndex).Indicies, Meshes(NewObj.VisualIndex).MaterialCount
'                                                Else
'                                                    ReDim Meshes(NewObj.VisualIndex).Textures(0 To 0) As Direct3DTexture8
'                                                    ReDim Meshes(NewObj.VisualIndex).Materials(0 To 0) As D3DMATERIAL8
'                                                    NewObj.VisualIndex = 0
'                                                End If
'                                            Else
'                                                For I = LBound(Meshes) To UBound(Meshes)
'                                                    If Meshes(I).Filename = LCase(inArg(1)) Then
'                                                        NewObj.VisualIndex = I
'                                                        Exit For
'                                                    End If
'                                                Next
'                                                If NewObj.VisualIndex = 0 Then
'                                                    MeshCount = MeshCount + 1
'                                                    ReDim Preserve Meshes(1 To MeshCount) As MyMesh
'                                                    NewObj.VisualIndex = MeshCount
'                                                    Meshes(NewObj.VisualIndex).Filename = LCase(inArg(1))
'                                                    If PathExists(AppPath & "Models\" & Meshes(NewObj.VisualIndex).Filename, True) Then
'                                                        CreateMesh AppPath & "Models\" & Meshes(NewObj.VisualIndex).Filename, Meshes(NewObj.VisualIndex).Mesh, Meshes(NewObj.VisualIndex).MaterialBuffer, _
'                                                                NewObj.Origin, NewObj.Scaled, Meshes(NewObj.VisualIndex).Materials, Meshes(NewObj.VisualIndex).Textures, _
'                                                                Meshes(NewObj.VisualIndex).Verticies, Meshes(NewObj.VisualIndex).Indicies, Meshes(NewObj.VisualIndex).MaterialCount
'                                                    Else
'                                                        ReDim Meshes(NewObj.VisualIndex).Textures(0 To 0) As Direct3DTexture8
'                                                        ReDim Meshes(NewObj.VisualIndex).Materials(0 To 0) As D3DMATERIAL8
'                                                        NewObj.VisualIndex = 0
'                                                    End If
'                                                End If
'                                            End If
'
'
'                                        Case "filename", "boundsobj"
'
'                                            If MeshCount = 0 Then
'                                                MeshCount = MeshCount + 1
'                                                ReDim Meshes(1 To MeshCount) As MyMesh
'                                                NewObj.MeshIndex = MeshCount
'                                                Meshes(NewObj.MeshIndex).Filename = LCase(inArg(1))
'                                                If PathExists(AppPath & "Models\" & Meshes(NewObj.MeshIndex).Filename, True) Then
'                                                    CreateMesh AppPath & "Models\" & Meshes(NewObj.MeshIndex).Filename, Meshes(NewObj.MeshIndex).Mesh, Meshes(NewObj.MeshIndex).MaterialBuffer, _
'                                                            NewObj.Origin, NewObj.Scaled, Meshes(NewObj.MeshIndex).Materials, Meshes(NewObj.MeshIndex).Textures, _
'                                                            Meshes(NewObj.MeshIndex).Verticies, Meshes(NewObj.MeshIndex).Indicies, Meshes(NewObj.MeshIndex).MaterialCount
'                                                Else
'                                                    ReDim Meshes(NewObj.MeshIndex).Textures(0 To 0) As Direct3DTexture8
'                                                    ReDim Meshes(NewObj.MeshIndex).Materials(0 To 0) As D3DMATERIAL8
'                                                    NewObj.MeshIndex = 0
'                                                End If
'                                            Else
'                                                For I = LBound(Meshes) To UBound(Meshes)
'                                                    If Meshes(I).Filename = LCase(inArg(1)) Then
'                                                        NewObj.MeshIndex = I
'                                                        Exit For
'                                                    End If
'                                                Next
'
'                                                If NewObj.MeshIndex = 0 Then
'                                                    MeshCount = MeshCount + 1
'                                                    ReDim Preserve Meshes(1 To MeshCount) As MyMesh
'                                                    NewObj.MeshIndex = MeshCount
'                                                    Meshes(NewObj.MeshIndex).Filename = LCase(inArg(1))
'                                                    If PathExists(AppPath & "Models\" & Meshes(NewObj.MeshIndex).Filename, True) Then
'                                                        CreateMesh AppPath & "Models\" & Meshes(NewObj.MeshIndex).Filename, Meshes(NewObj.MeshIndex).Mesh, Meshes(NewObj.MeshIndex).MaterialBuffer, _
'                                                                NewObj.Origin, NewObj.Scaled, Meshes(NewObj.MeshIndex).Materials, Meshes(NewObj.MeshIndex).Textures, _
'                                                                Meshes(NewObj.MeshIndex).Verticies, Meshes(NewObj.MeshIndex).Indicies, Meshes(NewObj.MeshIndex).MaterialCount
'                                                    Else
'                                                        ReDim Meshes(NewObj.MeshIndex).Textures(0 To 0) As Direct3DTexture8
'                                                        ReDim Meshes(NewObj.MeshIndex).Materials(0 To 0) As D3DMATERIAL8
'                                                        NewObj.MeshIndex = 0
'                                                    End If
'
'                                                End If
'
'                                            End If
'                                        Case "replacer"
'                                            If BillBoardCount > 0 Then
'                                                For o = 1 To BillBoardCount
'                                                    If BillBoards(o).Key = inArg(2) Then
'                                                        NewObj.ReplacerVals.Add o, NewObj.Key & "_" & Replace(inArg(1), ".", "")
'                                                        NewObj.ReplacerKeys.Add NewObj.Key & "_" & Replace(inArg(1), ".", "")
'                                                    End If
'                                                Next
'                                            End If
'                                        Case "activity"
'
'                                            Select Case LCase(CStr(inArg(1)))
'                                                Case "direct"
'                                                    AddActivity NewObj, Actions.Directing, inArg(2), makepoint(CSng(inArg(3)), CSng(inArg(4)), CSng(inArg(5))), _
'                                                                            IIf(IsNumeric(inArg(6)), inArg(6), 0), IIf(IsNumeric(inArg(7)), inArg(7), 0), _
'                                                                            IIf(IsNumeric(inArg(8)), inArg(8), -1), IIf(IsNumeric(inArg(9)), inArg(9), -1)
'                                                Case "rotate"
'
'                                                    AddActivity NewObj, Actions.Rotating, inArg(2), makepoint(CSng(inArg(3)), CSng(inArg(4)), CSng(inArg(5))), _
'                                                                            IIf(IsNumeric(inArg(6)), inArg(6), 0), IIf(IsNumeric(inArg(7)), inArg(7), 0), _
'                                                                            IIf(IsNumeric(inArg(8)), inArg(8), -1), IIf(IsNumeric(inArg(9)), inArg(9), -1)
'                                                Case "scale"
'                                                    AddActivity NewObj, Actions.Scaling, inArg(2), makepoint(CSng(inArg(3)), CSng(inArg(4)), CSng(inArg(5))), _
'                                                                            IIf(IsNumeric(inArg(6)), inArg(6), 0), IIf(IsNumeric(inArg(7)), inArg(7), 0), _
'                                                                            IIf(IsNumeric(inArg(8)), inArg(8), -1), IIf(IsNumeric(inArg(9)), inArg(9), -1)
'                                                Case "script"
'                                                    If InStr(inData, "[") > 0 Then
'                                                        inLine = inLine + CountWord(RemoveLineArg(inData, "["), vbCrLf)
'                                                        inTrig = RemoveQuotedArg(inData, "[", "]", True)
'                                                        inTrig = inLine & ":" & inTrig
'
'                                                        AddActivity NewObj, Actions.Script, inArg(2), makepoint(0, 0, 0), _
'                                                                                , , IIf(IsNumeric(inArg(3)), inArg(3), -1), _
'                                                                                IIf(IsNumeric(inArg(4)), inArg(4), -1), inTrig
'                                                    Else
'                                                        AddMessage "Warning, Brackets Required at Line " & inLine
'                                                    End If
'
'                                            End Select
'                                        Case "boundary"
'                                            If inItem = "player" Then Space.Boundary = CSng(inArg(1))
'                                        Case "movespeed"
'                                            If inItem = "player" Then Player.MoveSpeed = CSng(inArg(1))
'                                        Case "camerapitch"
'                                            If inItem = "player" Then Player.CameraPitch = CSng(inArg(1))
'                                        Case "cameraangle"
'                                            If inItem = "player" Then Player.CameraAngle = CSng(inArg(1))
'
'                                        Case Else
'                                            If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
'                                                AddMessage "Warning, Unknown Object at Line " & inLine
'                                            End If
'                                    End Select
'                                ElseIf Left(Trim(inData), 1) = "[" Then
'                                    AddMessage "Warning, Itemless Brackets at Line " & inLine
'                                    GoTo throwerror
'                                End If
'                            End If
'                        Loop
'
'                        If (NewObj.MeshIndex > 0) Then
'
'                            D3DXMatrixIdentity NewObj.Matrix
'                            D3DXMatrixTranslation NewObj.Matrix, NewObj.Offset.X, NewObj.Offset.Y, NewObj.Offset.Z
'                            D3DXMatrixRotationX NewObj.Matrix, NewObj.Rotate.X * (PI / 180)
'                            D3DXMatrixRotationY NewObj.Matrix, NewObj.Rotate.Y * (PI / 180)
'                            D3DXMatrixRotationZ NewObj.Matrix, NewObj.Rotate.Z * (PI / 180)
'                            D3DXMatrixScaling NewObj.Matrix, NewObj.Scaled.X, NewObj.Scaled.Y, NewObj.Scaled.Z
'
'                            D3DXMatrixIdentity matWorld
'                            DDevice.SetTransform D3DTS_WORLD, matWorld
'
'                        End If
'
'                        If (NewObj.CollideIndex > -1) And (NewObj.MeshIndex > 0) Then
'                            AddCollision NewObj
'                        Else
'                            NewObj.CollideIndex = -1
'                        End If
'
'                        If (inItem = "player") Then
'                            PlayerObject = NewObj
'                        Else
'                            ObjectCount = ObjectCount + 1
'                            ReDim Preserve Objects(1 To ObjectCount) As MyObject
'                            Objects(ObjectCount) = NewObj
'                        End If


'    Key As String
'    Visible As Boolean
'
'    MeshIndex As Long
'    VisualIndex As Long
'
'    BlackAlpha As Boolean
'    WireFrame As Boolean
'
'    Effect As Collides
'
'    Origin As D3DVECTOR
'    Direct As D3DVECTOR
'    Offset As D3DVECTOR
'
'    Folcrum() As D3DVECTOR
'    FolcrumCount As Long
'
'    Rotate As D3DVECTOR
'    Twists As D3DVECTOR
'
'    Scaled As D3DVECTOR
'    Scalar As D3DVECTOR
'
'    Activities() As MyActivity
'    ActivityCount As Long
'    Gravitational As Boolean
'
'    ReplacerVals As New VBA.Collection
'    ReplacerKeys As New VBA.Collection
'
'
'    '
'
'
'    CollideIndex As Long
'    CollideObject As Long
'    CollideFaces As Long
'    CulledFaces As Long
'
'    Matrix As D3DMATRIX
'
'    'Public Type MyStates
'    IsMoving As Moving
'    InLiquid As Boolean
'    OnLadder As Boolean
'    'End Type
    
Private Sub Class_Initialize()
    Stats_Object_Count = Stats_Object_Count + 1
    Key = "K" & ObjPtr(Me)
'                        NewObj.VisualIndex = 0
'                        NewObj.MeshIndex = 0
'                        NewObj.Effect = Collides.None
'                        NewObj.Gravitational = False
'                        NewObj.BlackAlpha = False
'                        NewObj.CollideIndex = 0
'                        NewObj.CollideObject = -1
'                        NewObj.Key = ""
'                        NewObj.WireFrame = False
'
'                        Do Until NewObj.ActivityCount = 0
'                            DeleteActivity NewObj, NewObj.Activities(1).Ident
'                        Loop
'                        NewObj.Origin = MakeVector(0, 0, 0)
                        
End Sub

Private Sub Class_Terminate()
    Stats_Object_Count = Stats_Object_Count - 1
End Sub
