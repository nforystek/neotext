VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Board"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pFileNames As VBA.Collection
Private pSkin() As Direct3DTexture8


Private pPlaq() As MyVertex
Private pVBuf As Direct3DVertexBuffer8

Private pOrigin As Point

Private pPoint1 As Point
Private pPoint2 As Point
Private pPoint3 As Point
Private pPoint4 As Point

Private pNormal1 As Point
Private pNormal2 As Point
Private pNormal3 As Point
Private pNormal4 As Point

Private pTexture1 As Coord
Private pTexture2 As Coord
Private pTexture3 As Coord
Private pTexture4 As Coord

Private pScaleX As Single
Private pScaleY As Single

Private pTranslucent As Boolean
Private pAlphablend As Boolean

Private pAnimateMSecs As Single
Private pAnimateTimer As Double
Private pAnimatePoint As Long

Private pVisible As Boolean
Private pMirror As Boolean

Private pKey As String
Private pSerialize As Boolean
Private pAlign As Long

Public Property Get Point1() As Point
    Set Point1 = pPoint1
End Property
Public Property Set Point1(ByRef RHS As Point)
    Set pPoint1 = RHS
    If pAlign < 2 Then pAlign = 1
    RecreateSkins
End Property
Public Property Get Point2() As Point
    Set Point2 = pPoint2
End Property
Public Property Set Point2(ByRef RHS As Point)
    Set pPoint2 = RHS
    If pAlign < 2 Then pAlign = 1
    RecreateSkins
End Property
Public Property Get Point3() As Point
    Set Point3 = pPoint3
End Property
Public Property Set Point3(ByRef RHS As Point)
    Set pPoint3 = RHS
    If pAlign < 2 Then pAlign = 1
    RecreateSkins
End Property
Public Property Get Point4() As Point
    Set Point4 = pPoint4
End Property
Public Property Set Point4(ByRef RHS As Point)
    Set pPoint4 = RHS
    If pAlign < 2 Then pAlign = 1
    RecreateSkins
End Property



Public Property Get Normal1() As Point
    Set Normal1 = pNormal1
End Property
Public Property Set Normal1(ByRef RHS As Point)
    Set pNormal1 = RHS
    pAlign = 2
    RecreateSkins
End Property
Public Property Get Normal2() As Point
    Set Normal2 = pNormal2
End Property
Public Property Set Normal2(ByRef RHS As Point)
    Set pNormal2 = RHS
    pAlign = 2
    RecreateSkins
End Property
Public Property Get Normal3() As Point
    Set Normal3 = pNormal3
End Property
Public Property Set Normal3(ByRef RHS As Point)
    Set pNormal3 = RHS
    pAlign = 2
    RecreateSkins
End Property
Public Property Get Normal4() As Point
    Set Normal4 = pNormal4
End Property
Public Property Set Normal4(ByRef RHS As Point)
    Set pNormal4 = RHS
    pAlign = 2
    RecreateSkins
End Property


Public Property Get Texture1() As Coord
    Set Texture1 = pTexture1
End Property
Public Property Set Texture1(ByRef RHS As Coord)
    Set pTexture1 = RHS
    pAlign = 2
    RecreateSkins
End Property

Public Property Get Texture2() As Coord
    Set Texture2 = pTexture2
End Property
Public Property Set Texture2(ByRef RHS As Coord)
    Set pTexture2 = RHS
    pAlign = 2
    RecreateSkins
End Property

Public Property Get Texture3() As Coord
    Set Texture3 = pTexture3
End Property
Public Property Set Texture3(ByRef RHS As Coord)
    Set pTexture3 = RHS
    pAlign = 2
    RecreateSkins
End Property

Public Property Get Texture4() As Coord
    Set Texture4 = pTexture4
End Property
Public Property Set Texture4(ByRef RHS As Coord)
    Set pTexture4 = RHS
    pAlign = 2
    RecreateSkins
End Property

Public Property Get Origin() As Point
    Set Origin = pOrigin
End Property
Public Property Set Origin(ByRef RHS As Point)
    Set pOrigin = RHS
    RecreateSkins
End Property

Public Property Get ScaleX() As Single
    ScaleX = pScaleX
End Property
Public Property Let ScaleX(ByVal RHS As Single)
    pScaleX = RHS
    pAlign = 1
    RecreateSkins
End Property

Public Property Get ScaleY() As Single
    ScaleY = pScaleY
End Property
Public Property Let ScaleY(ByVal RHS As Single)
    pScaleY = RHS
    pAlign = 1
    RecreateSkins
End Property

Public Property Get Mirror() As Boolean
    Mirror = pMirror
End Property
Public Property Let Mirror(ByVal RHS As Boolean)
    pMirror = RHS
    RecreateSkins
End Property
Public Property Get Translucent() As Boolean
    Translucent = pTranslucent
End Property
Public Property Let Translucent(ByVal RHS As Boolean)
    pTranslucent = RHS
End Property

Public Property Get Alphablend() As Boolean
    Alphablend = pAlphablend
End Property
Public Property Let Alphablend(ByVal RHS As Boolean)
    pAlphablend = RHS
End Property

Public Property Get Animated() As Single
    Animated = pAnimateMSecs
End Property
Public Property Let Animated(ByVal RHS As Single)
    pAnimateMSecs = RHS
    pAnimateTimer = GetTimer
    pAnimatePoint = 1
End Property

Friend Property Get AnimateTimer() As Double
    AnimateTimer = pAnimateTimer
End Property

Friend Property Let AnimateTimer(ByVal RHS As Double)
    pAnimateTimer = RHS
End Property

Friend Property Get AnimatePoint() As Long
    AnimatePoint = pAnimatePoint
End Property
Friend Property Let AnimatePoint(ByVal RHS As Long)
    pAnimatePoint = RHS
End Property


Public Property Get FileNames() As VBA.Collection
    Set FileNames = pFileNames
End Property
Public Property Get FileName(Optional ByVal Index As Long = 0) As String
    If pFileNames.Count > 0 And Index <= pFileNames.Count And Index > 0 Then
        FileName = pFileNames.Item(Index)
    ElseIf pFileNames.Count > 0 Then
        FileName = pFileNames.Item(1)
    End If
End Property
Public Property Let FileName(Optional ByVal Index As Long = 0, ByVal RHS As String)
    If Index = 0 Then
        If PathExists(ScriptRoot & "Models\" & RHS, True) Then
            pFileNames.Add RHS
            ReDim Preserve pSkin(1 To pFileNames.Count) As Direct3DTexture8
            Set pSkin(pFileNames.Count) = LoadTexture(ScriptRoot & "Models\" & RHS)
            RecreateSkins
        Else
            Err.Raise 8, App.EXEName, "File not found."
        End If
    ElseIf pFileNames.Count > 0 And Index <= pFileNames.Count And Index > 0 Then
        If PathExists(ScriptRoot & "Models\" & RHS, True) Then
            Set pSkin(Index) = Nothing
            Set pSkin(Index) = LoadTexture(ScriptRoot & "Models\" & RHS)
            pFileNames.Add RHS, , , Index
            pFileNames.Remove Index
            RecreateSkins
        Else
            Err.Raise 8, App.EXEName, "File not found."
        End If
    End If
    
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
        If pKey <> RHS And pKey <> "" Then If All.Exists(pKey) Then All.Remove pKey
        If All.Exists(RHS) Then
            All.Remove RHS
            All.Add Me, RHS
        End If
        pKey = RHS
    End If
End Property

Public Property Get Visible() As Boolean
    Visible = pVisible And ((Not pVBuf Is Nothing) Or pMirror)
End Property
Public Property Let Visible(ByVal RHS As Boolean)
    pVisible = RHS
End Property

Public Sub SetTexture(ByVal Stage As Long, ByVal Index As Long)
    If Index = -1 Then
        DDevice.SetTexture Stage, Nothing
    Else
        If FileNames.Count > 0 Then
            DDevice.SetTexture Stage, pSkin(Index)
        End If
    End If
End Sub
Public Sub Render()
        
    DDevice.SetStreamSource 0, pVBuf, LenB(pPlaq(0))
    DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2

End Sub

Private Function MakeMyVertex(ByRef Point As Point, ByRef Normal As Point, ByRef Texture As Coord) As MyVertex
    With MakeMyVertex
        .X = Point.X
        .Y = Point.Y
        .Z = Point.Z
        .nx = Normal.X
        .ny = Normal.Y
        .nz = Normal.Z
        .tu = Texture.U
        .tv = Texture.V
    End With
End Function


Private Sub RecreateSkins()


    If pAlign = 2 Then
        CreateSquareEx pPlaq, 0, MakeMyVertex(Point1, Normal1, Texture1), _
                                 MakeMyVertex(Point2, Normal2, Texture2), _
                                 MakeMyVertex(Point3, Normal3, Texture3), _
                                MakeMyVertex(Point4, Normal4, Texture4)
    Else 'If pAlign = 1 Then
        CreateSquare pPlaq, 0, MakeVector(pPoint1.X, pPoint1.Y, pPoint1.Z), _
                                        MakeVector(pPoint2.X, pPoint2.Y, pPoint2.Z), _
                                        MakeVector(pPoint3.X, pPoint3.Y, pPoint3.Z), _
                                        MakeVector(pPoint4.X, pPoint4.Y, pPoint4.Z), _
                                        pScaleX, pScaleY
        pNormal1.X = pPlaq(0).nx
        pNormal1.Y = pPlaq(0).ny
        pNormal1.Z = pPlaq(0).nz
        pTexture1.U = 0
        pTexture1.V = pPlaq(0).tv

        pNormal2.X = pPlaq(1).nx
        pNormal2.Y = pPlaq(1).ny
        pNormal2.Z = pPlaq(1).nz
        pTexture2.U = pPlaq(1).tu
        pTexture2.V = pPlaq(1).tv
        
        pNormal3.X = pPlaq(4).nx
        pNormal3.Y = pPlaq(4).ny
        pNormal3.Z = pPlaq(4).nz
        pTexture3.U = pPlaq(4).tu
        pTexture3.V = 0
        
        pNormal4.X = pPlaq(5).nx
        pNormal4.Y = pPlaq(5).ny
        pNormal4.Z = pPlaq(5).nz
        pTexture4.U = 0
        pTexture4.V = 0
        
    End If

    'If pAlign = 1 Or pAlign = 2 Then
        Set pOrigin = SquareCenter(MakePoint(pPoint1.X, pPoint1.Y, pPoint1.Z), _
                                        MakePoint(pPoint2.X, pPoint2.Y, pPoint2.Z), _
                                        MakePoint(pPoint3.X, pPoint3.Y, pPoint3.Z), _
                                        MakePoint(pPoint4.X, pPoint4.Y, pPoint4.Z))


        
        D3DVertexBuffer8SetData pVBuf, 0, LenB(pPlaq(0)) * (UBound(pPlaq) + 1), 0, pPlaq(0)

    'End If
End Sub

Private Sub Class_Initialize()
    Set pOrigin = New Point

    Set pPoint1 = New Point
    Set pPoint2 = New Point
    Set pPoint3 = New Point
    Set pPoint4 = New Point

    Set pNormal1 = New Point
    Set pNormal2 = New Point
    Set pNormal3 = New Point
    Set pNormal4 = New Point

    Set pTexture1 = New Coord
    Set pTexture2 = New Coord
    Set pTexture3 = New Coord
    Set pTexture4 = New Coord

    Stats_Board_Count = Stats_Board_Count + 1
    
    pAlign = 1
    
    ReDim pPlaq(0 To 5) As MyVertex

    Set pFileNames = New VBA.Collection
    
    Set pVBuf = DDevice.CreateVertexBuffer(LenB(pPlaq(0)) * (UBound(pPlaq) + 1), 0, FVF_RENDER, D3DPOOL_DEFAULT)
    
    Key = "K" & ObjPtr(Me)
End Sub

Private Sub Class_Terminate()
    If pFileNames.Count > 0 Then
        Dim i As Long
        For i = 1 To pFileNames.Count
            Set pSkin(i) = Nothing
        Next
        Erase pSkin
        Do Until pFileNames.Count = 0
            pFileNames.Remove 1
        Loop
    End If
    
    Erase pPlaq
    
    Set pVBuf = Nothing
    Set pPoint1 = Nothing
    Set pPoint2 = Nothing
    Set pPoint3 = Nothing
    Set pPoint4 = Nothing

    Set pNormal1 = Nothing
    Set pNormal2 = Nothing
    Set pNormal3 = Nothing
    Set pNormal4 = Nothing

    Set pTexture1 = Nothing
    Set pTexture2 = Nothing
    Set pTexture3 = Nothing
    Set pTexture4 = Nothing
    
    Set pOrigin = Nothing
    
    Stats_Board_Count = Stats_Board_Count - 1
End Sub

