VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Billboard"
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

Private pPoint1 As MyVertex
Private pPoint2 As MyVertex
Private pPoint3 As MyVertex
Private pPoint4 As MyVertex

Private pScaleX As Single
Private pScaleY As Single

Private pTranslucent As Boolean
Private pBlackAlpha As Boolean

Private pAnimateMSecs As Single
Private pAnimateTimer As Double
Private pAnimatePoint As Long

Private pVisible As Boolean

Private pKey As String
Private pSerialize As Boolean
Private pAlign As Long

Public Property Get Point1() As Point
    Set Point1 = MakePoint(pPoint1.X, pPoint1.Y, pPoint1.Z)
End Property
Public Property Set Point1(ByRef RHS As Point)
    With pPoint1
        .X = RHS.X
        .Y = RHS.Y
        .Z = RHS.Z
    End With
    If pAlign < 2 Then pAlign = 1
    RecreateSkins
End Property
Public Property Get Point2() As Point
    Set Point2 = MakePoint(pPoint2.X, pPoint2.Y, pPoint2.Z)
End Property
Public Property Set Point2(ByRef RHS As Point)
    With pPoint2
        .X = RHS.X
        .Y = RHS.Y
        .Z = RHS.Z
    End With
    If pAlign < 2 Then pAlign = 1
    RecreateSkins
End Property
Public Property Get Point3() As Point
    Set Point3 = MakePoint(pPoint3.X, pPoint3.Y, pPoint3.Z)
End Property
Public Property Set Point3(ByRef RHS As Point)
    With pPoint3
        .X = RHS.X
        .Y = RHS.Y
        .Z = RHS.Z
    End With
    If pAlign < 2 Then pAlign = 1
    RecreateSkins
End Property
Public Property Get Point4() As Point
    Set Point4 = MakePoint(pPoint4.X, pPoint4.Y, pPoint4.Z)
End Property
Public Property Set Point4(ByRef RHS As Point)
    With pPoint4
        .X = RHS.X
        .Y = RHS.Y
        .Z = RHS.Z
    End With
    If pAlign < 2 Then pAlign = 1
    RecreateSkins
End Property

Public Property Get Point1UV() As Coord
    Set Point1UV = MakeCoord(pPoint1.tu, pPoint1.tv)
End Property
Public Property Set Point1UV(ByRef RHS As Coord)
    With pPoint1
        .tu = RHS.X
        .tv = RHS.Y
    End With
    pAlign = 2
    RecreateSkins
End Property

Public Property Get Point2UV() As Coord
    Set Point2UV = MakeCoord(pPoint2.tu, pPoint2.tv)
End Property
Public Property Set Point2UV(ByRef RHS As Coord)
    With pPoint2
        .tu = RHS.X
        .tv = RHS.Y
    End With
    pAlign = 2
    RecreateSkins
End Property

Public Property Get Point3UV() As Coord
    Set Point3UV = MakeCoord(pPoint3.tu, pPoint3.tv)
End Property
Public Property Set Point3UV(ByRef RHS As Coord)
    With pPoint3
        .tu = RHS.X
        .tv = RHS.Y
    End With
    pAlign = 2
    RecreateSkins
End Property

Public Property Get Point4UV() As Coord
    Set Point4UV = MakeCoord(pPoint4.tu, pPoint4.tv)
End Property
Public Property Set Point4UV(ByRef RHS As Coord)
    With pPoint4
        .tu = RHS.X
        .tv = RHS.Y
    End With
    pAlign = 2
    RecreateSkins
End Property

Public Property Get Origin() As Point
    Set Origin = pOrigin
End Property
Public Property Set Origin(ByRef RHS As Point)
    Set pOrigin = RHS
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
Public Property Get Translucent() As Boolean
    Translucent = pTranslucent
End Property
Public Property Let Translucent(ByVal RHS As Boolean)
    pTranslucent = RHS
End Property

Public Property Get BlackAlpha() As Boolean
    BlackAlpha = pBlackAlpha
End Property
Public Property Let BlackAlpha(ByVal RHS As Boolean)
    pBlackAlpha = RHS
End Property

Public Property Get Animated() As Single
    Animated = pAnimateMSecs
End Property
Public Property Let Animated(ByVal RHS As Single)
    pAnimateMSecs = RHS
    pAnimateTimer = GetTimer
    pAnimatePoint = 1
End Property

Public Property Get AnimateTimer() As Double
Attribute AnimateTimer.VB_MemberFlags = "40"
    AnimateTimer = pAnimateTimer
End Property
Public Property Let AnimateTimer(ByVal RHS As Double)
    pAnimateTimer = RHS
End Property

Public Property Get AnimatePoint() As Long
Attribute AnimatePoint.VB_MemberFlags = "40"
    AnimatePoint = pAnimatePoint
End Property
Public Property Let AnimatePoint(ByVal RHS As Long)
    pAnimatePoint = RHS
End Property


Public Property Get FileNames() As VBA.Collection
    Set FileNames = pFileNames
End Property
Public Property Get Filename(Optional ByVal Index As Long = 0) As String
    If pFileNames.Count > 0 And Index <= pFileNames.Count And Index > 0 Then
        Filename = pFileNames.Item(Index)
    ElseIf pFileNames.Count > 0 Then
        Filename = pFileNames.Item(1)
    End If
End Property
Public Property Let Filename(Optional ByVal Index As Long = 0, ByVal RHS As String)
    If Index = 0 Then
        If PathExists(ScriptRoot & "\Models\" & RHS, True) Then
            pFileNames.Add RHS
            ReDim Preserve pSkin(1 To pFileNames.Count) As Direct3DTexture8
            Set pSkin(pFileNames.Count) = LoadTexture(ScriptRoot & "\Models\" & RHS)
        Else
            Err.Raise 8, App.EXEName, "File not found."
        End If
    ElseIf pFileNames.Count > 0 And Index <= pFileNames.Count And Index > 0 Then
        If PathExists(ScriptRoot & "\Models\" & RHS, True) Then
            Set pSkin(Index) = Nothing
            Set pSkin(Index) = LoadTexture(ScriptRoot & "\Models\" & RHS)
            pFileNames.Add RHS, , , Index
            pFileNames.Remove Index
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
    Visible = pVisible And (Not (pVBuf Is Nothing))
End Property
Public Property Let Visible(ByVal RHS As Boolean)
    pVisible = RHS
End Property

Public Sub SetTexture(ByVal Stage As Long, ByVal Index As Long)
Attribute SetTexture.VB_MemberFlags = "40"
    If Index = -1 Then
        DDevice.SetTexture Stage, Nothing
    Else
        DDevice.SetTexture Stage, pSkin(Index)
    End If
End Sub
Public Sub Render()
Attribute Render.VB_MemberFlags = "40"
    DDevice.SetStreamSource 0, pVBuf, Len(pPlaq(0))
    DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
End Sub

Private Sub RecreateSkins()
    If pAlign = 1 Then
        CreateSquare pPlaq, 0, MakeVector(pPoint1.X, pPoint1.Y, pPoint1.Z), _
                                        MakeVector(pPoint2.X, pPoint2.Y, pPoint2.Z), _
                                        MakeVector(pPoint3.X, pPoint3.Y, pPoint3.Z), _
                                        MakeVector(pPoint4.X, pPoint4.Y, pPoint4.Z), _
                                        pScaleX, pScaleY
    ElseIf pAlign = 2 Then
        CreateSquareEx pPlaq, 0, pPoint1, pPoint2, pPoint3, pPoint4

    End If

    If pAlign = 1 Or pAlign = 2 Then
        Set pOrigin = ToPoint(SquareCenter(MakeVector(pPoint1.X, pPoint1.Y, pPoint1.Z), _
                                        MakeVector(pPoint2.X, pPoint2.Y, pPoint2.Z), _
                                        MakeVector(pPoint3.X, pPoint3.Y, pPoint3.Z), _
                                        MakeVector(pPoint4.X, pPoint4.Y, pPoint4.Z)))


        D3DVertexBuffer8SetData pVBuf, 0, Len(pPlaq(0)) * 6, 0, pPlaq(0)
    End If
End Sub

Private Sub Class_Initialize()
    Set pOrigin = New Point
    Stats_Board_Count = Stats_Board_Count + 1

    ReDim pPlaq(0 To 5) As MyVertex
    Set pVBuf = DDevice.CreateVertexBuffer(Len(pPlaq(0)) * 6, 0, FVF_RENDER, D3DPOOL_DEFAULT)

    Set pFileNames = New VBA.Collection
    
    Key = "K" & ObjPtr(Me)
End Sub

Private Sub Class_Terminate()
    If pFileNames.Count > 0 Then
        Dim I As Long
        For I = 1 To pFileNames.Count
            Set pSkin(I) = Nothing
        Next
        Erase pSkin
        Do Until pFileNames.Count = 0
            pFileNames.Remove 1
        Loop
    End If
    
    Erase pPlaq
    
    Set pVBuf = Nothing
    
    Set pOrigin = Nothing
    
    Stats_Board_Count = Stats_Board_Count - 1
End Sub

'Public Type MyBoard
'    Skin() As Direct3DTexture8
'    SkinCount As Long
'
'    Plaq() As MyVertex
'    VBuf As Direct3DVertexBuffer8
'
'    Center As D3DVECTOR
'
'    Point1 As MyVertex
'    Point2 As MyVertex
'    Point3 As MyVertex
'    Point4 As MyVertex
'
'    ScaleX As Single
'    ScaleY As Single
'
'    Translucent As Boolean
'    BlackAlpha As Boolean
'
'    AnimateMSecs As Single
'    AnimateTimer As Double
'    AnimatePoint As Long
'
'End Type


'                        o = 0
'                        BillBoardCount = BillBoardCount + 1
'                        ReDim Preserve BillBoards(1 To BillBoardCount) As MyBoard
'                        With BillBoards(BillBoardCount)
'                            ReDim .Plaq(0 To 5) As MyVertex
'
'                            Do Until inData = ""
'                                If (Left(Replace(Replace(inData, " ", ""), vbTab, ""), 1) = ";") Then
'                                    RemoveNextArg inData, vbCrLf
'                                Else
'                                    inName = RemoveLineArg(inData, "[")
'                                    inLine = (NumLines - CountWord(inData & inText, vbCrLf))
'                                    If (Not (Trim(inName) = "")) Then
'                                        inArg() = Split(Trim(inName), " ")
'                                        ReDim Preserve inArg(0 To 5)
'                                        Select Case inArg(0)

'                                            Case "filename"
'                                                .SkinCount = .SkinCount + 1
'                                                ReDim Preserve .Skin(1 To .SkinCount) As Direct3DTexture8
'                                                Set .Skin(.SkinCount) = LoadTexture(AppPath & "Models\" & inArg(1))
'                                            Case "point1"
'                                                .Point1.X = CSng(inArg(1))
'                                                .Point1.Y = CSng(inArg(2))
'                                                .Point1.Z = CSng(inArg(3))
'                                                o = 1
'                                                If IsNumeric(inArg) = 5 Then
'                                                    o = 2
'                                                    .Point1.tu = CSng(inArg(4))
'                                                    .Point1.tv = CSng(inArg(5))
'                                                End If
'                                            Case "point2"
'                                                .Point2.X = CSng(inArg(1))
'                                                .Point2.Y = CSng(inArg(2))
'                                                .Point2.Z = CSng(inArg(3))
'                                                o = 1
'                                                If IsNumeric(inArg) = 5 Then
'                                                    o = 2
'                                                    .Point2.tu = CSng(inArg(4))
'                                                    .Point2.tv = CSng(inArg(5))
'                                                End If
'                                            Case "point3"
'                                                .Point3.X = CSng(inArg(1))
'                                                .Point3.Y = CSng(inArg(2))
'                                                .Point3.Z = CSng(inArg(3))
'                                                o = 1
'                                                If IsNumeric(inArg) = 5 Then
'                                                    o = 2
'                                                    .Point3.tu = CSng(inArg(4))
'                                                    .Point3.tv = CSng(inArg(5))
'                                                End If
'                                            Case "point4"
'                                                .Point4.X = CSng(inArg(1))
'                                                .Point4.Y = CSng(inArg(2))
'                                                .Point4.Z = CSng(inArg(3))
'                                                o = 1
'                                                If IsNumeric(inArg) = 5 Then
'                                                    o = 2
'                                                    .Point4.tu = CSng(inArg(4))
'                                                    .Point4.tv = CSng(inArg(5))
'                                                End If
'                                            Case "scalex"
'                                                .ScaleX = CSng(inArg(1))
'                                                o = 1
'                                            Case "scaley"
'                                                .ScaleY = CSng(inArg(1))

'                                            Case Else
'                                                If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
'                                                    AddMessage "Warning, Unknown Object at Line " & inLine
'                                                End If
'                                        End Select
'                                    ElseIf Left(Trim(inData), 1) = "[" Then
'                                        AddMessage "Warning, Itemless Brackets at Line " & inLine
'                                        GoTo throwerror
'                                    End If
'                                End If
'                            Loop
'
'                            If o = 1 Then
'                                CreateSquare .Plaq, 0, MakeVector(.Point1.X, .Point1.Y, .Point1.Z), _
'                                                                MakeVector(.Point2.X, .Point2.Y, .Point2.Z), _
'                                                                MakeVector(.Point3.X, .Point3.Y, .Point3.Z), _
'                                                                MakeVector(.Point4.X, .Point4.Y, .Point4.Z), _
'                                                                .ScaleX, .ScaleY
'                            ElseIf o = 2 Then
'                                CreateSquareEx .Plaq, 0, .Point1, .Point2, .Point3, .Point4
'
'                            End If
'
'                            If o = 1 Or o = 2 Then
'                                .Center = SquareCenter(MakeVector(.Point1.X, .Point1.Y, .Point1.Z), _
'                                                                MakeVector(.Point2.X, .Point2.Y, .Point2.Z), _
'                                                                MakeVector(.Point3.X, .Point3.Y, .Point3.Z), _
'                                                                MakeVector(.Point4.X, .Point4.Y, .Point4.Z))
'
'                                Set .VBuf = DDevice.CreateVertexBuffer(Len(.Plaq(0)) * 6, 0, FVF_RENDER, D3DPOOL_DEFAULT)
'                                D3DVertexBuffer8SetData .VBuf, 0, Len(.Plaq(0)) * 6, 0, .Plaq(0)
'                            End If
'                        End With
