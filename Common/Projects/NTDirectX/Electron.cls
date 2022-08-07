VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Electron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'##################################################################################################################
'## Electron Description: Anything auditable but not with in the 3D world by itself, such as texture/mesh/sounds ##
'##################################################################################################################

Private pFaceIndex As Long

Private pForm As ElectronTypes

Private WithEvents pCoordinate As Coordinate
Attribute pCoordinate.VB_VarHelpID = -1

Private WithEvents pCenter As Vector
Attribute pCenter.VB_VarHelpID = -1

Private WithEvents pPoint1 As Textcoord
Attribute pPoint1.VB_VarHelpID = -1
Private WithEvents pPoint2 As Textcoord
Attribute pPoint2.VB_VarHelpID = -1
Private WithEvents pPoint3 As Textcoord
Attribute pPoint3.VB_VarHelpID = -1
Private WithEvents pPoint4 As Textcoord
Attribute pPoint4.VB_VarHelpID = -1

Private WithEvents pScaled As Coordinate
Attribute pScaled.VB_VarHelpID = -1

Private Area As Double

Public Visible As Boolean

Private pTransposing As Molecule
Public Translucent As Boolean
Public Transparent As Boolean

Private pAnimateMSecs As Single
Private pAnimateTimer As Single
Private pAnimatePoint As Long

Private pSerialize As Boolean


Public Property Set Transposing(ByRef RHS As Molecule)
    Set pTransposing = RHS
End Property
Public Property Get Transposing() As Molecule
    Set Transposing = pTransposing
End Property

Public Property Get Serialize() As Boolean
    Serialize = pSerialize
End Property
Public Property Let Serialize(ByVal RHS As Boolean)
    pSerialize = RHS
End Property

Public Property Get ToString(Optional ByVal Indent As String = "") As String

    ToString = ToString & Indent & "<Form>" & pForm & "</Form>" & vbCrLf
    
    ToString = ToString & Indent & "<Visible>" & Visible & "</Visible>" & vbCrLf
    ToString = ToString & Indent & "<Translucent>" & Translucent & "</Translucent>" & vbCrLf
    ToString = ToString & Indent & "<Transparent>" & Transparent & "</Transparent>" & vbCrLf

    ToString = ToString & Indent & "<Animated>" & Animated & "</Animated>" & vbCrLf
    ToString = ToString & Indent & "<AnimateTimer>" & AnimateTimer & "</AnimateTimer>" & vbCrLf
    ToString = ToString & Indent & "<AnimatePoint>" & AnimatePoint & "</AnimatePoint>" & vbCrLf

    If Not pCoordinate Is Nothing Then ToString = ToString & Indent & "<Coordinate>" & pCoordinate.ToString() & "</Coordinate>" & vbCrLf
    If Not pCenter Is Nothing Then ToString = ToString & Indent & "<Center>" & pCenter.ToString() & "</Center>" & vbCrLf
    If Not pScaled Is Nothing Then ToString = ToString & Indent & "<Scaled>" & pScaled.ToString() & "</Scaled>" & vbCrLf
    If Not pPoint1 Is Nothing Then ToString = ToString & Indent & "<Point1>" & pPoint1.ToString() & "</Point1>" & vbCrLf
    If Not pPoint2 Is Nothing Then ToString = ToString & Indent & "<Point2>" & pPoint2.ToString() & "</Point2>" & vbCrLf
    If Not pPoint3 Is Nothing Then ToString = ToString & Indent & "<Point3>" & pPoint3.ToString() & "</Point3>" & vbCrLf
    If Not pPoint4 Is Nothing Then ToString = ToString & Indent & "<Point4>" & pPoint4.ToString() & "</Point4>" & vbCrLf

End Property


Public Property Get Form() As Long
    
    If ((pForm And TwoDimensions) = TwoDimensions) Then
        If Not (pCoordinate Is Nothing) Then
            Form = TwoDimensions
        Else
            Form = NotDisplayed
        End If
    End If
    If ((pForm And ThreeDimensions) = ThreeDimensions) Then
        If Not (pPoint1 Is Nothing Or pPoint2 Is Nothing Or pPoint3 Is Nothing Or pPoint4 Is Nothing) Then
            Form = Form Or ThreeDimensions
        Else
            Form = Form Or NotDisplayed
        End If
    End If
    
End Property
Public Property Let Form(ByVal RHS As Long)
    If Not ((pForm And RHS) = RHS) Then
        pForm = pForm Or RHS
        With Faces(pFaceIndex)
            If RHS = TwoDimensions Then
            
                ReDim .Screen2D(0 To 3) As MyScreen
                .Screen2D(0) = MakeScreen(0, 0, -1, 0, 0)
                .Screen2D(1) = MakeScreen(0, 0, -1, 1, 0)
                .Screen2D(2) = MakeScreen(0, 0, -1, 0, 1)
                .Screen2D(3) = MakeScreen(0, 0, -1, 1, 1)

            ElseIf RHS = ThreeDimensions Then
                
                ReDim .Verticies(0 To 5) As MyVertex

            End If
        End With
        ReSquare
    End If

End Property

Public Property Get SurfaceArea() As Double
    SurfaceArea = Area
End Property

Public Property Get Animated() As Double
    Animated = pAnimateMSecs
End Property
Public Property Let Animated(ByVal RHS As Double)
    pAnimateMSecs = RHS
    pAnimateTimer = 0
End Property

Public Property Get AnimateTimer() As Double
    AnimateTimer = pAnimateTimer
End Property
Public Property Let AnimateTimer(ByVal RHS As Double)
    pAnimateTimer = RHS
End Property
Public Property Get AnimatePoint() As Double
    If pAnimateMSecs = 0 Then
        AnimatePoint = 1
    Else
        AnimatePoint = pAnimatePoint
    End If
End Property
Public Property Let AnimatePoint(ByVal RHS As Double)
    If RHS > Faces(pFaceIndex).ImageCount Then
        pAnimatePoint = 1
    Else
        pAnimatePoint = RHS
    End If
End Property

Public Property Get Coordinate() As Coordinate
    If pCoordinate Is Nothing Then Set pCoordinate = New Coordinate
    Form = TwoDimensions
    Set Coordinate = pCoordinate
End Property
Public Property Let Coordinate(ByRef RHS)
    If pCoordinate Is Nothing Then Set pCoordinate = New Coordinate
    Form = TwoDimensions
    pCoordinate = RHS
End Property

Public Property Get Height() As Long
    Height = Faces(pFaceIndex).Dimension.Height
End Property
Public Property Get Width() As Long
    Width = Faces(pFaceIndex).Dimension.Width
End Property

Public Property Get Center() As Vector
    If pCenter Is Nothing Then Set pCenter = New Vector
    Form = ThreeDimensions
    Set Center = pCenter
End Property
Public Property Let Center(ByRef RHS)
    If pCenter Is Nothing Then Set pCenter = New Vector
    Form = ThreeDimensions
    pCenter = RHS
End Property

Public Property Get Scaled() As Coordinate
    If pScaled Is Nothing Then Set pScaled = New Coordinate
    Form = ThreeDimensions
    Set Scaled = pScaled
End Property
Public Property Let Scaled(ByRef RHS)
    If pScaled Is Nothing Then Set pScaled = New Coordinate
    Form = ThreeDimensions
    pScaled = RHS
End Property

Public Property Get Point1() As Textcoord
    If pPoint1 Is Nothing Then Set pPoint1 = New Textcoord
    Form = ThreeDimensions
    Set Point1 = pPoint1
End Property
Public Property Let Point1(ByRef RHS)
    If pPoint1 Is Nothing Then Set pPoint1 = New Textcoord
    Form = ThreeDimensions
    pPoint1 = RHS
End Property
Public Property Get Point2() As Textcoord
    If pPoint2 Is Nothing Then Set pPoint2 = New Textcoord
    Form = ThreeDimensions
    Set Point2 = pPoint2
End Property
Public Property Let Point2(ByRef RHS)
    If pPoint2 Is Nothing Then Set pPoint2 = New Textcoord
    Form = ThreeDimensions
    pPoint2 = RHS
End Property
Public Property Get Point3() As Textcoord
    If pPoint3 Is Nothing Then Set pPoint3 = New Textcoord
    Form = ThreeDimensions
    Set Point3 = pPoint3
End Property
Public Property Let Point3(ByRef RHS)
    If pPoint3 Is Nothing Then Set pPoint3 = New Textcoord
    Form = ThreeDimensions
    pPoint3 = RHS
End Property
Public Property Get Point4() As Textcoord
    If pPoint4 Is Nothing Then Set pPoint4 = New Textcoord
    Form = ThreeDimensions
    Set Point4 = pPoint4
End Property
Public Property Let Point4(ByRef RHS)
    If pPoint4 Is Nothing Then Set pPoint4 = New Textcoord
    Form = ThreeDimensions
    pPoint4 = RHS
End Property

Public Property Get FaceIndex() As Long
    FaceIndex = pFaceIndex
End Property
Friend Property Let FaceIndex(ByVal RHS As Long)
    pFaceIndex = RHS
End Property

Public Property Get File(Optional ByVal Number As Long = 0) As Variant
    If Number = 0 And Faces(FaceIndex).ImageCount > 0 Then
        File = GetIndexFile(Faces(FaceIndex).Images(1))
    Else
        File = GetIndexFile(Faces(FaceIndex).Images(Number))
    End If
End Property
Public Property Let File(Optional ByVal Number As Long = 0, ByVal RHS As Variant)
    If PathExists(ScriptRoot & "\" & RHS, True) Then
        RHS = ScriptRoot & "\" & RHS
    End If
    If PathExists(RHS, True) Then
        With Faces(FaceIndex)
            If (Number = 0) Or (Number > 0 And .ImageCount < Number) Then
                .ImageCount = .ImageCount + 1
                ReDim Preserve .Images(1 To .ImageCount) As Long
                Number = .ImageCount
            End If
            .Images(Number) = GetFileIndex(RHS)
            Set Files(.Images(Number)).Data = LoadTextureEx(RHS, .Dimension)
        End With
    End If
End Property

Private Sub Class_Initialize()
    FaceCount = FaceCount + 1
    ReDim Preserve Faces(1 To FaceCount) As MyFace
    FaceIndex = FaceCount
    AnimatePoint = 1
End Sub

Private Sub Class_Terminate()
    If FaceIndex <= FaceCount And FaceIndex > 0 Then
        Set Faces(FaceIndex).VBuffer = Nothing
        Erase Faces(FaceIndex).Images
        Faces(FaceIndex).ImageCount = 0
        Erase Faces(FaceIndex).Verticies
        Erase Faces(FaceIndex).Screen2D
        If FaceCount > 0 Then
            If FaceIndex < FaceCount Then
                Dim cnt As Long
                For cnt = FaceIndex + 1 To FaceCount
                    Faces(cnt - 1) = Faces(cnt)
                Next
            End If
            FaceCount = FaceCount - 1
            If FaceCount > 0 Then
                ReDim Preserve Faces(1 To FaceCount) As MyFace
                Dim brill As Electron
                For Each brill In Electrons
                    If brill.FaceIndex > FaceIndex Then brill.FaceIndex = brill.FaceIndex - 1
                Next
            End If
        End If
    End If

    Set pCoordinate = Nothing
    Set pScaled = Nothing
    Set pCenter = Nothing
    Set pPoint1 = Nothing
    Set pPoint2 = Nothing
    Set pPoint3 = Nothing
    Set pPoint4 = Nothing
End Sub

Private Sub pCoordinate_Changed(X As Single, Y As Single)
    ReScreen
End Sub

Private Sub pPoint1_Changed(X As Single, Y As Single, Z As Single, U As Single, V As Single)
    ReSquare
End Sub

Private Sub pPoint2_Changed(X As Single, Y As Single, Z As Single, U As Single, V As Single)
    ReSquare
End Sub

Private Sub pPoint3_Changed(X As Single, Y As Single, Z As Single, U As Single, V As Single)
    ReSquare
End Sub

Private Sub pPoint4_Changed(X As Single, Y As Single, Z As Single, U As Single, V As Single)
    ReSquare
End Sub
Private Sub ReScreen()
    If Not (pCoordinate Is Nothing) Then
    
        With Faces(FaceIndex)
            .Screen2D(0).X = (Coordinate.X / PointPerPixelX) * Screen.TwipsPerPixelX
            .Screen2D(0).Y = (Coordinate.Y / PointPerPixelY) * Screen.TwipsPerPixelY
            .Screen2D(1).X = .Screen2D(0).X + .Dimension.Width
            .Screen2D(1).Y = (Coordinate.Y / PointPerPixelY) * Screen.TwipsPerPixelY
            .Screen2D(2).X = (Coordinate.X / PointPerPixelX) * Screen.TwipsPerPixelX
            .Screen2D(2).Y = .Screen2D(0).Y + .Dimension.Height
            .Screen2D(3).X = .Screen2D(2).X + .Dimension.Width
            .Screen2D(3).Y = .Screen2D(1).Y + .Dimension.Height
        End With
    End If
End Sub
Private Sub ReSquare()
    If Not (pPoint1 Is Nothing Or pPoint2 Is Nothing Or pPoint3 Is Nothing Or pPoint4 Is Nothing) Then

        Dim vec As D3DVECTOR
        vec = SquareCenter(tox(pPoint1), tox(pPoint2), tox(pPoint3), tox(pPoint4))
        Center.X = vec.X
        Center.Y = vec.Y
        Center.Z = vec.Z
    
        Dim ScaleX As Single
        Dim ScaleY As Single
        
        If pScaled Is Nothing Then
            ScaleX = 1
            ScaleY = 1
        Else
            ScaleX = pScaled.X
            ScaleY = pScaled.Y
        End If
        
        Area = CreateSquare(Faces(FaceIndex).Verticies, 0, _
            MakeVector(pPoint1.X, pPoint1.Y, pPoint1.Z), _
            MakeVector(pPoint2.X, pPoint2.Y, pPoint2.Z), _
            MakeVector(pPoint3.X, pPoint3.Y, pPoint3.Z), _
            MakeVector(pPoint4.X, pPoint4.Y, pPoint4.Z), _
            ScaleX, ScaleY)
            
        Set Faces(FaceIndex).VBuffer = DDevice.CreateVertexBuffer(Len(Faces(FaceIndex).Verticies(0)) * 6, 0, FVF_RENDER, D3DPOOL_DEFAULT)
        D3DVertexBuffer8SetData Faces(FaceIndex).VBuffer, 0, Len(Faces(FaceIndex).Verticies(0)) * 6, 0, Faces(FaceIndex).Verticies(0)

    End If
End Sub


Private Sub pScaled_Changed(X As Single, Y As Single)
    ReSquare
End Sub
