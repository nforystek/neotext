VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Planet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Implements Molecule

Private pForm As Long

Private pTransparent As Boolean
Private pTranslucent As Boolean
Private pAlphablend As Boolean

Private pHeight As Single
Private pWidth As Single
Private pLength As Single
Private pSegments As Long
Private pFollow As Boolean
Private pHoning As Boolean
Private pSnapping As Boolean

Private pOuterEdge As Single
Private pInnerEdge As Single
Private pField As Single
Private pRows As Long
Private pColumns As Long

Private pFileIndex() As Long

Private pColor As Color
Private pFog As Single

Private pKey As String

Private pPlanets As Planets
Private pParent As Molecule

Private pSerialize As Boolean

Public Property Get Moved() As Boolean
    Moved = Molecule.Moved
End Property
Public Property Let Moved(ByVal RHS As Boolean)
    Molecule.Moved = RHS
End Property

Public Property Let Snapping(ByVal RHS As Boolean)
    Molecule.Snapping = RHS
End Property
Public Property Get Snapping() As Boolean
    Snapping = Molecule.Snapping
End Property

Public Property Get Planets() As Planets
    Set Planets = pPlanets
End Property
Public Property Set Planets(ByRef RHS As Planets)
    Set pPlanets = RHS
End Property

Public Property Get Serialize() As Boolean
Attribute Serialize.VB_MemberFlags = "40"
    Serialize = pSerialize
End Property
Public Property Let Serialize(ByVal RHS As Boolean)
    pSerialize = RHS
End Property

Public Property Get Fog() As Single
    Fog = pFog
End Property
Public Property Let Fog(ByVal RHS As Single)
    pFog = RHS
End Property

Public Property Get Honing() As Boolean
    Honing = pHoning
End Property
Public Property Let Honing(ByVal RHS As Boolean)
    pHoning = RHS
End Property

Public Property Get Follow() As Boolean
    Follow = pFollow
End Property
Public Property Let Follow(ByVal RHS As Boolean)
    pFollow = RHS
End Property

Friend Function RelativeColorFactor(ByVal Distance As Single) As Single
    Dim Checkto As Single
    Checkto = Round(Distance, 0)
    If PlateauInfinite Then
        RelativeColorFactor = 1
    ElseIf Checkto <= OuterEdge + Field Then
        If PlateauIsland Or (PlateauDoughnut And (Checkto >= InnerEdge)) Then
'            If Checkto <= OuterEdge And Checkto >= InnerEdge Then
'                RelativeColorFactor = 1
'            ElseIf Checkto - OuterEdge < Field And (OuterEdge - InnerEdge) - Distance > 0 Then
'                RelativeColorFactor = Round(CSng((CSng(Distance - OuterEdge)) / Field), 6)
'            Else
'                RelativeColorFactor = 1
'            End If
            If Checkto <= OuterEdge And Checkto >= InnerEdge Then
                RelativeColorFactor = 1
            ElseIf Checkto - OuterEdge < Field And (OuterEdge - InnerEdge) > 0 Then
                RelativeColorFactor = Round(CSng((CSng(Distance - OuterEdge)) / Field), 6)
            End If
            
        ElseIf PlateauHole Or (PlateauDoughnut And Checkto < OuterEdge) Then
            If Checkto >= (OuterEdge - (InnerEdge + Field)) Then
                RelativeColorFactor = Round(CSng(CSng(Field - CSng(Distance - CSng(OuterEdge - CSng(InnerEdge + Field)))) / Field), 6)
            End If
        ElseIf (Field > 0) Then
        
            If Distance = 0 Then
                RelativeColorFactor = 1
            ElseIf Checkto <= Field Then
                If Distance = 0 Then
                    RelativeColorFactor = 1
                Else
                    RelativeColorFactor = ((Field - Distance) / Field)
                End If
            Else
                RelativeColorFactor = 0
            End If
        End If
    End If
    If RelativeColorFactor > 1 Then RelativeColorFactor = 1
    If RelativeColorFactor < 0 Then RelativeColorFactor = 0
    RelativeColorFactor = Round(RelativeColorFactor, 6)
End Function

Public Property Get Clone() As Orbit
    Set Clone = Molecule.Clone_
End Property
Public Property Get Color() As Color
    If pColor Is Nothing Then Set pColor = New Color
    Set Color = pColor
End Property
Public Property Set Color(ByRef RHS As Color)
    Set pColor = RHS
End Property

Public Property Get Motions() As Motions
    Set Motions = Molecule.Motions
End Property
Public Property Set Motions(ByRef RHS As Motions)
    Set Molecule.Motions = RHS
End Property

Public Property Get Parent() As Orbit
    Set Parent = pParent
End Property
Public Property Set Parent(ByRef RHS As Orbit)
    Set pParent = RHS
End Property
'Public Property Get Parent() As Orbit
'    Set Parent = Molecule.Parent
'End Property
'Public Property Set Parent(ByRef RHS As Orbit)
'    Set Molecule.Parent = RHS
'End Property
Public Property Get Relative() As Orbit
    Set Relative = Molecule.Relative
End Property
Public Property Set Relative(ByRef RHS As Orbit)
    Set Molecule.Relative = RHS
End Property

Public Property Get Absolute() As Orbit
    Set Absolute = Molecule.Absolute
End Property
Public Property Set Absolute(ByRef RHS As Orbit)
    Set Molecule.Absolute = RHS
End Property

Public Property Get Alphablend() As Boolean
    Alphablend = pAlphablend
End Property
Public Property Let Alphablend(ByVal RHS As Boolean)
    pAlphablend = RHS
End Property
Public Property Get Transparent() As Boolean
    Transparent = pTransparent
End Property
Public Property Let Transparent(ByVal RHS As Boolean)
    pTransparent = RHS
End Property
Public Property Get Translucent() As Boolean
    Translucent = pTranslucent
End Property
Public Property Let Translucent(ByVal RHS As Boolean)
    pTranslucent = RHS
End Property


Public Property Get Volume() As Volume
    Set Volume = Molecule.Volume
End Property
Public Property Set Volume(ByRef RHS As Volume)
    Set Molecule.Volume = RHS
End Property
Public Property Get Molecules() As Molecules
   Set Molecules = Molecule.Molecules
End Property
Public Property Set Molecules(ByRef RHS As Molecules)
    Set Molecule.Molecules = RHS
End Property
Public Property Get Origin() As Point
    Set Origin = Molecule.Origin
End Property
Public Property Set Origin(ByRef RHS As Point)
    Set Molecule.Origin = RHS
End Property

Public Property Get Offset() As Point
    If Molecule.Offset Is Nothing Then Set Molecule.Offset = New Point
    Set Offset = Molecule.Offset
End Property
Public Property Set Offset(ByRef RHS As Point)
    Set Molecule.Offset = RHS
End Property

Public Property Get Rotate() As Point
    If Molecule.Rotate Is Nothing Then Set Molecule.Rotate = New Point
    Set Rotate = Molecule.Rotate
End Property
Public Property Set Rotate(ByRef RHS As Point)
    Set Molecule.Rotate = RHS
End Property


Public Property Get Scaled() As Point
    If Molecule.Scaled Is Nothing Then Set Molecule.Scaled = New Point
    Set Scaled = Molecule.Scaled
End Property
Public Property Set Scaled(ByRef RHS As Point)
    Set Molecule.Scaled = RHS
End Property

Public Property Get Ranges() As Range
    Set Ranges = Molecule.Ranges
End Property
Public Property Set Ranges(ByRef RHS As Range)
    Set Molecule.Ranges = RHS
End Property


Public Property Get Visible() As Boolean
    Visible = Molecule.Visible
End Property
Public Property Let Visible(ByVal RHS As Boolean)
    Molecule.Visible = RHS
End Property

Public Property Get Width() As Single
    If pWidth <= 0 Then
        Width = 1
    Else
        Width = pWidth
    End If
End Property
Public Property Let Width(ByVal RHS As Single)
    pWidth = RHS
    If (pFileIndex(0) > 0) Then RebuildPlanet
End Property

Public Property Get Height() As Single
    If pHeight <= 0 Then
        Height = 1
    Else
        Height = pHeight
    End If
End Property
Public Property Let Height(ByVal RHS As Single)
    pHeight = RHS
    If (pFileIndex(0) > 0) Then RebuildPlanet
End Property

Public Property Get Length() As Single
    If pLength <= 0 Then
        Length = 1
    Else
        Length = pLength
    End If
End Property
Public Property Let Length(ByVal RHS As Single)
    pLength = RHS
    If (pFileIndex(0) > 0) Then RebuildPlanet
End Property

Public Property Get OuterEdge() As Single
    OuterEdge = pOuterEdge
End Property
Public Property Let OuterEdge(ByVal RHS As Single)
    pOuterEdge = RHS
    Ranges.r = IIf(Field > -1, Field, 0) + IIf(pOuterEdge > -1, pOuterEdge, 0)
    If RHS > pInnerEdge Or RHS = -1 Then
        If (pFileIndex(0) > 0) Then RebuildPlanet
    End If
End Property

Public Property Get InnerEdge() As Single
    
    InnerEdge = pInnerEdge
End Property
Public Property Let InnerEdge(ByVal RHS As Single)
    pInnerEdge = RHS
    If RHS < pOuterEdge And RHS > -1 Then
        If (pFileIndex(0) > 0) Then RebuildPlanet
    End If
End Property

Public Property Get Segments() As Long
    If pSegments < 3 Then
        If ((Length <> 0) And (Height <> 0)) Then
            Segments = 4
        Else
            Segments = 360
        End If
    Else
        Segments = pSegments
    End If
End Property
Public Property Let Segments(ByVal RHS As Long)
    If RHS >= 3 Then
        pSegments = RHS
        If (pFileIndex(0) > 0) Then RebuildPlanet
    Else
        Err.Raise 8, "Planet", "Segments must be a value of 3 or greater."
    End If
End Property

Public Property Get Field() As Single
    If (pField = 0 And Length > 0 And Width > 0) Then
        Field = Far '(Length * Width) ^ (1 / 2)
    Else
        Field = pField
    End If
End Property
Public Property Let Field(ByVal RHS As Single)
    pField = RHS
    Ranges.r = IIf(pField > -1, pField, 0) + IIf(pOuterEdge > -1, pOuterEdge, 0)
End Property

Public Property Get Rows() As Single
    Rows = pRows
End Property
Public Property Let Rows(ByVal RHS As Single)
    pRows = RHS
    If (pFileIndex(0) > 0) Then RebuildPlanet
End Property

Public Property Get Columns() As Single
    Columns = pColumns
End Property
Public Property Let Columns(ByVal RHS As Single)
    pColumns = RHS
    If (pFileIndex(0) > 0) Then RebuildPlanet
End Property

Public Property Get Key() As String
Attribute Key.VB_UserMemId = 0
Attribute Key.VB_MemberFlags = "600"
    Key = Molecule.Key
End Property
Public Property Let Key(ByVal RHS As String)
    Molecule.Key = RHS
End Property

Public Property Get Form() As Long
    Form = pForm
End Property
Public Property Let Form(ByRef RHS As Long)
    pForm = RHS
'    Select Case pForm
'        Case PlanetTypes.Plateau
'
'        Case PlanetTypes.Screen
'
'        Case PlanetTypes.Shade
'
'        Case PlanetTypes.World
'
'    End Select
    If (pFileIndex(0) > 0) Then RebuildPlanet
End Property


Friend Property Get FileIndex(ByVal GridIndex As Long) As Long
    FileIndex = pFileIndex(GridIndex)
End Property

Public Property Get FileName(Optional ByVal Row As Long = 0, Optional ByVal Column As Long = 0) As String
    If (pFileIndex(Row * Column) > 0) Then
        FileName = GetIndexFile(pFileIndex(Row * Column))
    End If
End Property
Public Property Let FileName(Optional ByVal Row As Long = 0, Optional ByVal Column As Long = 0, ByVal RHS As String)
    If RHS = "" And pFileIndex(Row * Column) > 0 Then
        pFileIndex(Row * Column) = 0
    ElseIf PathExists(RHS, True) Then
        pFileIndex(Row * Column) = GetFileIndex(RHS)
        Set Files(pFileIndex(Row * Column)).Data = LoadTexture(RHS)
        ImageDimensions RHS, Files(pFileIndex(Row * Column)).Size
        If pWidth = 0 Then pWidth = Files(pFileIndex(Row * Column)).Size.Width
        If pHeight = 0 Then pHeight = Files(pFileIndex(Row * Column)).Size.Height
        If pLength = 0 Then pLength = ((pWidth * pHeight) ^ (1 / 2))
        RebuildPlanet
    End If
End Property

Friend Property Get PlateauInfinite()
    PlateauInfinite = ((pFileIndex(0) > 0) And ((((Form = Plateau) And ((Field = 0) Or (Field = Far))) Or PlateauHole) And (Not (PlateauIsland Or PlateauDoughnut))))
End Property
Friend Property Get PlateauIsland()
    PlateauIsland = ((pFileIndex(0) > 0) And ((OuterEdge > 0) And (InnerEdge = 0)))
End Property
Friend Property Get PlateauDoughnut()
    PlateauDoughnut = ((pFileIndex(0) > 0) And (((OuterEdge > InnerEdge) And (InnerEdge > 0))))
End Property
Friend Property Get PlateauHole()
    PlateauHole = ((pFileIndex(0) > 0) And ((InnerEdge > 0) And (OuterEdge = 0)))
End Property

Private Sub RebuildPlanet()
                    
    Dim m As Molecule
    Dim uuu As Volume
    Dim v As Matter
    Dim p1 As Point
    Dim p2 As Point
    Dim p3 As Point
    Dim P4 As Point
    
    If pColumns > 0 And pRows > 0 Then
        ReDim Preserve pFileIndex(0 To pColumns * pRows) As Long
    End If

    Dim Size As ImageDimensions
    Dim tmp As String
    If pFileIndex(0) > 0 Then
        tmp = Files(pFileIndex(0)).Path
        Size = Files(pFileIndex(0)).Size
    Else
        Size.Height = 1
        Size.Width = 1
    End If
    If tmp = "" Then Exit Sub
    
    If Volume Is Nothing Then Set Volume = New Volume
    
    If Form = PlanetTypes.World Then
            
          If (pFileIndex(0) > 0) And (Width > 0) And (Height > 0) And (Length > 0) Then '3d sky, i.e. in a box
              If Height / 3 = Width / 4 Then
                
                  Volume.Clear

                  '3x4 cube texture
                  frmMain.Picture2.Cls
                  
                  Set frmMain.Picture2.Picture = LoadPicture(tmp) '
                 ' Debug.Print frmMain.Picture2.Width / VB.Screen.TwipsPerPixelX; frmMain.Picture2.Height / VB.Screen.TwipsPerPixelY
                  
'                  Stop
                  frmMain.Picture3.Width = (Size.Width / 4) * VB.Screen.TwipsPerPixelX
                  frmMain.Picture3.Height = (Size.Height / 3) * VB.Screen.TwipsPerPixelY
                  Size.Height = (frmMain.Picture3.Width / VB.Screen.TwipsPerPixelX)
                  Size.Width = (frmMain.Picture3.Height / VB.Screen.TwipsPerPixelY)
                  
                  frmMain.Picture3.Cls
                  'Debug.Print frmMain.Picture3.Width / VB.Screen.TwipsPerPixelX; frmMain.Picture3.Height / VB.Screen.TwipsPerPixelY
                  BitBlt frmMain.Picture3.hdc, 0, 0, (frmMain.Picture3.Width \ VB.Screen.TwipsPerPixelX), (frmMain.Picture3.Height \ VB.Screen.TwipsPerPixelY), frmMain.Picture2.hdc, ((frmMain.Picture3.Width \ VB.Screen.TwipsPerPixelX) * 2), (frmMain.Picture3.Height \ VB.Screen.TwipsPerPixelY), SRCCOPY
                  tmp = GetTemporaryFolder & "\" & GetFileTitle(tmp) & "Left.bmp"
                  'tmp = "C:\Development\Neotext\Common\Projects\NTDirectX\Test\ToggleBox\Temp\" & "Left.bmp"
                  
                  SavePicture frmMain.Picture3.Image, tmp

                  Set p3 = MakePoint((Width / 2), -(Height / 2), (Length / 2)) ' right    4    __   '
                  Set p2 = MakePoint((Width / 2), -(Height / 2), -(Length / 2)) '       3    /_|/|
                  Set p1 = MakePoint((Width / 2), (Height / 2), -(Length / 2)) '          1  |/_|/
                  Set P4 = MakePoint((Width / 2), (Height / 2), (Length / 2)) '         2
                   
                  Set uuu = CreateVolumeFace(tmp, P4, p1, p2, p3, Size.Width / Width, Size.Height / Length)

                  For Each v In uuu
                      Volume.Add v
                  Next
                  uuu.Clear
                  Set uuu = Nothing

                  frmMain.Picture3.Cls
                  BitBlt frmMain.Picture3.hdc, 0, 0, (frmMain.Picture3.Width \ VB.Screen.TwipsPerPixelX), (frmMain.Picture3.Height \ VB.Screen.TwipsPerPixelY), frmMain.Picture2.hdc, ((frmMain.Picture3.Width \ VB.Screen.TwipsPerPixelX) * 3), (frmMain.Picture3.Height \ VB.Screen.TwipsPerPixelY), SRCCOPY
                  tmp = GetTemporaryFolder & "\" & GetFileTitle(tmp) & "Back.bmp"
                  'tmp = "C:\Development\Neotext\Common\Projects\NTDirectX\Test\ToggleBox\Temp\" & "Back.bmp"
                  SavePicture frmMain.Picture3.Image, tmp

                  Set p1 = MakePoint(-(Width / 2), -(Height / 2), -(Length / 2)) '      __   back  '
                  Set p2 = MakePoint((Width / 2), -(Height / 2), -(Length / 2)) '     /_|/|   4  3   '
                  Set p3 = MakePoint((Width / 2), (Height / 2), -(Length / 2)) '      |/_|/   1  2   '
                  Set P4 = MakePoint(-(Width / 2), (Height / 2), -(Length / 2)) '                    '
                  
                  Set uuu = CreateVolumeFace(tmp, p3, P4, p1, p2, Size.Width / Width, Size.Height / Height)

                  For Each v In uuu
                      Volume.Add v
                  Next
                  uuu.Clear
                  Set uuu = Nothing

                  frmMain.Picture3.Cls
                  BitBlt frmMain.Picture3.hdc, 0, 0, (frmMain.Picture3.Width \ VB.Screen.TwipsPerPixelX), (frmMain.Picture3.Height \ VB.Screen.TwipsPerPixelY), frmMain.Picture2.hdc, 0, (frmMain.Picture3.Height \ VB.Screen.TwipsPerPixelY), SRCCOPY

                  tmp = GetTemporaryFolder & "\" & GetFileTitle(tmp) & "Right.bmp"
                  'tmp = "C:\Development\Neotext\Common\Projects\NTDirectX\Test\ToggleBox\Temp\" & "Right.bmp"
                  SavePicture frmMain.Picture3.Image, tmp

                  Set p1 = MakePoint(-(Width / 2), -(Height / 2), -(Length / 2)) '      __          3  left    '
                  Set p2 = MakePoint(-(Width / 2), -(Height / 2), (Length / 2)) '     /_|/|       4            '
                  Set p3 = MakePoint(-(Width / 2), (Height / 2), (Length / 2)) '      |/_|/         2         '
                  Set P4 = MakePoint(-(Width / 2), (Height / 2), -(Length / 2)) '                 1               '
                  
                  Set uuu = CreateVolumeFace(tmp, P4, p3, p2, p1, Size.Height / Length, Size.Width / Width)

                  For Each v In uuu
                      Volume.Add v
                  Next
                  uuu.Clear
                  Set uuu = Nothing

                  frmMain.Picture3.Cls
                  BitBlt frmMain.Picture3.hdc, 0, 0, (frmMain.Picture3.Width \ VB.Screen.TwipsPerPixelX), (frmMain.Picture3.Height \ VB.Screen.TwipsPerPixelY), frmMain.Picture2.hdc, (frmMain.Picture3.Width \ VB.Screen.TwipsPerPixelX), ((frmMain.Picture3.Height \ VB.Screen.TwipsPerPixelY) * 2), SRCCOPY
                  
                  tmp = GetTemporaryFolder & "\" & GetFileTitle(tmp) & "Bottom.bmp"
                  'tmp = "C:\Development\Neotext\Common\Projects\NTDirectX\Test\ToggleBox\Temp\" & "Bottom.bmp"
                  SavePicture frmMain.Picture3.Image, tmp


                  Set p1 = MakePoint((Width / 2), -(Height / 2), -(Length / 2)) '        __    bottom       '
                  Set p2 = MakePoint(-(Width / 2), -(Height / 2), -(Length / 2)) '     /_|/|               '
                  Set p3 = MakePoint(-(Width / 2), -(Height / 2), (Length / 2)) '      |/_|/   3  4       '
                  Set P4 = MakePoint((Width / 2), -(Height / 2), (Length / 2)) '              2  1           '
                  
                  Set uuu = CreateVolumeFace(tmp, p1, p2, p3, P4, Size.Width / Width, Size.Height / Length)

                  For Each v In uuu
                      Volume.Add v
                  Next
                  uuu.Clear
                  Set uuu = Nothing

                  frmMain.Picture3.Cls
                  BitBlt frmMain.Picture3.hdc, 0, 0, (frmMain.Picture3.Width \ VB.Screen.TwipsPerPixelX), (frmMain.Picture3.Height \ VB.Screen.TwipsPerPixelY), frmMain.Picture2.hdc, (frmMain.Picture3.Width \ VB.Screen.TwipsPerPixelX), (frmMain.Picture3.Height \ VB.Screen.TwipsPerPixelY), SRCCOPY
                  
                  tmp = GetTemporaryFolder & "\" & GetFileTitle(tmp) & "Front.bmp"
                  'tmp = "C:\Development\Neotext\Common\Projects\NTDirectX\Test\ToggleBox\Temp\" & "Front.bmp"
                  SavePicture frmMain.Picture3.Image, tmp

                  Set p1 = MakePoint((Width \ 2), -(Height \ 2), (Length \ 2)) '        __      front  '
                  Set p2 = MakePoint(-(Width \ 2), -(Height \ 2), (Length \ 2)) '     /_|/|     3  4     '
                  Set p3 = MakePoint(-(Width \ 2), (Height \ 2), (Length \ 2)) '      |/_|/     2  1    '
                  Set P4 = MakePoint((Width \ 2), (Height \ 2), (Length \ 2)) '                           '
                  
                  Set uuu = CreateVolumeFace(tmp, p3, P4, p1, p2, Size.Width / Width, Size.Height / Height)

                  For Each v In uuu
                      Volume.Add v
                  Next
                  uuu.Clear
                  Set uuu = Nothing

                  frmMain.Picture3.Cls
                  BitBlt frmMain.Picture3.hdc, 0, 0, (frmMain.Picture3.Width \ VB.Screen.TwipsPerPixelX), (frmMain.Picture3.Height \ VB.Screen.TwipsPerPixelY), frmMain.Picture2.hdc, (frmMain.Picture3.Width \ VB.Screen.TwipsPerPixelX), 0, SRCCOPY
                  tmp = GetTemporaryFolder & "\" & GetFileTitle(tmp) & "Top.bmp"
                  'tmp = "C:\Development\Neotext\Common\Projects\NTDirectX\Test\ToggleBox\Temp\" & "Top.bmp"
                  SavePicture frmMain.Picture3.Image, tmp

                  Set p1 = MakePoint((Width \ 2), (Height \ 2), (Length \ 2)) '          __     top    '
                  Set p2 = MakePoint(-(Width \ 2), (Height \ 2), (Length \ 2)) '       /_|/|    2  1   '
                  Set p3 = MakePoint(-(Width \ 2), (Height \ 2), -(Length \ 2)) '      |/_|/   3  4   '
                  Set P4 = MakePoint((Width \ 2), (Height \ 2), -(Length \ 2)) '                      '
                  
                  Set uuu = CreateVolumeFace(tmp, p3, P4, p1, p2, Size.Width / Width, Size.Height / Length)

                  For Each v In uuu
                      Volume.Add v
                  Next
                  uuu.Clear
                  Set uuu = Nothing
                  
                  
              End If
              
          End If
    ElseIf Form = PlanetTypes.Plateau Then 'all the flat types
   
        If ((Rows > 0) And (Columns > 0)) And (PlateauIsland Or PlateauDoughnut) Then

            If UBound(pFileIndex) < Rows * Columns Then
                ReDim Preserve pFileIndex(0 To Rows * Columns) As Long
            End If

            Dim i As Single
            For i = 1 To Rows * Columns
                If pFileIndex(i) = 0 Then pFileIndex(i) = pFileIndex(0)
            Next

        End If
        
        'a circle at 360 segments, a square at 4 segments, a triangle at 3 segments
        If PlateauIsland() Then 'island
        
            Set m = CreateMoleculeLanding(tmp, OuterEdge, Segments, 0, Size.Width, Size.Height)
            
            Volume.Clear
            For Each v In m.Volume
                Volume.Add v
            Next
            m.Volume.Clear
            Set m = Nothing

        ElseIf PlateauDoughnut() Then 'doughnut

            Set m = CreateMoleculeLanding(tmp, OuterEdge, Segments, InnerEdge, Size.Width, Size.Height)
            
            Volume.Clear
            For Each v In m.Volume
                Volume.Add v
            Next
            m.Volume.Clear
            Set m = Nothing
                        
        ElseIf PlateauHole Then 'hole
        
            Volume.Clear

            Set m = CreateMoleculeLanding(tmp, Far, Segments, 0, Size.Width, Size.Height)

            For Each v In m.Volume
                Volume.Add v
            Next
            m.Volume.Clear
            Set m = Nothing
            
            Set m = CreateMoleculeLanding(tmp, Far, Segments, InnerEdge, Size.Width, Size.Height)

            For Each v In m.Volume
                Volume.Add v
            Next
            m.Volume.Clear
            Set m = Nothing
                        
        ElseIf PlateauInfinite Then 'stretched texture island

            Volume.Clear

            Set m = CreateMoleculeLanding(tmp, Far, Segments, 0, Size.Width * (Far / Size.Width), Size.Height * (Far / Size.Height))

            For Each v In m.Volume
                Volume.Add v
            Next
            m.Volume.Clear
            Set m = Nothing
                            
        End If

    End If
    

End Sub


Private Sub Class_Initialize()
    Stats_Planet_Count = Stats_Planet_Count + 1
    vbaObjSetAddref Molecule, ObjPtr(New Molecule)
    Key = "K" & ObjPtr(Me)
    ReDim pFileIndex(0 To 0) As Long
End Sub

Private Sub Class_Terminate()
    Stats_Planet_Count = Stats_Planet_Count - 1
    
    Set pParent = Nothing
    'Set Molecule.Parent = Nothing

    If Not Molecule.Molecules Is Nothing Then
        Molecule.Molecules.Clear
    End If
    Set Molecule.Molecules = Nothing

    If Not Molecule.Volume Is Nothing Then
        Molecule.Volume.Clear
    End If
    Set Molecule.Volume = Nothing

    vbaObjSet Molecule, ObjPtr(Nothing)
    
    Erase pFileIndex
    If Not pPlanets Is Nothing Then
        pPlanets.Clear
    End If
    Set pPlanets = Nothing

    Set pColor = Nothing
End Sub

Private Property Set Molecule_Motions(RHS As Motions)
    Set Molecule.Motions = RHS
End Property

Private Property Get Molecule_Motions() As Motions
    Set Molecule_Motions = Molecule.Motions
End Property

Private Property Let Molecule_FileName(ByVal RHS As String)
    Molecule.FileName = RHS
End Property

Private Property Get Molecule_FileName() As String
    Molecule_FileName = Molecule.FileName
End Property

Private Property Let Molecule_Key(ByVal RHS As String)
    Molecule.Key = RHS
End Property

Private Property Get Molecule_Key() As String
    Molecule_Key = Molecule.Key
End Property

Private Property Let Molecule_Moved(ByVal RHS As Boolean)
    Molecule.Moved = RHS
End Property

Private Property Get Molecule_Moved() As Boolean
    Molecule_Moved = Molecule.Moved
End Property

Private Property Set Molecule_Parent(RHS As Molecule)
    Set Molecule.Parent = RHS
End Property

Private Property Get Molecule_Parent() As Molecule
    Set Molecule_Parent = Molecule.Parent
End Property

Private Property Set Molecule_Relative(RHS As Orbit)
    Set Molecule.Relative = RHS
End Property

Private Property Get Molecule_Relative() As Orbit
    Set Molecule_Relative = Molecule.Relative
End Property

Private Property Set Molecule_Absolute(RHS As Orbit)
    Set Molecule.Absolute = RHS
End Property

Private Property Get Molecule_Absolute() As Orbit
    Set Molecule_Absolute = Molecule.Absolute
End Property

Private Property Set Molecule_Molecues(ByVal RHS As Molecules)
    Set Molecule.Molecules = RHS
End Property

Private Property Get Molecule_Molecues() As Molecules
    Set Molecule_Molecues = Molecule.Molecules
End Property

Private Property Set Molecule_Molecules(RHS As Molecules)
    Set Molecule.Molecules = RHS
End Property

Private Property Get Molecule_Molecules() As Molecules
    Set Molecule_Molecules = Molecule.Molecules
End Property

Private Property Set Molecule_Offset(ByRef RHS As Point)
    Set Molecule.Offset = RHS
End Property

Private Property Get Molecule_Offset() As Point
    Set Molecule_Offset = Molecule.Offset
End Property

Private Property Set Molecule_Origin(ByRef RHS As Point)
    Set Molecule.Origin = RHS
End Property

Private Property Get Molecule_Origin() As Point
    Set Molecule_Origin = Molecule.Origin
End Property

Private Property Set Molecule_Ranges(ByRef RHS As Range)
    Set Molecule.Ranges = RHS
End Property

Private Property Get Molecule_Ranges() As Range
    Set Molecule_Ranges = Molecule.Ranges
End Property

Private Property Set Molecule_Rotate(ByRef RHS As Point)
    Set Molecule.Rotate = RHS
End Property

Private Property Get Molecule_Rotate() As Point
    Set Molecule_Rotate = Molecule.Rotate
End Property

Private Property Set Molecule_Scaled(ByRef RHS As Point)
    Set Molecule.Scaled = RHS
End Property

Private Property Get Molecule_Scaled() As Point
    Set Molecule_Scaled = Molecule.Scaled
End Property

Private Property Let Molecule_Serialize(ByVal RHS As Boolean)
    Molecule.Serialize = RHS
End Property

Private Property Get Molecule_Serialize() As Boolean
    Molecule_Serialize = Molecule.Serialize
End Property


Private Property Let Molecule_Snapping(ByVal RHS As Boolean)
    Molecule.Snapping = RHS
End Property

Private Property Get Molecule_Snapping() As Boolean
    Molecule_Snapping = Molecule.Snapping
End Property

Private Property Let Molecule_ToString(ByVal RHS As String)
    Molecule.ToString = RHS
End Property

Private Property Get Molecule_ToString() As String
    Molecule.ToString = Molecule_ToString
End Property

Private Property Let Molecule_Visible(ByVal RHS As Boolean)
    Molecule.Visible = RHS
End Property

Private Property Get Molecule_Visible() As Boolean
    Molecule_Visible = Molecule.Visible
End Property

Private Property Set Molecule_Volume(ByRef RHS As Volume)
    Set Molecule.Volume = RHS
End Property

Private Property Get Molecule_Volume() As Volume
    Set Molecule_Volume = Molecule.Volume
End Property
