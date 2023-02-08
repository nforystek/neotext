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

Private pHeight As Single
Private pWidth As Single
Private pLength As Single
Private pSegments As Long

Private pOuterRadius As Single
Private pInnerRadius As Single
Private pRows As Long
Private pColumns As Long

Private pImageIndex() As Long

Private pFileIndex As Long

Private pColor As Color
Private pFog As Single

Private pKey As String

Public Property Get Fog() As Single
    Fog = pFog
End Property
Public Property Let Fog(ByVal RHS As Single)
    pFog = RHS
End Property

Public Property Get Color() As Color
    If pColor Is Nothing Then Set pColor = New Color
    Set Color = pColor
End Property
Public Property Set Color(ByRef RHS As Color)
    Set pColor = RHS
End Property

Public Property Get Parent() As Orbit
    Set Parent = Molecule.Parent
End Property
Public Property Set Parent(ByRef RHS As Orbit)
    Set Molecule.Parent = RHS
End Property

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

Public Property Get Serialize() As Boolean
    Serialize = Molecule.Serialize
End Property
Public Property Let Serialize(ByVal RHS As Boolean)
    Molecule.Serialize = RHS
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
    If (pFileIndex > 0) Then RebuildPlanet
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
    If (pFileIndex > 0) Then RebuildPlanet
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
    If (pFileIndex > 0) Then RebuildPlanet
End Property

Public Property Get OuterRadius() As Single
    OuterRadius = pOuterRadius
End Property
Public Property Let OuterRadius(ByVal RHS As Single)
    pOuterRadius = RHS
    Ranges.W = IIf(pInnerRadius > -1, pInnerRadius, 0) + IIf(pOuterRadius > -1, pOuterRadius, 0)
    If RHS > pInnerRadius Or RHS = -1 Then
        If (pFileIndex > 0) Then RebuildPlanet
    End If
End Property

Public Property Get InnerRadius() As Single
    InnerRadius = pInnerRadius
End Property
Public Property Let InnerRadius(ByVal RHS As Single)
    pInnerRadius = RHS
    Ranges.W = IIf(pInnerRadius > -1, pInnerRadius, 0) + IIf(pOuterRadius > -1, pOuterRadius, 0)
    If RHS < pOuterRadius And RHS > -1 Then
        If (pFileIndex > 0) Then RebuildPlanet
    End If
End Property

Public Property Get Segments() As Long
    If pSegments < 3 Then
        Segments = 360
    Else
        Segments = pSegments
    End If
End Property
Public Property Let Segments(ByVal RHS As Long)
    If RHS >= 3 Then
        pSegments = RHS
        If (pFileIndex > 0) Then RebuildPlanet
    Else
        Err.Raise 8, "Planet", "Segments must be a value of 3 or greater."
    End If
End Property

Public Property Get Rows() As Single
    Rows = pRows
End Property
Public Property Let Rows(ByVal RHS As Single)
    pRows = RHS
    If (pFileIndex > 0) Then RebuildPlanet
End Property

Public Property Get Columns() As Single
    Columns = pColumns
End Property
Public Property Let Columns(ByVal RHS As Single)
    pColumns = RHS
    If (pFileIndex > 0) Then RebuildPlanet
End Property

Public Property Get Key() As String
    Key = pKey
End Property
Public Property Let Key(ByVal RHS As String)

'    If RHS = "" And Planets.Exists(pKey) Then
'        Planets.Remove pKey
'        pKey = ""
'        Molecule.Key = ""
'    Else
        If pKey <> "" And (Not Planets.Exists(RHS)) Then
           ' If All.Exists(pKey) Then All.Remove pKey
            If Planets.Exists(pKey) Then Planets.Remove pKey
          ' All.Add Me, RHS
            Planets.Add Me, RHS
            pKey = RHS
        ElseIf pKey = "" Then
            pKey = RHS
        End If
'        Molecule.Key = pKey & "_Molecule"
'    End If
End Property

Public Property Get Form() As Long
    Form = pForm
End Property
Public Property Let Form(ByRef RHS As Long)
    pForm = RHS
    If (pFileIndex > 0) Then RebuildPlanet
End Property



Public Property Get FileName() As String
    If (pFileIndex > 0) Then
        FileName = GetIndexFile(pFileIndex)
    End If
End Property
Public Property Let FileName(ByVal RHS As String)
    If RHS = "" And pFileIndex > 0 Then
        pFileIndex = 0
    ElseIf PathExists(RHS, True) Then
        pFileIndex = GetFileIndex(RHS)
        Set Files(pFileIndex).Data = LoadTexture(RHS)
        ImageDimensions RHS, Files(pFileIndex).Size
        If pWidth = 0 Then pWidth = Files(pFileIndex).Size.Width
        If pHeight = 0 Then pHeight = Files(pFileIndex).Size.Height
        If pLength = 0 Then pLength = ((pWidth + pHeight) \ 2)
        RebuildPlanet
    End If
End Property


Private Sub RebuildPlanet()
                    
    Dim m As Molecule
    Dim uuu As Volume
    Dim V As Matter
    Dim p1 As Point
    Dim p2 As Point
    Dim p3 As Point
    Dim P4 As Point
    
    If pColumns > 0 And pRows > 0 Then
        ReDim Preserve pImageIndex(pColumns * pRows) As Long
    End If

    
    Dim tmp As String
    If pFileIndex > 0 Then
        tmp = Files(pFileIndex).path
    End If
    If tmp = "" Then Exit Sub
    
    If Form = PlanetTypes.World Then
            
          If (pFileIndex > 0) And (Width > 0) And (Height > 0) And (Length > 0) Then
              If Height / 3 = Width / 4 Then

                  Volume.Clear
                  
                  '3x4 cube texture
                  Set frmMain.Picture2.Picture = LoadPicture(tmp)
                  frmMain.Picture3.Width = frmMain.Picture2.Width / 4
                  frmMain.Picture3.Height = frmMain.Picture2.Height / 3

                  BitBlt frmMain.Picture3.hdc, 0, 0, frmMain.Picture3.Width / VB.Screen.TwipsPerPixelX, frmMain.Picture3.Height / VB.Screen.TwipsPerPixelY, frmMain.Picture2.hdc, ((frmMain.Picture3.Width / VB.Screen.TwipsPerPixelX) * 2), (frmMain.Picture3.Height / VB.Screen.TwipsPerPixelY), SRCCOPY
                  tmp = GetTemporaryFolder & "\" & GetFileTitle(tmp) & "Left.bmp"
                  SavePicture frmMain.Picture3.Image, tmp

                  Set p1 = MakePoint(-(Width / 2), -(Height / 2), (Length / 2))
                  Set p2 = MakePoint(-(Width / 2), -(Height / 2), -(Length / 2))
                  Set p3 = MakePoint(-(Width / 2), (Height / 2), -(Length / 2))
                  Set P4 = MakePoint(-(Width / 2), (Height / 2), (Length / 2))

                  Set uuu = CreateVolumeFace(tmp, p1, p2, p3, P4)

                  For Each V In uuu
                      Volume.Add V
                  Next
                  uuu.Clear
                  Set uuu = Nothing

                  BitBlt frmMain.Picture3.hdc, 0, 0, frmMain.Picture3.Width / VB.Screen.TwipsPerPixelX, frmMain.Picture3.Height / VB.Screen.TwipsPerPixelY, frmMain.Picture2.hdc, ((frmMain.Picture3.Width / VB.Screen.TwipsPerPixelX) * 3), (frmMain.Picture3.Height / VB.Screen.TwipsPerPixelY), SRCCOPY
                  tmp = GetTemporaryFolder & "\" & GetFileTitle(tmp) & "Back.bmp"
                  SavePicture frmMain.Picture3.Image, tmp

                  Set p1 = MakePoint(-(Width / 2), -(Height / 2), -(Length / 2))
                  Set p2 = MakePoint((Width / 2), -(Height / 2), -(Length / 2))
                  Set p3 = MakePoint((Width / 2), (Height / 2), -(Length / 2))
                  Set P4 = MakePoint(-(Width / 2), (Height / 2), -(Length / 2))

                  Set uuu = CreateVolumeFace(tmp, p1, p2, p3, P4)

                  For Each V In uuu
                      Volume.Add V
                  Next
                  uuu.Clear
                  Set uuu = Nothing

                  BitBlt frmMain.Picture3.hdc, 0, 0, frmMain.Picture3.Width / VB.Screen.TwipsPerPixelX, frmMain.Picture3.Height / VB.Screen.TwipsPerPixelY, frmMain.Picture2.hdc, 0, (frmMain.Picture3.Height / VB.Screen.TwipsPerPixelY), SRCCOPY

                  tmp = GetTemporaryFolder & "\" & GetFileTitle(tmp) & "Right.bmp"
                  SavePicture frmMain.Picture3.Image, tmp

                  Set p1 = MakePoint((Width / 2), -(Height / 2), -(Length / 2))
                  Set p2 = MakePoint((Width / 2), -(Height / 2), (Length / 2))
                  Set p3 = MakePoint((Width / 2), (Height / 2), (Length / 2))
                  Set P4 = MakePoint((Width / 2), (Height / 2), -(Length / 2))

                  Set uuu = CreateVolumeFace(tmp, p1, p2, p3, P4)

                  For Each V In uuu
                      Volume.Add V
                  Next
                  uuu.Clear
                  Set uuu = Nothing

                  BitBlt frmMain.Picture3.hdc, 0, 0, frmMain.Picture3.Width / VB.Screen.TwipsPerPixelX, frmMain.Picture3.Height / VB.Screen.TwipsPerPixelY, frmMain.Picture2.hdc, (frmMain.Picture3.Width / VB.Screen.TwipsPerPixelX), ((frmMain.Picture3.Height / VB.Screen.TwipsPerPixelY) * 2), SRCCOPY
                  
                  tmp = GetTemporaryFolder & "\" & GetFileTitle(tmp) & "Bottom.bmp"
                  SavePicture frmMain.Picture3.Image, tmp

                  Set p1 = MakePoint((Width / 2), -(Height / 2), -(Length / 2))
                  Set p2 = MakePoint(-(Width / 2), -(Height / 2), -(Length / 2))
                  Set p3 = MakePoint(-(Width / 2), -(Height / 2), (Length / 2))
                  Set P4 = MakePoint((Width / 2), -(Height / 2), (Length / 2))

                  Set uuu = CreateVolumeFace(tmp, p1, p2, p3, P4)

                  For Each V In uuu
                      Volume.Add V
                  Next
                  uuu.Clear
                  Set uuu = Nothing

                  BitBlt frmMain.Picture3.hdc, 0, 0, frmMain.Picture3.Width / VB.Screen.TwipsPerPixelX, frmMain.Picture3.Height / VB.Screen.TwipsPerPixelY, frmMain.Picture2.hdc, (frmMain.Picture3.Width / VB.Screen.TwipsPerPixelX), (frmMain.Picture3.Height / VB.Screen.TwipsPerPixelY), SRCCOPY
                  
                  tmp = GetTemporaryFolder & "\" & GetFileTitle(tmp) & "Front.bmp"
                  SavePicture frmMain.Picture3.Image, tmp

                  Set p1 = MakePoint((Width / 2), -(Height / 2), (Length / 2))
                  Set p2 = MakePoint(-(Width / 2), -(Height / 2), (Length / 2))
                  Set p3 = MakePoint(-(Width / 2), (Height / 2), (Length / 2))
                  Set P4 = MakePoint((Width / 2), (Height / 2), (Length / 2))

                  Set uuu = CreateVolumeFace(tmp, p1, p2, p3, P4)

                  For Each V In uuu
                      Volume.Add V
                  Next
                  uuu.Clear
                  Set uuu = Nothing

                  BitBlt frmMain.Picture3.hdc, 0, 0, frmMain.Picture3.Width / VB.Screen.TwipsPerPixelX, frmMain.Picture3.Height / VB.Screen.TwipsPerPixelY, frmMain.Picture2.hdc, (frmMain.Picture3.Width / VB.Screen.TwipsPerPixelX), 0, SRCCOPY
                  tmp = GetTemporaryFolder & "\" & GetFileTitle(tmp) & "Top.bmp"
                  SavePicture frmMain.Picture3.Image, tmp

                  Set p1 = MakePoint((Width / 2), (Height / 2), (Length / 2))
                  Set p2 = MakePoint(-(Width / 2), (Height / 2), (Length / 2))
                  Set p3 = MakePoint(-(Width / 2), (Height / 2), -(Length / 2))
                  Set P4 = MakePoint((Width / 2), (Height / 2), -(Length / 2))

                  Set uuu = CreateVolumeFace(tmp, p1, p2, p3, P4)

                  For Each V In uuu
                      Volume.Add V
                  Next
                  uuu.Clear
                  Set uuu = Nothing
                  
              End If
              
          End If
        
    ElseIf Form = PlanetTypes.Plateau And pInnerRadius > 0 Then
        Dim cutoff As Single
        

        cutoff = (pInnerRadius * 2)  '* MILE
       ' pWidth = (cutoff * ((Size.Height / 2) / 100))
      '  pHeight = (cutoff * (Size.Height / 100))
      '  cutoff = (Far / 2) '* MILE
        
        Volume.Clear

        Set uuu = CreateVolumeFace(tmp, MakePoint((cutoff / 2), 0, -(cutoff / 2)), _
                                MakePoint(-(cutoff / 2), 0, -(cutoff / 2)), _
                                MakePoint(-(cutoff / 2), 0, (cutoff / 2)), _
                                MakePoint((cutoff / 2), 0, (cutoff / 2)), _
                                 (cutoff * ((Height / 2) / 100)), (cutoff * ((Width / 2) / 100)))

        For Each V In uuu
            Volume.Add V
        Next
        uuu.Clear
        Set uuu = Nothing


      
        Set uuu = CreateVolumeLanding(tmp, (cutoff / 4) * 3, Segments, pInnerRadius / 4, (cutoff * (Height / 100)), (cutoff * (Width / 100)))
       ' Set uuu = CreateVolumeLanding(tmp, cutoff, Segments, cutoff / 4, (cutoff * (Height / 100)), (cutoff * (Width / 100)))
        
        For Each V In uuu

'            Set v.Point1 = VectorRotateAxis(v.Point1, MakePoint(0, (PI / 2), 0))
'            Set v.Point2 = VectorRotateAxis(v.Point2, MakePoint(0, (PI / 2), 0))
'            Set v.Point3 = VectorRotateAxis(v.Point3, MakePoint(0, (PI / 2), 0))

            Volume.Add V
        Next
        uuu.Clear
        Set uuu = Nothing
                
'        Set uuu = CreateVolumeFace(tmp, MakePoint((cutoff / 3), 0, -(cutoff / 3)), _
'                                MakePoint(-(cutoff / 3), 0, -(cutoff / 3)), _
'                                MakePoint(-(cutoff / 3), 0, (cutoff / 3)), _
'                                MakePoint((cutoff / 3), 0, (cutoff / 3)), _
'                                ((cutoff / 1.5) * ((Size.Width / 2) / 100)), ((cutoff / 1.5) * ((Size.Height / 2) / 100)))
'
'        For Each v In uuu
'
''            Set v.Point1 = VectorRotateAxis(v.Point1, MakePoint(0, (PI / 2), 0))
''            Set v.Point2 = VectorRotateAxis(v.Point2, MakePoint(0, (PI / 2), 0))
''            Set v.Point3 = VectorRotateAxis(v.Point3, MakePoint(0, (PI / 2), 0))
'
'            Volume.Add v
'        Next
'        uuu.Clear
'        Set uuu = Nothing

    End If

End Sub


Private Sub Class_Initialize()
    vbaObjSetAddref Molecule, ObjPtr(New Molecule)
    Molecule.Serialize = True
End Sub

Private Sub Class_Terminate()
    If pKey <> "" Then
        If Planets.Exists(pKey) Then Planets.Remove pKey
    End If
    Set pColor = Nothing
End Sub

Private Property Let Molecule_FileName(ByVal RHS As String)
    Me.FileName = RHS
End Property

Private Property Get Molecule_FileName() As String
    Molecule_FileName = Me.FileName
End Property

Private Property Let Molecule_Key(ByVal RHS As String)
    Key = RHS
End Property

Private Property Get Molecule_Key() As String
    Molecule_Key = Key
End Property

Private Property Set Molecule_Parent(RHS As Orbit)
    Set Me.Parent = RHS
End Property

Private Property Get Molecule_Parent() As Orbit
    Set Molecule_Parent = Me.Parent
End Property

Private Property Set Molecule_Relative(RHS As Orbit)
    Set Me.Relative = RHS
End Property

Private Property Get Molecule_Relative() As Orbit
    Set Molecule_Relative = Me.Relative
End Property

Private Property Set Molecule_Absolute(RHS As Orbit)
    Set Me.Absolute = RHS
End Property

Private Property Get Molecule_Absolute() As Orbit
    Set Molecule_Absolute = Me.Absolute
End Property

Private Property Set Molecule_Molecues(ByVal RHS As Molecules)
    Set Me.Molecules = RHS
End Property

Private Property Get Molecule_Molecues() As Molecules
    Set Molecule_Molecues = Me.Molecules
End Property

Private Property Set Molecule_Molecules(RHS As Molecules)
    Set Me.Molecules = RHS
End Property

Private Property Get Molecule_Molecules() As Molecules
    Set Molecule_Molecules = Me.Molecules
End Property

Private Property Set Molecule_Offset(ByRef RHS As Point)
    Set Me.Offset = RHS
End Property

Private Property Get Molecule_Offset() As Point
    Set Molecule_Offset = Me.Offset
End Property

Private Property Set Molecule_Origin(ByRef RHS As Point)
    Set Me.Origin = RHS
End Property

Private Property Get Molecule_Origin() As Point
    Set Molecule_Origin = Me.Origin
End Property

Private Property Set Molecule_Ranges(ByRef RHS As Range)
    Set Me.Ranges = RHS
End Property

Private Property Get Molecule_Ranges() As Range
    Set Molecule_Ranges = Me.Ranges
End Property

Private Property Set Molecule_Rotate(ByRef RHS As Point)
    Set Me.Rotate = RHS
End Property

Private Property Get Molecule_Rotate() As Point
    Set Molecule_Rotate = Me.Rotate
End Property

Private Property Set Molecule_Scaled(ByRef RHS As Point)
    Set Me.Scaled = RHS
End Property

Private Property Get Molecule_Scaled() As Point
    Set Molecule_Scaled = Me.Scaled
End Property

Private Property Let Molecule_Serialize(ByVal RHS As Boolean)
    Me.Serialize = RHS
End Property

Private Property Get Molecule_Serialize() As Boolean
    Molecule_Serialize = Me.Serialize
End Property


Private Property Let Molecule_Visible(ByVal RHS As Boolean)
    Me.Visible = RHS
End Property

Private Property Get Molecule_Visible() As Boolean
    Molecule_Visible = Me.Visible
End Property

Private Property Set Molecule_Volume(ByRef RHS As Volume)
    Set Me.Volume = RHS
End Property

Private Property Get Molecule_Volume() As Volume
    Set Molecule_Volume = Me.Volume
End Property
