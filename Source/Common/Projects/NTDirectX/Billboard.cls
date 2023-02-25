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

Implements Molecule

Private pForm As Long

Private pTransparent As Boolean
Private pTranslucent As Boolean

Private pHeight As Single
Private pWidth As Single
Private pLength As Single

Private pOuterRadii As Single
Private pInnerRadii As Single
Private pSegments As Long

Private pRows As Long
Private pColumns As Long
Private pRadii As Single

Private pImageIndex() As Long

Private pFileIndex As Long

Private pSerialize As Boolean
Public Property Get Serialize() As Boolean
Attribute Serialize.VB_MemberFlags = "40"
    Serialize = pSerialize
End Property
Public Property Let Serialize(ByVal RHS As Boolean)
    pSerialize = RHS
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

Public Property Get Key() As String
    Key = Molecule.Key
End Property
Public Property Let Key(ByVal RHS As String)
    Molecule.Key = RHS
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
        RebuildBillboard
    End If
End Property

Public Property Get Form() As Long
    Form = pForm
End Property
Public Property Let Form(ByRef RHS As Long)
    pForm = RHS
    RebuildBillboard
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
    Set Offset = Molecule.Offset
End Property
Public Property Set Offset(ByRef RHS As Point)
    Set Molecule.Offset = RHS
End Property

Public Property Get Rotate() As Point
    Set Rotate = Molecule.Rotate
End Property
Public Property Set Rotate(ByRef RHS As Point)
    Set Molecule.Rotate = RHS
End Property


Public Property Get Scaled() As Point
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

Public Property Get Radii() As Single
    Radii = pRadii
End Property
Public Property Let Radii(ByVal RHS As Single)
    pRadii = RHS
    RebuildBillboard
End Property

Public Property Get OuterRadii() As Single
    OuterRadii = pOuterRadii
End Property
Public Property Let OuterRadii(ByVal RHS As Single)
    pOuterRadii = RHS
    RebuildBillboard
End Property
Public Property Get InnerRadii() As Single
    InnerRadii = pInnerRadii
End Property
Public Property Let InnerRadii(ByVal RHS As Single)
    pInnerRadii = RHS
    RebuildBillboard
End Property
Public Property Get Segments() As Long
    Segments = pSegments
End Property
Public Property Let Segments(ByVal RHS As Long)
    pSegments = RHS
    RebuildBillboard
End Property

Public Property Get Width() As Single
    If (pWidth = 0) And (Not ((Form And Inland) = Inland)) And (Not ((Form And Island) = Island)) Then
        Width = Far
    Else
        Width = pWidth
    End If
End Property
Public Property Let Width(ByVal RHS As Single)
    pWidth = RHS
    RebuildBillboard
End Property
Public Property Get Height() As Single
    If (pHeight = 0) And (Not ((Form And Inland) = Inland)) And (Not ((Form And Island) = Island)) Then
        Height = Far
    Else
        Height = pHeight
    End If
End Property
Public Property Let Height(ByVal RHS As Single)
    pHeight = RHS
    RebuildBillboard
End Property
Public Property Get Length() As Single
    If (pLength = 0) And (Not ((Form And Inland) = Inland)) And (Not ((Form And Island) = Island)) Then
        Length = Far
    Else
        Length = pLength
    End If
End Property
Public Property Let Length(ByVal RHS As Single)
    pLength = RHS
    RebuildBillboard
End Property

Public Property Get Rows() As Single
    Rows = pRows
End Property
Public Property Let Rows(ByVal RHS As Single)
    pRows = RHS
    RebuildBillboard
End Property
Public Property Get Columns() As Single
    Columns = pColumns
End Property
Public Property Let Columns(ByVal RHS As Single)
    pColumns = RHS
    RebuildBillboard
End Property

Public Property Get Field() As Single
    If (Molecule.Ranges.W = 0 And Length > 0 And Width > 0) And (Not ((Form And Inland) = Inland)) And (Not ((Form And Island) = Island)) Then
        Field = Far '(Length * Width) ^ (1 / 2)
    Else
        Field = Molecule.Ranges.W
    End If
End Property
Public Property Let Field(ByVal RHS As Single)
    Molecule.Ranges.W = RHS
End Property

Private Sub RebuildBillboard()
                    
    Dim m As Molecule
    Dim V As Matter
    Dim p1 As Point
    Dim p2 As Point
    Dim p3 As Point
    Dim P4 As Point
    
    Dim Size As ImageDimensions
    
    Dim tmp As String
    If pFileIndex > 0 Then
        tmp = Files(pFileIndex).path
        Size = Files(pFileIndex).Size
    Else
        Size.Height = 1
        Size.Width = 1
    End If
    If tmp = "" Then Exit Sub
    
    Dim cutoff As Single
    cutoff = Far
    
    If Form = PlanetTypes.Island Then

        If (pFileIndex > 0) And (((OuterRadii > 0) And (InnerRadii = 0)) Or (Radii > 0)) And (Segments > 0) Then

            If Radii > 0 Then
                Set m = CreateMoleculeLanding(tmp, Radii, Segments, 0, Size.Width, Size.Height)
            ElseIf OuterRadii > 0 Then
                Set m = CreateMoleculeLanding(tmp, OuterRadii, Segments, 0, Size.Width, Size.Height)
            End If
            
            Volume.Clear
            For Each V In m.Volume
                Volume.Add V
            Next
            m.Volume.Clear
            Set m = Nothing
                        
        End If
    ElseIf Form = PlanetTypes.Inland Then

        If (pFileIndex > 0) And (((OuterRadii > InnerRadii) And (InnerRadii > 0)) Or (Radii > 0)) And (Segments > 0) Then

            If InnerRadii > 0 And InnerRadii < Radii Then
                Set m = CreateMoleculeLanding(tmp, Radii, Segments, InnerRadii, Size.Width, Size.Height)
            ElseIf (Radii > 0) And Not (OuterRadii > 0) Then
                Set m = CreateMoleculeLanding(tmp, (cutoff / 2), Segments, Radii, Size.Width, Size.Height)
            ElseIf OuterRadii > InnerRadii And InnerRadii > 0 Then
                Set m = CreateMoleculeLanding(tmp, OuterRadii, Segments, InnerRadii, Size.Width, Size.Height)
            End If
            
            Volume.Clear
            For Each V In m.Volume
                Volume.Add V
            Next
            m.Volume.Clear
            Set m = Nothing

        End If
    End If


        
        
'            If (Segments > 0) Then
'
'                Set m = CreateMoleculeFace(Tmp, MakePoint((cutoff / 2), 0, -(cutoff / 2)), _
'                                        MakePoint(-(cutoff / 2), 0, -(cutoff / 2)), _
'                                        MakePoint(-(cutoff / 2), 0, (cutoff / 2)), _
'                                        MakePoint((cutoff / 2), 0, (cutoff / 2)), _
'                                        (cutoff * ((Size.Width / 2) / 100)), (cutoff * ((Size.Height / 2) / 100)))
'                Volume.Clear
'                For Each v In m.Volume
'                    Volume.Add v
'                Next
'                m.Volume.Clear
'                Set m = Nothing
'
'
'                If ((Form And Inland) = Inland) Then
'                    If InnerRadii > 0 And InnerRadii < Radii Then
'                        Set m = CreateMoleculeLanding(Tmp, Radii, Segments, InnerRadii, Size.Width, Size.Height)
'                    ElseIf (Radii > 0) And Not (OuterRadii > 0) Then
'                        Set m = CreateMoleculeLanding(Tmp, (cutoff / 2), Segments, Radii, Size.Width, Size.Height)
'                    ElseIf OuterRadii > InnerRadii And InnerRadii > 0 Then
'                        Set m = CreateMoleculeLanding(Tmp, OuterRadii, Segments, InnerRadii, Size.Width, Size.Height)
'                    End If
'
'                    If Not m Is Nothing Then
'                        For Each v In m.Volume
'                            Volume.Add v
'                        Next
'                        m.Volume.Clear
'                        Set m = Nothing
'                    End If
'                ElseIf ((Form And Island) = Island) Then
'
'                    If Radii > 0 Then
'                        Set m = CreateMoleculeLanding(Tmp, Radii, Segments, 0, Size.Width, Size.Height)
'                    ElseIf OuterRadii > 0 Then
'                        Set m = CreateMoleculeLanding(Tmp, OuterRadii, Segments, 0, Size.Width, Size.Height)
'                    End If
'
'                    If Not m Is Nothing Then
'                        For Each v In m.Volume
'                            Volume.Add v
'                        Next
'                        m.Volume.Clear
'                        Set m = Nothing
'                    End If
'
'                End If
'            Else
'
'                If ((Length <> 0) And (Width <> 0)) And ((Form And BillboardTypes.Island) = BillboardTypes.Island) Then
'                    If ((Rows > 0) And (Columns > 0)) Then
'
'                        Volume.Clear
'
'                        Dim i As Single
'                        Dim X As Single
'                        Dim Z As Single
'
'                        If UBound(pImageIndex) < Rows * Columns Then
'                            ReDim Preserve pImageIndex(1 To Rows * Columns) As Long
'                        End If
'
'                        For i = 2 To Rows * Columns
'                            If pImageIndex(i) = 0 Then pImageIndex(i) = pImageIndex(1)
'                        Next
'
'                        X = -((Rows \ 2) * Width) - Width
'                        Z = ((Columns \ 2) * Length)
'
'                        For i = 0 To (((Rows * Columns) * 6) - 1) Step 6
'
'                            If (i \ 6) Mod Columns = 0 Then
'                                Z = Z - Length
'                                X = -((Rows \ 2) * Width)
'                            Else
'                                X = X + Width
'                            End If
'                            If pImageIndex((i \ 6) + 1) > 0 Then
'                                Tmp = Files(pImageIndex((i \ 6) + 1)).path
'                            Else
'                                Tmp = ""
'                            End If
'
'                            Set m = CreateMoleculeFace(Tmp, MakePoint(((Width / 2) + X), 0, (-(Length / 2) + Z)), _
'                                                        MakePoint((-(Width / 2) + X), 0, (-(Length / 2) + Z)), _
'                                                        MakePoint((-(Width / 2) + X), 0, ((Length / 2) + Z)), _
'                                                        MakePoint(((Width / 2) + X), 0, ((Length / 2) + Z)), _
'                                                        (Width * ((Files(pImageIndex((i \ 6) + 1)).Size.Width / 2) / 100)), (Length * ((Files(pImageIndex((i \ 6) + 1)).Size.Height / 2) / 100)))
'
'                            For Each v In m.Volume
'                                Volume.Add v
'                            Next
'                            m.Volume.Clear
'                            Set m = Nothing
'
'                        Next
'                    Else
'                        Volume.Clear
'
'                        Set m = CreateMoleculeFace(Tmp, _
'                                                    MakePoint((Width / 2), 0, -(Length / 2)), _
'                                                    MakePoint(-(Width / 2), 0, -(Length / 2)), _
'                                                    MakePoint(-(Width / 2), 0, (Length / 2)), _
'                                                    MakePoint((Width / 2), 0, (Length / 2)), _
'                                                    (Width * ((Size.Width / 2) / 100)), (Length * ((Size.Height / 2) / 100)))
'                        For Each v In m.Volume
'                            Volume.Add v
'                        Next
'                        m.Volume.Clear
'                        Set m = Nothing
'
'                    End If
'
'                Else
'                    Volume.Clear
'
'                    Set m = CreateMoleculeFace(Tmp, _
'                                                MakePoint((cutoff / 2), 0, -(cutoff / 2)), _
'                                                MakePoint(-(cutoff / 2), 0, -(cutoff / 2)), _
'                                                MakePoint(-(cutoff / 2), 0, (cutoff / 2)), _
'                                                MakePoint((cutoff / 2), 0, (cutoff / 2)), _
'                                                (cutoff * ((Size.Width / 2) / 100)), (cutoff * ((Size.Height / 2) / 100)))
'
'                    For Each v In m.Volume
'                        Volume.Add v
'                    Next
'                    m.Volume.Clear
'                    Set m = Nothing
'
'                    If ((Length <> 0) And (Width <> 0)) And ((Form And Inland) = Inland) Then
'
'
'                        Set m = CreateMoleculeFace(Tmp, _
'                                                    MakePoint(-(Width / 2), 0, -(Length / 2)), _
'                                                    MakePoint(-((cutoff / 2) - (Width / 2)), 0, -(Length / 2)), _
'                                                    MakePoint(-((cutoff / 2) - (Width / 2)), 0, ((cutoff / 2) - (Length / 2))), _
'                                                    MakePoint(-(Width / 2), 0, ((cutoff / 2) - (Length / 2))), _
'                                                    (cutoff * ((Size.Width / 2) / 100)), (cutoff * ((Size.Height / 2) / 100)))
'                        For Each v In m.Volume
'                            Volume.Add v
'                        Next
'                        m.Volume.Clear
'                        Set m = Nothing
'
'                        Set m = CreateMoleculeFace(Tmp, _
'                                                    MakePoint(-(Width / 2), 0, (Length / 2)), _
'                                                    MakePoint(-(Width / 2), 0, ((cutoff / 2) - (Length / 2))), _
'                                                    MakePoint(((cutoff / 2) - (Width / 2)), 0, ((cutoff / 2) - (Length / 2))), _
'                                                    MakePoint(((cutoff / 2) - (Width / 2)), 0, -(Length / 2)), _
'                                                    (cutoff * ((Size.Width / 2) / 100)), (cutoff * ((Size.Height / 2) / 100)))
'
'                        For Each v In m.Volume
'                            Volume.Add v
'                        Next
'                        m.Volume.Clear
'                        Set m = Nothing
'
'                        Set m = CreateMoleculeFace(Tmp, _
'                                                    MakePoint((Width / 2), 0, (Length / 2)), _
'                                                    MakePoint(((cutoff / 2) - (Width / 2)), 0, (Length / 2)), _
'                                                    MakePoint(((cutoff / 2) - (Width / 2)), 0, -((cutoff / 2) - (Length / 2))), _
'                                                    MakePoint((Width / 2), 0, -((cutoff / 2) - (Length / 2))), _
'                                                    (cutoff * ((Size.Width / 2) / 100)), (cutoff * ((Size.Height / 2) / 100)))
'
'                        For Each v In m.Volume
'                            Volume.Add v
'                        Next
'                        m.Volume.Clear
'                        Set m = Nothing
'
'                        Set m = CreateMoleculeFace(Tmp, _
'                                                    MakePoint((Width / 2), 0, -(Length / 2)), _
'                                                    MakePoint((Width / 2), 0, -((cutoff / 2) - (Length / 2))), _
'                                                    MakePoint(-((cutoff / 2) - (Width / 2)), 0, -((cutoff / 2) - (Length / 2))), _
'                                                    MakePoint(-((cutoff / 2) - (Width / 2)), 0, -(Length / 2)), _
'                                                    (cutoff * ((Size.Width / 2) / 100)), (cutoff * ((Size.Height / 2) / 100)))
'                        For Each v In m.Volume
'                            Volume.Add v
'                        Next
'                        m.Volume.Clear
'                        Set m = Nothing
'
'                    End If
'
'                End If
'            End If
'
'        Else


End Sub


Private Sub Class_Initialize()
    vbaObjSetAddref Molecule, ObjPtr(New Molecule)
    Key = "K" & ObjPtr(Me)

End Sub

Private Sub Class_Terminate()
    If Key <> "" Then
        If Billboards.Exists(Key) Then Billboards.Remove Key
    End If

End Sub

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

Private Property Set Molecule_Parent(RHS As Orbit)
    Set Molecule.Parent = RHS
End Property

Private Property Get Molecule_Parent() As Orbit
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
    Set Molecule_Volume = Me.Volume
End Property


