VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Molecule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements Orbit

Private pVolume As Volume
Private pMolecules As Molecules
Private pParent As Molecule
Private pFileName As String
Private pKey As String

Private pSerialize As Boolean


Public Property Get Serialize() As Boolean
Attribute Serialize.VB_MemberFlags = "40"
    Serialize = pSerialize
End Property
Public Property Let Serialize(ByVal RHS As Boolean)
    pSerialize = RHS
End Property

Public Property Get Parent() As Molecule
    Set Parent = pParent
End Property
Public Property Set Parent(ByRef RHS As Molecule)
    Set pParent = RHS
End Property

Public Property Get Motions() As Motions
    Set Motions = Orbit.Motions
End Property
Public Property Set Motions(ByRef RHS As Motions)
    Set Orbit.Motions = RHS
End Property

'Public Property Get Parent() As Orbit
'    Set Parent = Orbit.Parent
'End Property
'Public Property Set Parent(ByRef RHS As Orbit)
'    Set Orbit.Parent = RHS
'End Property

Public Property Get Relative() As Orbit
    Set Relative = Orbit.Relative
End Property
Public Property Set Relative(ByRef RHS As Orbit)
    Set Orbit.Relative = RHS
End Property

Public Property Get Absolute() As Orbit
    Set Absolute = Orbit.Absolute
End Property
Public Property Set Absolute(ByRef RHS As Orbit)
    Set Orbit.Absolute = RHS
End Property

Public Property Let ToString(ByVal RHS As String)
    Orbit.ToString = RHS
End Property
Public Property Get ToString() As String
Attribute ToString.VB_MemberFlags = "40"
    ToString = Orbit.ToString
End Property

Public Property Get Key() As String
Attribute Key.VB_UserMemId = 0
Attribute Key.VB_MemberFlags = "600"
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

Public Property Get FileName() As String
    If (pFileName <> "") Then
        FileName = pFileName
    End If
End Property
Public Property Let FileName(ByVal RHS As String)
    If Volume Is Nothing Then Set pVolume = New Volume
    
    If RHS = "" And pFileName <> "" Then
        Volume.Clear
        pFileName = ""
    ElseIf PathExists(RHS, True) And LCase(RHS) <> LCase(pFileName) Then
        Volume.Clear
        pFileName = RHS
        Dim m As Molecule
        Set m = CreateMoleculeMesh(pFileName)
        Dim v As Matter
        For Each v In m.Volume
            Volume.Add v
        Next
        m.Volume.Clear
        Set m = Nothing
    End If
End Property

Public Property Get Volume() As Volume
    Set Volume = pVolume
End Property
Public Property Set Volume(ByRef RHS As Volume)
    Set pVolume = RHS
End Property
Public Property Get Molecules() As Molecules
    Set Molecules = pMolecules
End Property
Public Property Set Molecules(ByRef RHS As Molecules)
    Set pMolecules = RHS
End Property

Public Property Get Origin() As Point
    Set Origin = Orbit.Origin
End Property
Public Property Set Origin(ByRef RHS As Point)
    Set Orbit.Origin = RHS
End Property

Public Property Get Offset() As Point
    Set Offset = Orbit.Offset
End Property
Public Property Set Offset(ByRef RHS As Point)
    Set Orbit.Offset = RHS
End Property

Public Property Get Rotate() As Point
    Set Rotate = Orbit.Rotate
End Property
Public Property Set Rotate(ByRef RHS As Point)
    Set Orbit.Rotate = RHS
End Property

Public Property Get Scaled() As Point
    Set Scaled = Orbit.Scaled
End Property
Public Property Set Scaled(ByRef RHS As Point)
    Set Orbit.Scaled = RHS
End Property

Public Property Get Ranges() As Range
    Set Ranges = Orbit.Ranges
End Property
Public Property Set Ranges(ByRef RHS As Range)
    Set Orbit.Ranges = RHS
End Property

Public Property Get Visible() As Boolean
    Visible = Orbit.Visible
End Property
Public Property Let Visible(ByVal RHS As Boolean)
    Orbit.Visible = RHS
End Property

Private Sub Class_Initialize()
    Stats_Molecule_Count = Stats_Molecule_Count + 1
    vbaObjSetAddref Orbit, ObjPtr(New Orbit)
    'Set Orbit = New Orbit
    
    pKey = "K" & ObjPtr(Me)

End Sub


Friend Sub Class_Terminate()
    'Set Orbit = Nothing
    Stats_Molecule_Count = Stats_Molecule_Count - 1
        
    Set pParent = Nothing
    
    If Not pMolecules Is Nothing Then
        pMolecules.Clear
    End If
    Set pMolecules = Nothing
    If Not pVolume Is Nothing Then
        pVolume.Clear
    End If
    Set pVolume = Nothing
    
    Set Orbit.Offset = Nothing
    Set Orbit.Origin = Nothing
    Set Orbit.Ranges = Nothing
    Set Orbit.Rotate = Nothing
    Set Orbit.Scaled = Nothing

    Set Orbit.Parent = Nothing
    Set Orbit.Absolute = Nothing
    Set Orbit.Relative = Nothing

    'Set Orbit = Nothing
    'vbaObjSet Orbit, ObjPtr(Nothing)
End Sub

Private Property Set Orbit_Motions(RHS As Motions)
    Set Orbit.Motions = RHS
End Property

Private Property Get Orbit_Motions() As Motions
    Set Orbit_Motions = Orbit.Motions
End Property

Private Property Set Orbit_Relative(RHS As Orbit)
    Set Orbit.Relative = RHS
End Property

Private Property Get Orbit_Relative() As Orbit
    Set Orbit_Relative = Orbit.Relative
End Property


Private Property Set Orbit_Absolute(RHS As Orbit)
    Set Orbit.Absolute = RHS
End Property

Private Property Get Orbit_Absolute() As Orbit
    Set Orbit_Absolute = Orbit.Absolute
End Property

Private Property Get Orbit_Offset() As Point
    Set Orbit_Offset = Orbit.Offset
End Property

Private Property Set Orbit_Offset(ByRef RHS As Point)
    Set Orbit.Offset = RHS
End Property

Private Property Get Orbit_Origin() As Point
    Set Orbit_Origin = Orbit.Origin
End Property

Private Property Set Orbit_Origin(ByRef RHS As Point)
    Set Orbit.Origin = RHS
End Property

Private Property Get Orbit_Ranges() As Range
    Set Orbit_Ranges = Orbit.Ranges
End Property

Private Property Set Orbit_Ranges(ByRef RHS As Range)
    Set Orbit.Ranges = RHS
End Property

Private Property Get Orbit_Rotate() As Point
    Set Orbit_Rotate = Orbit.Rotate
End Property

Private Property Set Orbit_Rotate(ByRef RHS As Point)
    Set Orbit.Rotate = RHS
End Property

Private Property Get Orbit_Scaled() As Point
    Set Orbit_Scaled = Orbit.Scaled
End Property

Private Property Set Orbit_Scaled(ByRef RHS As Point)
    Set Orbit.Scaled = RHS
End Property

Private Property Get Orbit_Serialize() As Boolean
    Orbit_Serialize = Orbit.Serialize
End Property

Private Property Let Orbit_Serialize(ByVal RHS As Boolean)
    Orbit.Serialize = RHS
End Property

Private Property Get Orbit_ToString() As String
    Orbit_ToString = Orbit.ToString
End Property

Private Property Let Orbit_ToString(ByVal RHS As String)
    Orbit.ToString = RHS
End Property

Private Property Get Orbit_Visible() As Boolean
    Orbit_Visible = Orbit.Visible
End Property

Private Property Let Orbit_Visible(ByVal RHS As Boolean)
    Orbit.Visible = RHS
End Property
