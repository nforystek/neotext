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
Private pSerialize As Boolean
Private pFileName As String
Private pKey As String

Public Property Get Parent() As Orbit
    Set Parent = Orbit.Parent
End Property
Public Property Set Parent(ByRef RHS As Orbit)
    Set Orbit.Parent = RHS
End Property

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


Public Property Get Key() As String
    Key = pKey
End Property
Public Property Let Key(ByVal RHS As String)
'    If Molecules.Exists(pKey) And pKey <> "" Then Molecules.Remove pKey
'    pKey = RHS
'    If RHS <> "" And (Not Motions.Exists(RHS)) Then Molecules.Add Me, RHS
    
'    If RHS = "" And Molecules.Exists(pKey) Then
'        Molecules.Remove pKey
'        pKey = ""
'    Else
        If pKey <> "" And (Not Molecules.Exists(RHS)) Then
            If Molecules.Exists(pKey) Then Molecules.Remove pKey
            Molecules.Add Me, RHS
            pKey = RHS
        ElseIf pKey = "" Then
            pKey = RHS
        End If
'    End If
End Property

Public Property Get FileName() As String
    If (pFileName <> "") Then
        FileName = pFileName
    End If
End Property
Public Property Let FileName(ByVal RHS As String)
    If RHS = "" And pFileName <> "" Then
        Volume.Clear
        pFileName = ""
    ElseIf PathExists(RHS, True) Then
        Volume.Clear
        pFileName = RHS
        Dim m As Molecule
        Set m = CreateMoleculeMesh(pFileName)
        Dim V As Matter
        For Each V In m.Volume
            Volume.Add V
        Next
        m.Volume.Clear
        Set m = Nothing
    End If
End Property

Public Property Get Volume() As Volume
    If pVolume Is Nothing Then Set pVolume = New Volume
    Set Volume = pVolume
End Property
Public Property Set Volume(ByRef RHS As Volume)
    Set pVolume = RHS
End Property
Public Property Get Molecules() As Molecules
    If pMolecules Is Nothing Then
        Set pMolecules = New Molecules
        Set pMolecules.Parent = Me
    End If
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

Public Property Get Serialize() As Boolean
    Serialize = pSerialize
    'Serialize = Orbit.Serialize
End Property
Public Property Let Serialize(ByVal RHS As Boolean)
    pSerialize = RHS
    'Orbit.Serialize = RHS
End Property

Public Property Get Visible() As Boolean
    Visible = Orbit.Visible
End Property
Public Property Let Visible(ByVal RHS As Boolean)
    Orbit.Visible = RHS
End Property

Private Sub Class_Initialize()
    vbaObjSetAddref Orbit, ObjPtr(New Orbit)

End Sub

'Public Property Get NotEmpty() As Boolean
'    NotEmpty = pOrbit.NotEmpty
'End Property
'
'Public Sub StackAdd(RHS As Orbit)
'    pOrbit.StackAdd RHS
'End Sub
'
'Public Function StackDel() As Orbit
'    Set StackDel = pOrbit.StackDel
'End Function
'
'Private Sub Class_Initialize()
'    'Set Orbit = pOrbit
'End Sub

Private Sub Class_Terminate()
    Set pVolume = Nothing
    Set pMolecules = Nothing
    If pKey <> "" Then
        If Molecules.Exists(pKey) Then Molecules.Remove pKey
       ' If All.Exists(pKey) Then All.Remove pKey
    End If
End Sub

Private Property Set Orbit_Relative(RHS As Orbit)
    Set Me.Relative = RHS
End Property

Private Property Get Orbit_Relative() As Orbit
    Set Orbit_Relative = Me.Relative
End Property


Private Property Set Orbit_Absolute(RHS As Orbit)
    Set Me.Absolute = RHS
End Property

Private Property Get Orbit_Absolute() As Orbit
    Set Orbit_Absolute = Me.Absolute
End Property

Private Property Get Orbit_Offset() As Point
    Set Orbit_Offset = Me.Offset
End Property

Private Property Set Orbit_Offset(ByRef RHS As Point)
    Set Me.Offset = RHS
End Property

Private Property Get Orbit_Origin() As Point
    Set Orbit_Origin = Me.Origin
End Property

Private Property Set Orbit_Origin(ByRef RHS As Point)
    Set Me.Origin = RHS
End Property

Private Property Get Orbit_Ranges() As Range
    Set Orbit_Ranges = Me.Ranges
End Property

Private Property Set Orbit_Ranges(ByRef RHS As Range)
    Set Me.Ranges = RHS
End Property

Private Property Get Orbit_Rotate() As Point
    Set Orbit_Rotate = Me.Rotate
End Property

Private Property Set Orbit_Rotate(ByRef RHS As Point)
    Set Me.Rotate = RHS
End Property

Private Property Get Orbit_Scaled() As Point
    Set Orbit_Scaled = Me.Scaled
End Property

Private Property Set Orbit_Scaled(ByRef RHS As Point)
    Set Me.Scaled = RHS
End Property

Private Property Get Orbit_Serialize() As String
    Orbit_Serialize = Me.Serialize
End Property

Private Property Let Orbit_Serialize(ByVal RHS As String)
    Me.Serialize = RHS
End Property


Private Property Get Orbit_Visible() As Boolean
    Orbit_Visible = Me.Visible
End Property

Private Property Let Orbit_Visible(ByVal RHS As Boolean)
    Me.Visible = RHS
End Property
