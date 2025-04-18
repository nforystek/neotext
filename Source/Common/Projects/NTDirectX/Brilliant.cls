VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Brilliant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Implements Orbit

Public SunLight As Boolean

Public LightBlink As Single
Public LightTimer As Single
Public LightIsOn As Boolean
    
Public DiffuseRoll As Single
Public DiffuseTimer As Single
Public DIffuseMax As Single
Public DiffuseNow As Single
    
Private pLightIndex As Long

Public Diffuse As Color
Attribute Diffuse.VB_VarHelpID = -1
Public Ambience As Color
Attribute Ambience.VB_VarHelpID = -1
Public Specular As Color
Attribute Specular.VB_VarHelpID = -1
Public Attenuation As Point

Private pKey As String

Private pSerialize As Boolean


Public Property Get Serialize() As Boolean
Attribute Serialize.VB_MemberFlags = "40"
    Serialize = pSerialize
End Property
Public Property Let Serialize(ByVal RHS As Boolean)
    pSerialize = RHS
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

Public Property Get Motions() As Motions
    Set Motions = Orbit.Motions
End Property
Public Property Set Motions(ByRef RHS As Motions)
    Set Orbit.Motions = RHS
End Property

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

Public Property Get Moved() As Boolean
    Moved = Orbit.Moved
End Property
Public Property Let Moved(ByVal RHS As Boolean)
    Orbit.Moved = RHS
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

Public Property Get LightIndex() As Long
    LightIndex = pLightIndex
End Property
Friend Property Let LightIndex(ByVal RHS As Long)
    pLightIndex = RHS
End Property
Public Property Get Phi() As Single
    Phi = Lights(LightIndex).Phi
End Property
Public Property Let Phi(ByVal RHS As Single)
    Lights(LightIndex).Phi = RHS
    DDevice.SetLight LightIndex - 1, Lights(LightIndex)
End Property

Public Property Get Theta() As Single
    Theta = Lights(LightIndex).Theta
End Property
Public Property Let Theta(ByVal RHS As Single)
    Lights(LightIndex).Theta = RHS
    DDevice.SetLight LightIndex - 1, Lights(LightIndex)
End Property

Public Property Get FallOff() As Single
    FallOff = Lights(LightIndex).FallOff
End Property
Public Property Let FallOff(ByVal RHS As Single)
    Lights(LightIndex).FallOff = RHS
    DDevice.SetLight LightIndex - 1, Lights(LightIndex)
End Property

Public Property Get Range() As Single
    Range = Lights(LightIndex).Range
End Property
Public Property Let Range(ByVal RHS As Single)
    Lights(LightIndex).Range = RHS
    DDevice.SetLight LightIndex - 1, Lights(LightIndex)
End Property

Public Property Get Form() As Long
    Form = Lights(LightIndex).Type
End Property
Public Property Let Form(ByVal RHS As Long)
    Lights(LightIndex).Type = RHS
    DDevice.SetLight LightIndex - 1, Lights(LightIndex)
End Property

Private Property Get Clone() As Orbit
    Set Clone = Orbit.Clone
    Set Clone = New Orbit
    With Clone
        .SunLight = SunLight
        .LightBlink = LightBlink
        .LightTimer = LightTimer
        .LightIsOn = LightIsOn
        .DiffuseRoll = DiffuseRoll
        .DiffuseTimer = DiffuseTimer
        .DIffuseMax = DIffuseMax
        .DiffuseNow = DiffuseNow
        
        .LightIndex = LightIndex
        .Key = Key
        .Serialize = Serialize
    
        If Not Diffuse Is Nothing Then
            Set .Diffuse = Diffuse.Clone
        End If
        If Not Ambience Is Nothing Then
            Set .Ambience = Ambience.Clone
        End If
        If Not Specular Is Nothing Then
            Set .Specular = Specular.Clone
        End If
        If Not Attenuation Is Nothing Then
            Set .Attenuation = Attenuation.Clone
        End If

    End With
End Property

Friend Sub UpdateValues()
    Lights(LightIndex).Diffuse.A = Diffuse.Alpha
    Lights(LightIndex).Diffuse.r = Diffuse.red
    Lights(LightIndex).Diffuse.g = Diffuse.green
    Lights(LightIndex).Diffuse.b = Diffuse.blue
 
    Lights(LightIndex).Ambient.A = Ambience.Alpha
    Lights(LightIndex).Ambient.r = Ambience.red
    Lights(LightIndex).Ambient.g = Ambience.green
    Lights(LightIndex).Ambient.b = Ambience.blue
 
    Lights(LightIndex).Specular.A = Specular.Alpha
    Lights(LightIndex).Specular.r = Specular.red
    Lights(LightIndex).Specular.g = Specular.green
    Lights(LightIndex).Specular.b = Specular.blue
 
    Lights(LightIndex).Attenuation0 = Attenuation.X
    Lights(LightIndex).Attenuation1 = Attenuation.Y
    Lights(LightIndex).Attenuation2 = Attenuation.Z

    Lights(LightIndex).Position.X = Origin.X
    Lights(LightIndex).Position.Y = Origin.Y
    Lights(LightIndex).Position.Z = Origin.Z
    
    Lights(LightIndex).Direction.X = Rotate.X
    Lights(LightIndex).Direction.Y = Rotate.Y
    Lights(LightIndex).Direction.Z = Rotate.Z
    
End Sub

Private Sub Class_Initialize()
    Stats_Brilliant_Count = Stats_Brilliant_Count + 1
    vbaObjSetAddref Orbit, ObjPtr(New Orbit)
    'Set Orbit = New Orbit
    
    Key = "K" & ObjPtr(Me)
    
    LightCount = LightCount + 1
    ReDim Preserve Lights(1 To LightCount) As D3DLIGHT8
    LightIndex = LightCount
    Set Diffuse = New Color
    Set Ambience = New Color
    Set Specular = New Color
    Set Attenuation = New Point

End Sub

Private Sub Class_Terminate()
    Stats_Brilliant_Count = Stats_Brilliant_Count - 1

'    If LightCount > 0 Then
'        If LightIndex < LightCount Then
'            Dim cnt As Long
'            For cnt = LightIndex + 1 To LightCount
'                Lights(cnt - 1) = Lights(cnt)
'            Next
'        End If
'        LightCount = LightCount - 1
'        If LightCount > 0 Then
'            ReDim Preserve Lights(1 To LightCount) As D3DLIGHT8
'            Dim brill As Object
'            For Each brill In Brilliants
'                If brill.LightIndex > LightIndex Then brill.LightIndex = brill.LightIndex - 1
'            Next
'        End If
'    End If
'    If StopGame And LightCount = 0 Then Erase Lights
    
    Set Diffuse = Nothing
    Set Ambience = Nothing
    Set Specular = Nothing
    Set Attenuation = Nothing
    
    'Set Orbit = Nothing
    vbaObjSet Orbit, ObjPtr(Nothing)
End Sub

Private Property Get Orbit_Clone() As Orbit
    Set Orbit_Clone = Orbit.Clone
End Property

Private Property Set Orbit_Motions(RHS As Motions)
    Set Orbit.Motions = RHS
End Property

Private Property Get Orbit_Motions() As Motions
    Set Orbit_Motions = Orbit.Motions
End Property

Private Property Let Orbit_Moved(ByVal RHS As Boolean)
    Orbit.Moved = RHS
End Property

Private Property Get Orbit_Moved() As Boolean
    Orbit_Moved = Orbit.Moved
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

Private Property Let Orbit_Serialize(ByVal RHS As Boolean)
    Orbit.Serialize = RHS
End Property

Private Property Get Orbit_Serialize() As Boolean
    Orbit_Serialize = Orbit.Serialize
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

