Attribute VB_Name = "modMove"

Option Explicit

Private Sub ApplyOrigin(ByRef ApplyTo As Molecule, ByRef Parent As Molecule, ByVal Relative As Boolean)

    If Not Relative Then
        Set ApplyTo.Origin = VectorDeduction(ApplyTo.Absolute.Origin, ApplyTo.Origin)
        Set ApplyTo.Absolute.Origin = ApplyTo.Origin
    Else
        Set ApplyTo.Origin = VectorAddition(VectorRotateAxis(ApplyTo.Relative.Origin, ApplyTo.Rotate), ApplyTo.Origin)
        Set ApplyTo.Absolute.Origin = ApplyTo.Origin
        Set ApplyTo.Relative.Origin = Nothing
        
      
'        If ApplyTo.Parent Is Nothing Then
'            Set ApplyTo.Origin = VectorAddition(VectorRotateAxis(ApplyTo.Relative.Origin, ApplyTo.Rotate), ApplyTo.Origin)
'        Else
'            Set ApplyTo.Origin = VectorAddition(VectorRotateAxis(VectorRotateAxis(ApplyTo.Relative.Origin, ApplyTo.Parent.Rotate), ApplyTo.Rotate), ApplyTo.Origin)
'        End If
'        Set ApplyTo.Absolute.Origin = ApplyTo.Origin
'        Set ApplyTo.Relative.Origin = Nothing
        
'        Set ApplyTo.Origin = VectorAddition(ApplyTo.Absolute.Origin, ApplyTo.Origin)
'        Set ApplyTo.Absolute.Origin = ApplyTo.Origin
'        Set ApplyTo.Relative.Origin = Nothing
    End If


End Sub


Private Sub ApplyRotate(ByRef ApplyTo As Molecule, ByRef Parent As Molecule, ByVal Relative As Boolean)

    If Not Relative Then
        Set ApplyTo.Rotate = AngleAxisDeduction(ApplyTo.Absolute.Rotate, ApplyTo.Rotate)
        Set ApplyTo.Absolute.Rotate = ApplyTo.Rotate
    Else
        Set ApplyTo.Rotate = AngleAxisAddition(ApplyTo.Relative.Rotate, ApplyTo.Rotate)
        Set ApplyTo.Absolute.Rotate = ApplyTo.Rotate
        Set ApplyTo.Relative.Rotate = Nothing
    End If

   

'    Static stacked As Integer
'    stacked = stacked + 1
'    Dim m As Molecule
'    If TypeName(ApplyTo) = "Planet" Then
'        For Each m In Molecules
'            If ApplyTo.Ranges.R = -1 Then
'                ApplyRotate VectorAddition(m.Rotate, Scalar), m, ApplyTo
'            ElseIf ApplyTo.Ranges.R > 0 Then
'                If ApplyTo.Ranges.R - Distance(m.Origin.X, m.Origin.Y, m.Origin.z, ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z) > 0 Then
'                    ApplyRotate  VectorAddition(m.Rotate, Scalar), m, ApplyTo
'                End If
'            End If
'        Next
'    ElseIf TypeName(ApplyTo) = "Molecule" Then
'        For Each m In ApplyTo.Molecules
'            ApplyRotate  VectorAddition(m.Rotate, Rotate), m, ApplyTo
'        Next
'    End If
'    stacked = stacked - 1
End Sub

Private Static Sub ApplyScaled(ByRef ApplyTo As Molecule, ByRef Parent As Molecule, ByVal Relative As Boolean)

    If Not Relative Then
        Set ApplyTo.Scaled = VectorDeduction(ApplyTo.Absolute.Scaled, ApplyTo.Scaled)
        Set ApplyTo.Absolute.Scaled = ApplyTo.Scaled
    Else
        Set ApplyTo.Scaled = VectorAddition(ApplyTo.Relative.Scaled, ApplyTo.Scaled)
        Set ApplyTo.Absolute.Scaled = ApplyTo.Scaled
        Set ApplyTo.Relative.Scaled = Nothing
    End If

'    Static stacked As Integer
'    stacked = stacked + 1
'    Dim m As Molecule
'    If TypeName(ApplyTo) = "Planet" Then
'        For Each m In Molecules
'            If ApplyTo.Ranges.R = -1 Then
'                ApplyScaled VectorAddition(m.Scaled, Scalar), m, ApplyTo
'            ElseIf ApplyTo.Ranges.R > 0 Then
'                If ApplyTo.Ranges.R - Distance(m.Origin.X, m.Origin.Y, m.Origin.z, ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z) > 0 Then
'                    ApplyScaled  VectorAddition(m.Scaled, Scalar), m, ApplyTo
'                End If
'            End If
'        Next
'    ElseIf TypeName(ApplyTo) = "Molecule" Then
'        For Each m In ApplyTo.Molecules
'            ApplyScaled  VectorAddition(m.Scaled, Scalar), m, ApplyTo
'        Next
'    End If
'    stacked = stacked - 1
End Sub

Private Static Sub ApplyOffset(ByRef ApplyTo As Molecule, ByRef Parent As Molecule, ByVal Relative As Boolean)

    If Not Relative Then
        Set ApplyTo.Offset = VectorDeduction(ApplyTo.Absolute.Offset, ApplyTo.Offset)
        Set ApplyTo.Absolute.Offset = ApplyTo.Offset
    Else
        Set ApplyTo.Offset = VectorAddition(ApplyTo.Relative.Offset, ApplyTo.Offset)
        Set ApplyTo.Absolute.Offset = ApplyTo.Offset
        Set ApplyTo.Relative.Offset = Nothing
    End If

'    Static stacked As Integer
'    stacked = stacked + 1
'    Dim m As Molecule
'    If TypeName(ApplyTo) = "Planet" Then
'        For Each m In Molecules
'            If ApplyTo.Ranges.R = -1 Then
'                ApplyOrigin  VectorAddition(ApplyTo.Origin, Offset), m, ApplyTo
'            ElseIf ApplyTo.Ranges.R > 0 Then
'                If ApplyTo.Ranges.R - Distance(m.Origin.X, m.Origin.Y, m.Origin.z, ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z) > 0 Then
'                    ApplyOrigin VectorAddition(ApplyTo.Origin, Offset), m, ApplyTo
'                End If
'            End If
'        Next
'    ElseIf TypeName(ApplyTo) = "Molecule" Then
'        For Each m In ApplyTo.Molecules
'            ApplyOrigin  VectorAddition(ApplyTo.Origin, Offset), m, ApplyTo
'        Next
'    End If
'    stacked = stacked - 1
End Sub

Public Sub Location(ByRef Origin As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    LocPos Origin, False, ApplyTo, Parent 'location is changing the origin to absolute
End Sub
Public Sub Position(ByRef Origin As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    LocPos Origin, True, ApplyTo, Parent 'position is changing the origin relative
End Sub
Public Sub CommitOrigin(ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
    CommitRoutine ApplyTo, Parent, False, False, True, False, True, False
    CommitRoutine ApplyTo, Parent, False, False, True, False, False, True
End Sub
Private Sub LocPos(ByRef Origin As Point, ByVal Relative As Boolean, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    'modifies the locale data of an objects properties, quickly in uncommitted change of multiple calls for speed consideration per frame
    If Origin.X <> 0 Or Origin.Y <> 0 Or Origin.Z <> 0 Then
        Dim o As Orbit
        Dim m As Molecule
        Select Case TypeName(ApplyTo)
            Case "Nothing"
                'go retrieve all planets whos range and origin has (0,0,0) with in it
                'and call change all molucules with in each of those planets as well
                Dim p As Planet
                For Each p In Planets
                    If ApplyTo.Ranges.r = -1 Then
                        LocPos Origin, Relative, p, ApplyTo
                    ElseIf ApplyTo.Ranges.r > 0 Then
                        If ApplyTo.Ranges.r <= Distance(p.Origin.X, p.Origin.Y, p.Origin.Z, 0, 0, 0) > 0 Then
                            LocPos Origin, Relative, p, ApplyTo
                        End If
                    End If
                Next
            Case "Planet", "Molecule"
                'early commit multiple calls per frame
                CommitOrigin ApplyTo, Parent
                If Relative Then
                    Set ApplyTo.Relative.Origin = Origin
                Else
                    Set ApplyTo.Absolute.Origin = Origin
                End If
                ApplyTo.Moved = True
                'change all molecules with in the specified planets range
'                If TypeName(ApplyTo) = "Planet" Then
''                    For Each m In Molecules
''                        If ApplyTo.Ranges.R = -1 Then
''                            LocPos Origin, Relative, m
''                        ElseIf ApplyTo.Ranges.R > 0 Then
''                            If ApplyTo.Ranges.R <= Distance(m.Origin.X, m.Origin.Y, m.Origin.z, _
''                                ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z) Then
''                                LocPos Origin, Relative, m
''                            End If
''                        End If
''                    Next
''                ElseIf TypeName(ApplyTo) = "Molecule" Then
''                    For Each m In ApplyTo.Molecules
''                        LocPos Origin, Relative, m
''                    Next
'                End If
        End Select
    End If
End Sub
Public Sub Rotation(ByRef Degrees As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    RotOri Degrees, False, ApplyTo, Parent 'location is changing the origin to absolute
End Sub
Public Sub Orientate(ByRef Degrees As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    RotOri Degrees, True, ApplyTo, Parent 'position is changing the origin relative
End Sub
Public Sub CommitRotate(ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
    CommitRoutine ApplyTo, Parent, True, False, False, False, True, False
    CommitRoutine ApplyTo, Parent, True, False, False, False, False, True
End Sub
Private Sub RotOri(ByRef Degrees As Point, ByVal Relative As Boolean, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    'modifies the locale data of an objects properties, quickly in uncommitted change of multiple calls for speed consideration per frame
    If Degrees.X <> 0 Or Degrees.Y <> 0 Or Degrees.Z <> 0 Then
        Dim m As Molecule
        Dim o As Point
        Select Case TypeName(ApplyTo)
            Case "Nothing"
                'go retrieve all planets whos range and origin has (0,0,0) with in it
                'and call change all molucules with in each of those planets as well
                Dim p As Planet
                For Each p In Planets
                    If ApplyTo.Ranges.r = -1 Then
                        RotOri Degrees, Relative, p, ApplyTo
                    ElseIf ApplyTo.Ranges.r > 0 Then
                        If ApplyTo.Ranges.r <= Distance(p.Origin.X, p.Origin.Y, p.Origin.Z, 0, 0, 0) > 0 Then
                            RotOri Degrees, Relative, p, ApplyTo
                        End If
                    End If
                Next
            Case "Planet", "Molecule"
                'early commit multiple calls per frame
                CommitRotate ApplyTo, Parent
                If Relative Then
                    Set ApplyTo.Relative.Rotate = Degrees
                Else
                    Set ApplyTo.Absolute.Rotate = Degrees
                End If
                ApplyTo.Moved = True
                'change all molecules with in the specified planets range
'                If TypeName(ApplyTo) = "Planet" Then
''                    For Each m In Molecules
''                        If ApplyTo.Ranges.R = -1 Then
''                            RotOri Degrees, Relative, m
''                        ElseIf ApplyTo.Ranges.R > 0 Then
''                            If ApplyTo.Ranges.R <= Distance(m.Origin.X, m.Origin.Y, m.Origin.z, _
''                                ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z) Then
''                                RotOri Degrees, Relative, m
''                            End If
''                        End If
''                    Next
''                ElseIf TypeName(ApplyTo) = "Molecule" Then
''                    For Each m In ApplyTo.Molecules
''                        RotOri Degrees, Relative, m
''                    Next
'                End If
        End Select
    End If
End Sub

Public Sub Scaling(ByRef Ratios As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    ScaExp Ratios, False, ApplyTo, Parent 'location is changing the origin to absolute
End Sub
Public Sub Explode(ByRef Ratios As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    ScaExp Ratios, True, ApplyTo, Parent 'position is changing the origin relative
End Sub
Public Sub CommitScaling(ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
    CommitRoutine ApplyTo, Parent, False, True, False, False, True, False
    CommitRoutine ApplyTo, Parent, False, True, False, False, False, True
End Sub
Private Sub ScaExp(ByRef Scalar As Point, ByVal Relative As Boolean, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    'modifies the locale data of an objects properties, quickly in uncommitted change of multiple calls for speed consideration per frame
    If Abs(Scalar.X) <> 1 Or Abs(Scalar.Y) <> 1 Or Abs(Scalar.Z) <> 1 Then
        Dim m As Molecule
        Dim o As Orbit
        Select Case TypeName(ApplyTo)
            Case "Nothing"
                'go retrieve all planets whos range and origin has (0,0,0) with in it
                'and call change all molucules with in each of those planets as well
                Dim p As Planet
                For Each p In Planets
                    If ApplyTo.Ranges.r = -1 Then
                        ScaExp Scalar, Relative, p, ApplyTo
                    ElseIf ApplyTo.Ranges.r > 0 Then
                        If ApplyTo.Ranges.r <= Distance(p.Origin.X, p.Origin.Y, p.Origin.Z, 0, 0, 0) > 0 Then
                            ScaExp Scalar, Relative, p, ApplyTo
                        End If
                    End If
                Next
            Case "Planet", "Molecule"
                'early commit multiple calls per frame
                CommitOffset ApplyTo, Parent
                If Relative Then
                    Set ApplyTo.Relative.Scaled = Scalar
                Else
                    Set ApplyTo.Absolute.Scaled = Scalar
                End If
                ApplyTo.Moved = True
                'change all molecules with in the specified planets range
'                If TypeName(ApplyTo) = "Planet" Then
''                    For Each m In Molecules
''                        If ApplyTo.Ranges.R = -1 Then
''                            ScaExp Scalar, Relative, m
''                        ElseIf ApplyTo.Ranges.R > 0 Then
''                            If ApplyTo.Ranges.R <= Distance(m.Origin.X, m.Origin.Y, m.Origin.z, _
''                                ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z) Then
''                                ScaExp Scalar, Relative, m
''                            End If
''                        End If
''                    Next
''                ElseIf TypeName(ApplyTo) = "Molecule" Then
''                    For Each m In ApplyTo.Molecules
''                        ScaExp Scalar, Relative, m
''                    Next
'                End If
        End Select
    End If
End Sub
Public Sub Displace(ByRef Offset As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    DisBal Offset, False, ApplyTo, Parent  'location is changing the origin to absolute
End Sub
Public Sub Balanced(ByRef Offset As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    DisBal Offset, True, ApplyTo, Parent 'position is changing the origin relative
End Sub
Public Sub CommitOffset(ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
    CommitRoutine ApplyTo, Parent, False, False, False, True, True, False
    CommitRoutine ApplyTo, Parent, False, False, False, True, False, True
End Sub
Private Sub DisBal(ByRef Offset As Point, ByVal Relative As Boolean, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
    'modifies the locale data of an objects properties, quickly in uncommitted change of multiple calls for speed consideration per frame
    If Offset.X <> 0 Or Offset.Y <> 0 Or Offset.Z <> 0 Then
        Dim dist As Single
        Dim m As Molecule
        Dim o As Orbit
        Select Case TypeName(ApplyTo)
            Case "Nothing"
                'go retrieve all planets whos range and origin has (0,0,0) with in it
                'and call change all molucules with in each of those planets as well
                Dim p As Planet
                For Each p In Planets
                    If ApplyTo.Ranges.r = -1 Then
                        DisBal Offset, Relative, p, ApplyTo
                    ElseIf ApplyTo.Ranges.r > 0 Then
                        If ApplyTo.Ranges.r <= Distance(p.Origin.X, p.Origin.Y, p.Origin.Z, 0, 0, 0) > 0 Then
                            DisBal Offset, Relative, p, ApplyTo
                        End If
                    End If
                Next
            Case "Planet", "Molecule"
                'early commit multiple calls per frame
                CommitOffset ApplyTo, Parent
                If Relative Then
                    Set ApplyTo.Relative.Offset = Offset
                Else
                    Set ApplyTo.Absolute.Offset = Offset
                End If
                ApplyTo.Moved = True
                'change all molecules with in the specified planets range
'                If TypeName(ApplyTo) = "Planet" Then
''                    For Each m In Molecules
''                        If ApplyTo.Ranges.R = -1 Then
''                            DisBal Offset, Relative, m
''                        ElseIf ApplyTo.Ranges.R > 0 Then
''                            If ApplyTo.Ranges.R <= Distance(m.Origin.X, m.Origin.Y, m.Origin.z, _
''                                ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z) Then
''                                DisBal Offset, Relative, m
''                            End If
''                        End If
''                    Next
''                ElseIf TypeName(ApplyTo) = "Molecule" Then
''                    For Each m In ApplyTo.Molecules
''                        DisBal Offset, Relative, m
''                    Next
'                End If
        End Select
    End If
End Sub

Private Function RangedMolecules(ByRef ApplyTo As Molecule) As NTNodes10.Collection
    Set RangedMolecules = New NTNodes10.Collection
    Dim m As Molecule

    Dim dist As Single
    For Each m In Molecules
        If ((m.Parent Is Nothing) And (Not TypeName(ApplyTo) = "Planet")) Or (TypeName(ApplyTo) = "Planet") Then
            If ApplyTo.Ranges.r = -1 Then
                RangedMolecules.Add m, m.Key
            ElseIf ApplyTo.Ranges.r > 0 Then
                dist = Distance(m.Origin.X, m.Origin.Y, m.Origin.Z, ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.Z)
                If ApplyTo.Ranges.r <= dist Then
                    RangedMolecules.Add m, m.Key
                End If
            End If
        End If
    Next
    For Each m In ApplyTo.Molecules
        If Not RangedMolecules.Exists(m.Key) Then RangedMolecules.Add m, m.Key
    Next
End Function


Public Sub CommitRoutine(ByRef ApplyTo As Molecule, ByRef Parent As Molecule, ByVal DoRotate As Boolean, ByVal DoScaled As Boolean, ByVal DoOrigin As Boolean, ByVal DoOffset As Boolean, ByVal DoAbsolute As Boolean, ByVal DoRelative As Boolean)
    'partial to committing a 3d objects properties during calls that may not sum, for retaining other properties needing change first and entirety per frame
    Static stacked As Boolean
    If Not stacked Then
        stacked = True
        
        'any absolute position comes first, pending is a difference from the actual
        If ((DoOrigin And DoAbsolute) Or ((Not DoOrigin) And (Not DoAbsolute))) Then
            If (Not ApplyTo.Absolute.Origin.Equals(Nothing)) Then ' And (Not ApplyTo.Moved) Then
                If (Not ApplyTo.Origin.Equals(ApplyTo.Absolute.Origin)) Then
                    ApplyOrigin ApplyTo, Parent, False
                    ApplyTo.Moved = True
                End If
            End If
        End If
        If ((DoOffset And DoAbsolute) Or ((Not DoOffset) And (Not DoAbsolute))) Then
            If (Not ApplyTo.Absolute.Offset.Equals(Nothing)) Then 'And (Not ApplyTo.Moved) Then
                If (Not ApplyTo.Offset.Equals(ApplyTo.Absolute.Offset)) Then
                    ApplyOffset ApplyTo, Parent, False
                    ApplyTo.Moved = True
                End If
            End If
        End If
        If ((DoRotate And DoAbsolute) Or ((Not DoRotate) And (Not DoAbsolute))) Then
            If (Not ApplyTo.Absolute.Rotate.Equals(Nothing)) Then 'And (Not ApplyTo.Moved) Then
                If (Not ApplyTo.Rotate.Equals(ApplyTo.Absolute.Rotate)) Then
                    ApplyRotate ApplyTo, Parent, False
                    ApplyTo.Moved = True
                End If
            End If
        End If
        If ((DoScaled And DoAbsolute) Or ((Not DoScaled) And (Not DoAbsolute))) Then
            If (Not ApplyTo.Absolute.Scaled.Equals(Nothing)) Then 'And (Not ApplyTo.Moved) Then
                If (Not ApplyTo.Scaled.Equals(ApplyTo.Absolute.Scaled)) Then
                    ApplyScaled ApplyTo, Parent, False
                    ApplyTo.Moved = True
                End If
            End If
        End If


        'relative positioning comes secondly, pending is there is any value not empty
        If ((DoRotate And DoRelative) Or ((Not DoRotate) And (Not DoRelative))) Then
            If (Not ApplyTo.Relative.Rotate.Equals(Nothing)) Then 'And (Not ApplyTo.Moved) Then
                If (ApplyTo.Relative.Rotate.X <> 0 Or ApplyTo.Relative.Rotate.Y <> 0 Or ApplyTo.Relative.Rotate.Z <> 0) Then
                    ApplyRotate ApplyTo, Parent, True
                    ApplyTo.Moved = True
                End If
            End If
        End If
        If ((DoOrigin And DoRelative) Or ((Not DoOrigin) And (Not DoRelative))) Then
            If (Not ApplyTo.Relative.Origin.Equals(Nothing)) Then 'And (Not ApplyTo.Moved) Then
                If (ApplyTo.Relative.Origin.X <> 0 Or ApplyTo.Relative.Origin.Y <> 0 Or ApplyTo.Relative.Origin.Z <> 0) Then
                    ApplyOrigin ApplyTo, Parent, True
                    ApplyTo.Moved = True
                End If
            End If
        End If
        If ((DoOffset And DoRelative) Or ((Not DoOffset) And (Not DoRelative))) Then
            If (Not ApplyTo.Relative.Offset.Equals(Nothing)) Then ' And (Not ApplyTo.Moved) Then
                If (ApplyTo.Relative.Offset.X <> 0 Or ApplyTo.Relative.Offset.Y <> 0 Or ApplyTo.Relative.Offset.Z <> 0) Then
                    ApplyOffset ApplyTo, Parent, True
                    ApplyTo.Moved = True
                End If
            End If
        End If
        If ((DoScaled And DoRelative) Or ((Not DoScaled) And (Not DoRelative))) Then
            If (Not ApplyTo.Relative.Scaled.Equals(Nothing)) Then 'And (Not ApplyTo.Moved) Then
                If (Abs(ApplyTo.Relative.Scaled.X) <> 1 Or Abs(ApplyTo.Relative.Scaled.Y) <> 1 Or Abs(ApplyTo.Relative.Scaled.Z) <> 1) Then
                    ApplyScaled ApplyTo, Parent, True
                    ApplyTo.Moved = True
                End If
            End If
        End If
                
        stacked = False
    End If
End Sub
Private Sub AllCommitRoutine(ByRef ApplyTo As Molecule, Optional ByRef Parent As Molecule = Nothing)

    CommitRoutine ApplyTo, Parent, True, False, False, False, True, False
    CommitRoutine ApplyTo, Parent, False, True, False, False, True, False
    CommitRoutine ApplyTo, Parent, False, False, True, False, True, False
    CommitRoutine ApplyTo, Parent, False, False, False, True, True, False

    CommitRoutine ApplyTo, Parent, True, False, False, False, False, True
    CommitRoutine ApplyTo, Parent, False, True, False, False, False, True
    CommitRoutine ApplyTo, Parent, False, False, True, False, False, True
    CommitRoutine ApplyTo, Parent, False, False, False, True, False, True

    Set ApplyTo.Relative = Nothing
End Sub

Public Sub RenderMotions(ByRef UserControl As Macroscopic, ByRef Camera As Camera)
    'called once per frame committing changes the last frame has waiting in object properties in entirety
    Dim m As Molecule
    Dim p As Planet
    
    For Each m In Molecules
        If m.Parent Is Nothing Then
            AllCommitRoutine m, Nothing
        End If
    Next
    
    
    For Each p In Planets
        AllCommitRoutine p, Nothing
    Next

    
End Sub


'Option Explicit
'
'Private Sub ApplyOrigin(ByRef ApplyTo As Molecule, ByRef Parent As Molecule, ByVal Relative As Boolean)
'
'    If Not Relative Then
'        Set ApplyTo.Origin = VectorDeduction(ApplyTo.Absolute.Origin, ApplyTo.Origin)
'        Set ApplyTo.Absolute.Origin = ApplyTo.Origin
'    Else
'        Set ApplyTo.Origin = VectorAddition(VectorRotateAxis(ApplyTo.Relative.Origin, ApplyTo.Rotate), ApplyTo.Origin)
'        Set ApplyTo.Absolute.Origin = ApplyTo.Origin
'        Set ApplyTo.Relative.Origin = Nothing
'    End If
'
''    Static stacked As Integer
''    stacked = stacked + 1
''    Dim m As Molecule
''    If TypeName(ApplyTo) = "Planet" Then
''        For Each m In Molecules
''            If ApplyTo.Ranges.R = -1 Then
''                ApplyOrigin m, ApplyTo, Relative
''            ElseIf ApplyTo.Ranges.R > 0 Then
''                If ApplyTo.Ranges.R - Distance(m.Origin.X, m.Origin.Y, m.Origin.z, ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z) > 0 Then
''                    ApplyOrigin m, ApplyTo, Relative
''                End If
''            End If
''        Next
''    ElseIf TypeName(ApplyTo) = "Molecule" Then
''        For Each m In ApplyTo.Molecules
''            ApplyOrigin m, ApplyTo, Relative
''        Next
''    End If
''    stacked = stacked - 1
'End Sub
'
'
'Private Sub ApplyRotate(ByRef ApplyTo As Molecule, ByRef Parent As Molecule, ByVal Relative As Boolean)
'
'    If Not Relative Then
'        Set ApplyTo.Rotate = AngleAxisDeduction(ApplyTo.Absolute.Rotate, ApplyTo.Rotate)
'        Set ApplyTo.Absolute.Rotate = ApplyTo.Rotate
'    Else
'        Set ApplyTo.Rotate = AngleAxisAddition(ApplyTo.Relative.Rotate, ApplyTo.Rotate)
'        Set ApplyTo.Absolute.Rotate = ApplyTo.Rotate
'        Set ApplyTo.Relative.Rotate = Nothing
'    End If
'
'
'End Sub
'
'Private Static Sub ApplyScaled(ByRef ApplyTo As Molecule, ByRef Parent As Molecule, ByVal Relative As Boolean)
'
'    If Not Relative Then
'        Set ApplyTo.Scaled = VectorDeduction(ApplyTo.Absolute.Scaled, ApplyTo.Scaled)
'        Set ApplyTo.Absolute.Scaled = ApplyTo.Scaled
'    Else
'        Set ApplyTo.Scaled = VectorAddition(ApplyTo.Relative.Scaled, ApplyTo.Scaled)
'        Set ApplyTo.Absolute.Scaled = ApplyTo.Scaled
'        Set ApplyTo.Relative.Scaled = Nothing
'    End If
'
''    Static stacked As Integer
''    stacked = stacked + 1
''    Dim m As Molecule
''    If TypeName(ApplyTo) = "Planet" Then
''        For Each m In Molecules
''            If ApplyTo.Ranges.R = -1 Then
''                ApplyScaled VectorAddition(m.Scaled, Scalar), m, ApplyTo
''            ElseIf ApplyTo.Ranges.R > 0 Then
''                If ApplyTo.Ranges.R - Distance(m.Origin.X, m.Origin.Y, m.Origin.z, ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z) > 0 Then
''                    ApplyScaled  VectorAddition(m.Scaled, Scalar), m, ApplyTo
''                End If
''            End If
''        Next
''    ElseIf TypeName(ApplyTo) = "Molecule" Then
''        For Each m In ApplyTo.Molecules
''            ApplyScaled  VectorAddition(m.Scaled, Scalar), m, ApplyTo
''        Next
''    End If
''    stacked = stacked - 1
'End Sub
'
'Private Static Sub ApplyOffset(ByRef ApplyTo As Molecule, ByRef Parent As Molecule, ByVal Relative As Boolean)
'
'    If Not Relative Then
'        Set ApplyTo.Offset = VectorDeduction(ApplyTo.Absolute.Offset, ApplyTo.Offset)
'        Set ApplyTo.Absolute.Offset = ApplyTo.Offset
'    Else
'        Set ApplyTo.Offset = VectorAddition(ApplyTo.Relative.Offset, ApplyTo.Offset)
'        Set ApplyTo.Absolute.Offset = ApplyTo.Offset
'        Set ApplyTo.Relative.Offset = Nothing
'    End If
'
''    Static stacked As Integer
''    stacked = stacked + 1
''    Dim m As Molecule
''    If TypeName(ApplyTo) = "Planet" Then
''        For Each m In Molecules
''            If ApplyTo.Ranges.R = -1 Then
''                ApplyOrigin  VectorAddition(ApplyTo.Origin, Offset), m, ApplyTo
''            ElseIf ApplyTo.Ranges.R > 0 Then
''                If ApplyTo.Ranges.R - Distance(m.Origin.X, m.Origin.Y, m.Origin.z, ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z) > 0 Then
''                    ApplyOrigin VectorAddition(ApplyTo.Origin, Offset), m, ApplyTo
''                End If
''            End If
''        Next
''    ElseIf TypeName(ApplyTo) = "Molecule" Then
''        For Each m In ApplyTo.Molecules
''            ApplyOrigin  VectorAddition(ApplyTo.Origin, Offset), m, ApplyTo
''        Next
''    End If
''    stacked = stacked - 1
'End Sub
'
'Public Sub Location(ByRef Origin As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
'    LocPos Origin, False, ApplyTo, Parent 'location is changing the origin to absolute
'End Sub
'Public Sub Position(ByRef Origin As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
'    LocPos Origin, True, ApplyTo, Parent 'position is changing the origin relative
'End Sub
'Public Sub CommitOrigin(ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
'    CommitRoutine ApplyTo, Parent, False, False, True, False, True, False
'    CommitRoutine ApplyTo, Parent, False, False, True, False, False, True
'End Sub
'Private Sub LocPos(ByRef Origin As Point, ByVal Relative As Boolean, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
'    'modifies the locale data of an objects properties, quickly in uncommitted change of multiple calls for speed consideration per frame
'    If Origin.X <> 0 Or Origin.Y <> 0 Or Origin.Z <> 0 Then
'        Dim o As Orbit
'        Dim m As Molecule
'        Select Case TypeName(ApplyTo)
'            Case "Nothing"
'                'go retrieve all planets whos range and origin has (0,0,0) with in it
'                'and call change all molucules with in each of those planets as well
'                Dim p As Planet
'                For Each p In Planets
'                    If ApplyTo.Ranges.r = -1 Then
'                        LocPos Origin, Relative, p, ApplyTo
'                    ElseIf ApplyTo.Ranges.r > 0 Then
'                        If ApplyTo.Ranges.r <= Distance(p.Origin.X, p.Origin.Y, p.Origin.Z, 0, 0, 0) > 0 Then
'                            LocPos Origin, Relative, p, ApplyTo
'                        End If
'                    End If
'                Next
'            Case "Planet", "Molecule"
'                'early commit multiple calls per frame
'                CommitOrigin ApplyTo, Parent
'                If Relative Then
'                    Set ApplyTo.Relative.Origin = Origin
'                Else
'                    Set ApplyTo.Absolute.Origin = Origin
'                End If
'                'change all molecules with in the specified planets range
''                If TypeName(ApplyTo) = "Planet" Then
'''                    For Each m In Molecules
'''                        If ApplyTo.Ranges.R = -1 Then
'''                            LocPos Origin, Relative, m
'''                        ElseIf ApplyTo.Ranges.R > 0 Then
'''                            If ApplyTo.Ranges.R <= Distance(m.Origin.X, m.Origin.Y, m.Origin.z, _
'''                                ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z) Then
'''                                LocPos Origin, Relative, m
'''                            End If
'''                        End If
'''                    Next
'''                ElseIf TypeName(ApplyTo) = "Molecule" Then
'''                    For Each m In ApplyTo.Molecules
'''                        LocPos Origin, Relative, m
'''                    Next
''                End If
'        End Select
'    End If
'End Sub
'Public Sub Rotation(ByRef Degrees As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
'    RotOri Degrees, False, ApplyTo, Parent 'location is changing the origin to absolute
'End Sub
'Public Sub Orientate(ByRef Degrees As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
'    RotOri Degrees, True, ApplyTo, Parent 'position is changing the origin relative
'End Sub
'Public Sub CommitRotate(ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
'    CommitRoutine ApplyTo, Parent, True, False, False, False, True, False
'    CommitRoutine ApplyTo, Parent, True, False, False, False, False, True
'End Sub
'Private Sub RotOri(ByRef Degrees As Point, ByVal Relative As Boolean, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
'    'modifies the locale data of an objects properties, quickly in uncommitted change of multiple calls for speed consideration per frame
'    If Degrees.X <> 0 Or Degrees.Y <> 0 Or Degrees.Z <> 0 Then
'        Dim m As Molecule
'        Dim o As Point
'        Select Case TypeName(ApplyTo)
'            Case "Nothing"
'                'go retrieve all planets whos range and origin has (0,0,0) with in it
'                'and call change all molucules with in each of those planets as well
'                Dim p As Planet
'                For Each p In Planets
'                    If ApplyTo.Ranges.r = -1 Then
'                        RotOri Degrees, Relative, p, ApplyTo
'                    ElseIf ApplyTo.Ranges.r > 0 Then
'                        If ApplyTo.Ranges.r <= Distance(p.Origin.X, p.Origin.Y, p.Origin.Z, 0, 0, 0) > 0 Then
'                            RotOri Degrees, Relative, p, ApplyTo
'                        End If
'                    End If
'                Next
'            Case "Planet", "Molecule"
'                'early commit multiple calls per frame
'                CommitRotate ApplyTo, Parent
'                If Relative Then
'                    Set ApplyTo.Relative.Rotate = Degrees
'                Else
'                    Set ApplyTo.Absolute.Rotate = Degrees
'                End If
'                'change all molecules with in the specified planets range
''                If TypeName(ApplyTo) = "Planet" Then
'''                    For Each m In Molecules
'''                        If ApplyTo.Ranges.R = -1 Then
'''                            RotOri Degrees, Relative, m
'''                        ElseIf ApplyTo.Ranges.R > 0 Then
'''                            If ApplyTo.Ranges.R <= Distance(m.Origin.X, m.Origin.Y, m.Origin.z, _
'''                                ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z) Then
'''                                RotOri Degrees, Relative, m
'''                            End If
'''                        End If
'''                    Next
'''                ElseIf TypeName(ApplyTo) = "Molecule" Then
'''                    For Each m In ApplyTo.Molecules
'''                        RotOri Degrees, Relative, m
'''                    Next
''                End If
'        End Select
'    End If
'End Sub
'
'Public Sub Scaling(ByRef Ratios As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
'    ScaExp Ratios, False, ApplyTo, Parent 'location is changing the origin to absolute
'End Sub
'Public Sub Explode(ByRef Ratios As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
'    ScaExp Ratios, True, ApplyTo, Parent 'position is changing the origin relative
'End Sub
'Public Sub CommitScaling(ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
'    CommitRoutine ApplyTo, Parent, False, True, False, False, True, False
'    CommitRoutine ApplyTo, Parent, False, True, False, False, False, True
'End Sub
'Private Sub ScaExp(ByRef Scalar As Point, ByVal Relative As Boolean, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
'    'modifies the locale data of an objects properties, quickly in uncommitted change of multiple calls for speed consideration per frame
'    If Abs(Scalar.X) <> 1 Or Abs(Scalar.Y) <> 1 Or Abs(Scalar.Z) <> 1 Then
'        Dim m As Molecule
'        Dim o As Orbit
'        Select Case TypeName(ApplyTo)
'            Case "Nothing"
'                'go retrieve all planets whos range and origin has (0,0,0) with in it
'                'and call change all molucules with in each of those planets as well
'                Dim p As Planet
'                For Each p In Planets
'                    If ApplyTo.Ranges.r = -1 Then
'                        ScaExp Scalar, Relative, p, ApplyTo
'                    ElseIf ApplyTo.Ranges.r > 0 Then
'                        If ApplyTo.Ranges.r <= Distance(p.Origin.X, p.Origin.Y, p.Origin.Z, 0, 0, 0) > 0 Then
'                            ScaExp Scalar, Relative, p, ApplyTo
'                        End If
'                    End If
'                Next
'            Case "Planet", "Molecule"
'                'early commit multiple calls per frame
'                CommitOffset ApplyTo, Parent
'                If Relative Then
'                    Set ApplyTo.Relative.Scaled = Scalar
'                Else
'                    Set ApplyTo.Absolute.Scaled = Scalar
'                End If
'                'change all molecules with in the specified planets range
''                If TypeName(ApplyTo) = "Planet" Then
'''                    For Each m In Molecules
'''                        If ApplyTo.Ranges.R = -1 Then
'''                            ScaExp Scalar, Relative, m
'''                        ElseIf ApplyTo.Ranges.R > 0 Then
'''                            If ApplyTo.Ranges.R <= Distance(m.Origin.X, m.Origin.Y, m.Origin.z, _
'''                                ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z) Then
'''                                ScaExp Scalar, Relative, m
'''                            End If
'''                        End If
'''                    Next
'''                ElseIf TypeName(ApplyTo) = "Molecule" Then
'''                    For Each m In ApplyTo.Molecules
'''                        ScaExp Scalar, Relative, m
'''                    Next
''                End If
'        End Select
'    End If
'End Sub
'Public Sub Displace(ByRef Offset As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
'    DisBal Offset, False, ApplyTo, Parent  'location is changing the origin to absolute
'End Sub
'Public Sub Balanced(ByRef Offset As Point, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
'    DisBal Offset, True, ApplyTo, Parent 'position is changing the origin relative
'End Sub
'Public Sub CommitOffset(ByRef ApplyTo As Molecule, ByRef Parent As Molecule)
'    CommitRoutine ApplyTo, Parent, False, False, False, True, True, False
'    CommitRoutine ApplyTo, Parent, False, False, False, True, False, True
'End Sub
'Private Sub DisBal(ByRef Offset As Point, ByVal Relative As Boolean, Optional ByRef ApplyTo As Molecule = Nothing, Optional ByRef Parent As Molecule = Nothing)
'    'modifies the locale data of an objects properties, quickly in uncommitted change of multiple calls for speed consideration per frame
'    If Offset.X <> 0 Or Offset.Y <> 0 Or Offset.Z <> 0 Then
'        Dim dist As Single
'        Dim m As Molecule
'        Dim o As Orbit
'        Select Case TypeName(ApplyTo)
'            Case "Nothing"
'                'go retrieve all planets whos range and origin has (0,0,0) with in it
'                'and call change all molucules with in each of those planets as well
'                Dim p As Planet
'                For Each p In Planets
'                    If ApplyTo.Ranges.r = -1 Then
'                        DisBal Offset, Relative, p, ApplyTo
'                    ElseIf ApplyTo.Ranges.r > 0 Then
'                        If ApplyTo.Ranges.r <= Distance(p.Origin.X, p.Origin.Y, p.Origin.Z, 0, 0, 0) > 0 Then
'                            DisBal Offset, Relative, p, ApplyTo
'                        End If
'                    End If
'                Next
'            Case "Planet", "Molecule"
'                'early commit multiple calls per frame
'                CommitOffset ApplyTo, Parent
'                If Relative Then
'                    Set ApplyTo.Relative.Offset = Offset
'                Else
'                    Set ApplyTo.Absolute.Offset = Offset
'                End If
'                'change all molecules with in the specified planets range
''                If TypeName(ApplyTo) = "Planet" Then
'''                    For Each m In Molecules
'''                        If ApplyTo.Ranges.R = -1 Then
'''                            DisBal Offset, Relative, m
'''                        ElseIf ApplyTo.Ranges.R > 0 Then
'''                            If ApplyTo.Ranges.R <= Distance(m.Origin.X, m.Origin.Y, m.Origin.z, _
'''                                ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.z) Then
'''                                DisBal Offset, Relative, m
'''                            End If
'''                        End If
'''                    Next
'''                ElseIf TypeName(ApplyTo) = "Molecule" Then
'''                    For Each m In ApplyTo.Molecules
'''                        DisBal Offset, Relative, m
'''                    Next
''                End If
'        End Select
'    End If
'End Sub
'
'Private Function RangedMolecules(ByRef ApplyTo As Molecule) As NTNodes10.Collection
'    Set RangedMolecules = New NTNodes10.Collection
'    Dim m As Molecule
'
'    Dim dist As Single
'    For Each m In Molecules
'        If ((m.Parent Is Nothing) And (Not TypeName(ApplyTo) = "Planet")) Or (TypeName(ApplyTo) = "Planet") Then
'            If ApplyTo.Ranges.r = -1 Then
'                RangedMolecules.Add m, m.Key
'            ElseIf ApplyTo.Ranges.r > 0 Then
'                dist = Distance(m.Origin.X, m.Origin.Y, m.Origin.Z, ApplyTo.Origin.X, ApplyTo.Origin.Y, ApplyTo.Origin.Z)
'                If ApplyTo.Ranges.r <= dist Then
'                    RangedMolecules.Add m, m.Key
'                End If
'            End If
'        End If
'    Next
'    For Each m In ApplyTo.Molecules
'        If Not RangedMolecules.Exists(m.Key) Then RangedMolecules.Add m, m.Key
'    Next
'End Function
'
'Public Sub CommitRoutine(ByRef ApplyTo As Molecule, ByRef Parent As Molecule, ByVal DoRotate As Boolean, ByVal DoScaled As Boolean, ByVal DoOrigin As Boolean, ByVal DoOffset As Boolean, ByVal DoAbsolute As Boolean, ByVal DoRelative As Boolean)
'    'partial to committing a 3d objects properties during calls that may not sum, for retaining other properties needing change first and entirety per frame
'    Static stacked As Boolean
'    If Not stacked Then
'        stacked = True
'
'        'any absolute position comes first, pending is a difference from the actual
'        If Not ApplyTo.Absolute.Origin.Equals(Nothing) Then
'            If (Not ApplyTo.Origin.Equals(ApplyTo.Absolute.Origin)) And ((DoOrigin And DoAbsolute) Or ((Not DoOrigin) And (Not DoAbsolute))) Then
'                ApplyOrigin ApplyTo, Parent, False
'            End If
'        End If
'        If Not ApplyTo.Absolute.Offset.Equals(Nothing) Then
'            If (Not ApplyTo.Offset.Equals(ApplyTo.Absolute.Offset)) And ((DoOffset And DoAbsolute) Or ((Not DoOffset) And (Not DoAbsolute))) Then
'                ApplyOffset ApplyTo, Parent, False
'            End If
'        End If
'        If Not ApplyTo.Absolute.Rotate.Equals(Nothing) Then
'            If (Not ApplyTo.Rotate.Equals(ApplyTo.Absolute.Rotate)) And ((DoRotate And DoAbsolute) Or ((Not DoRotate) And (Not DoAbsolute))) Then
'                ApplyRotate ApplyTo, Parent, False
'            End If
'        End If
'        If Not ApplyTo.Absolute.Scaled.Equals(Nothing) Then
'            If (Not ApplyTo.Scaled.Equals(ApplyTo.Absolute.Scaled)) And ((DoScaled And DoAbsolute) Or ((Not DoScaled) And (Not DoAbsolute))) Then
'                ApplyScaled ApplyTo, Parent, False
'            End If
'        End If
'
'
'        'relative positioning comes secondly, pending is there is any value not empty
'        If Not ApplyTo.Relative.Rotate.Equals(Nothing) Then
'            If (ApplyTo.Relative.Rotate.X <> 0 Or ApplyTo.Relative.Rotate.Y <> 0 Or ApplyTo.Relative.Rotate.Z <> 0) And ((DoRotate And DoRelative) Or ((Not DoRotate) And (Not DoRelative))) Then
'                ApplyRotate ApplyTo, Parent, True
'            End If
'        End If
'        If Not ApplyTo.Relative.Origin.Equals(Nothing) Then
'            If (ApplyTo.Relative.Origin.X <> 0 Or ApplyTo.Relative.Origin.Y <> 0 Or ApplyTo.Relative.Origin.Z <> 0) And ((DoOrigin And DoRelative) Or ((Not DoOrigin) And (Not DoRelative))) Then
'                ApplyOrigin ApplyTo, Parent, True
'            End If
'        End If
'        If Not ApplyTo.Relative.Offset.Equals(Nothing) Then
'            If (ApplyTo.Relative.Offset.X <> 0 Or ApplyTo.Relative.Offset.Y <> 0 Or ApplyTo.Relative.Offset.Z <> 0) And ((DoOffset And DoRelative) Or ((Not DoOffset) And (Not DoRelative))) Then
'                ApplyOffset ApplyTo, Parent, True
'            End If
'        End If
'        If Not ApplyTo.Relative.Scaled.Equals(Nothing) Then
'            If (Abs(ApplyTo.Relative.Scaled.X) <> 1 Or Abs(ApplyTo.Relative.Scaled.Y) <> 1 Or Abs(ApplyTo.Relative.Scaled.Z) <> 1) And ((DoScaled And DoRelative) Or ((Not DoScaled) And (Not DoRelative))) Then
'                ApplyScaled ApplyTo, Parent, True
'            End If
'        End If
'
'        stacked = False
'    End If
'End Sub
'Private Sub AllCommitRoutine(ByRef ApplyTo As Molecule, Optional ByRef Parent As Molecule = Nothing)
'
'    CommitRoutine ApplyTo, Parent, True, False, False, False, True, False
'    CommitRoutine ApplyTo, Parent, False, True, False, False, True, False
'    CommitRoutine ApplyTo, Parent, False, False, True, False, True, False
'    CommitRoutine ApplyTo, Parent, False, False, False, True, True, False
'
'    CommitRoutine ApplyTo, Parent, True, False, False, False, False, True
'    CommitRoutine ApplyTo, Parent, False, True, False, False, False, True
'    CommitRoutine ApplyTo, Parent, False, False, True, False, False, True
'    CommitRoutine ApplyTo, Parent, False, False, False, True, False, True
'
'    Set ApplyTo.Relative = Nothing
'End Sub
'
'Public Sub RenderMotions(ByRef UserControl As Macroscopic, ByRef Camera As Camera)
'    'called once per frame committing changes the last frame has waiting in object properties in entirety
'    Dim m As Molecule
'    Dim p As Planet
'
'    For Each m In Molecules
'        If m.Parent Is Nothing Then
'            AllCommitRoutine m, Nothing
'        End If
'    Next
'
'
'    For Each p In Planets
'        AllCommitRoutine p, Nothing
'    Next
'
'
'
'
'
'
'End Sub
'
'
'
