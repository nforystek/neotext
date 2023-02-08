VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Camera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pColor As Long
Private pPlayer As Molecule
Private pPlanet As Planet

Private pSerialize As Boolean

Public Property Get Player() As Molecule
    If pPlayer Is Nothing Then Set pPlayer = New Molecule
    Set Player = pPlayer
End Property
Public Property Set Player(ByRef RHS As Molecule)
    Set pPlayer = RHS
End Property

Public Property Get Planet() As Planet
    Set Planet = pPlanet
End Property
Public Property Set Planet(ByRef RHS As Planet)
    Set pPlanet = RHS
End Property

Public Property Get Serialize()
    If pSerialize Then
        Serialize = ""
        If Not pPlayer Is Nothing Then
            Serialize = Serialize & "Set Camera.Player = Molecules(""" & pPlayer.Key & """)" & vbCrLf
            Serialize = Serialize & "Camera.Player.Origin = """ & pPlayer.Origin & """" & vbCrLf
            Serialize = Serialize & "Camera.Player.Rotate = """ & pPlayer.Rotate & """" & vbCrLf
        End If
        If Not pPlanet Is Nothing Then
            Serialize = Serialize & "Set Camera.Planet = Planets(""" & pPlanet.Key & """)" & vbCrLf
        End If
        Serialize = Serialize & "Camera.Color = " & pColor & vbCrLf
        Serialize = Serialize & "Camera.Serialize = True" & vbCrLf
    Else
        Serialize = False
    End If
End Property

Public Property Let Serialize(ByVal RHS)
    If TypeName(RHS) = "Boolean" Then
        pSerialize = CBool(RHS)
    End If
End Property


Public Property Set Serialize(ByRef RHS)
    If TypeName(RHS) = "Molecule" Then
        Set pPlayer = RHS
    ElseIf TypeName(RHS) = "Planet" Then
        Set pPlanet = RHS
    End If
End Property

Public Property Let Color(ByVal RHS As Long)
    pColor = RHS
End Property
Public Property Get Color() As Long
    If pPlayer Is Nothing Then
        Color = pColor
    Else
'        If pPlanet Is Nothing Then
    
    
            Color = D3DColorARGB(1, 0, 0, 0)
'        Else
'
'            Dim dist As Single
'
'            dist = 1 - (1 * ((pPlayer.Origin.Y - pPlanet.LevelLow) / (pPlanet.LevelTop - pPlanet.LevelLow)))
'
'            Dim clr1 As Long
'            Dim clr2 As Long
'            Dim clr3 As Long
'
'            If dist > 1 Then dist = 1
'            clr1 = RGB(0, 0, 0)
'            clr2 = RGB(pPlanet.Color.Red, pPlanet.Color.Green, pPlanet.Color.Blue)
'
'            clr3 = Blend(clr1, clr2, dist)
'
'            Color = D3DColorARGB(1, rgbRed(clr3), rgbGreen(clr3), rgbBlue(clr3)) ', pPlanet.Color.Green, pPlanet.Color.Blue)
'        End If
        
        
'        Dim p As Planet
'        Dim p1 As Planet
'        Dim p2 As Planet
'        Dim p3 As Planet
'
'        Dim dist1 As Single
'        Dim dist2 As Single
'        Dim dist3 As Single
'
'        For Each p In Planets
'            dist1 = Distance(Player.Origin.X, Player.Origin.Y, Player.Origin.Z, p.Origin.X, p.Origin.Y, p.Origin.Z)
'            Set p1 = p
'            If dist1 = 0 Then
'                dist1 = dist3
'                Set p1 = p3
'            ElseIf dist2 = 0 Then
'                dist2 = dist3
'                Set p2 = p3
'            End If
'
'            If (dist1 < dist2) Then
'                modCommon.Swap dist2, dist1
'                modCommon.Swap p2, p1
'            End If
'            If (dist1 < dist3) Then
'                modCommon.Swap dist1, dist3
'                modCommon.Swap p1, p3
'            End If
'            If (dist2 < dist1) Then
'                modCommon.Swap dist1, dist2
'                modCommon.Swap p1, p2
'            End If
'
'        Next
'        Dim clr3 As Long
'
'        If (p1 Is Nothing) And (p2 Is Nothing) Then
'            clr3 = RGB(0, 0, 0)
'        Else
'            If Not ((p1 Is Nothing) Or (p2 Is Nothing)) Then
'
'                Dim clr1 As Long
'                Dim clr2 As Long
'
'                Dim pt1 As Point
'                Dim pt2 As Point
'
'                Set pt1 = p1.Origin
'                clr1 = p1.Color.RGB
'                Set pt2 = p2.Origin
'                clr2 = p2.Color.RGB
'
'                dist3 = Distance(Player.Origin.X, Player.Origin.Y, Player.Origin.Z, 0, (FAR / 90) + (FAR / 10), 0)
'                If dist3 < dist1 Then
'                    clr1 = RGB(0, 0, 0)
'                    Set pt1 = MakePoint(0, (FAR / 90) + (FAR / 10), 0)
'                ElseIf dist3 < dist2 Then
'                    clr2 = RGB(0, 0, 0)
'                    Set pt2 = MakePoint(0, (FAR / 90) + (FAR / 10), 0)
'                End If
'
'
'                dist3 = DistanceEx(pt2, pt1)
'
'
'                If dist3 > 0 Then
'                    dist1 = Distance(pt1.X, pt1.Y, pt1.Z, Player.Origin.X, Player.Origin.Y, Player.Origin.Z)
'                    dist2 = Distance(pt2.X, pt2.Y, pt2.Z, Player.Origin.X, Player.Origin.Y, Player.Origin.Z)
'
'                    clr3 = Blend(clr1, clr2, (dist3 / (dist1 + dist2)) * (dist1 / dist3))
'
'                Else
'                    clr3 = clr1
'                End If
'
'            ElseIf Not (p2 Is Nothing) Then
'                clr3 = p2.Color.RGB
'               ' Debug.Print Distance(p2.Origin.X, p2.Origin.Y, p2.Origin.Z, Player.Origin.X, Player.Origin.Y, Player.Origin.Z)
'            End If
'        End If
'        Color = D3DColorARGB(1, rgbRed(clr3), rgbGreen(clr3), rgbBlue(clr3))
    End If
End Property


Private Sub Class_Initialize()
    pColor = D3DColorARGB(0, 0, 0, 0)
End Sub

