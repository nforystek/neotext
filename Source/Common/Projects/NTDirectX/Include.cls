VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Include"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Location(ByRef Origin As Point, Optional ByRef ApplyTo As Object)
    mod3DObj.Location Origin, ApplyTo   'location is changing the origin to absolute
End Sub
Public Sub Position(ByRef Origin As Point, Optional ByRef ApplyTo As Object)
    mod3DObj.Position Origin, ApplyTo  'position is changing the origin relative
End Sub

Public Sub Rotation(ByRef Radians As Point, Optional ByRef ApplyTo As Object)
    mod3DObj.Rotation Radians, ApplyTo  'location is changing the origin to absolute
End Sub
Public Sub Orientate(ByRef Radians As Point, Optional ByRef ApplyTo As Object)
    mod3DObj.Orientate Radians, ApplyTo  'position is changing the origin relative
End Sub

Public Sub Scaling(ByRef Ratios As Point, Optional ByRef ApplyTo As Object)
    mod3DObj.Scaling Ratios, ApplyTo  'location is changing the origin to absolute
End Sub
Public Sub Explode(ByRef Ratios As Point, Optional ByRef ApplyTo As Object)
    mod3DObj.Explode Ratios, ApplyTo  'position is changing the origin relative
End Sub

Public Property Set Serialize(ByRef RHS)

    Dim code As String
    Dim line As String
    Dim o As Object 'NTNodes10.Collection
    Select Case LCase(TypeName(RHS))
        Case "bindings", "camera"
            If All.Exists(LCase(TypeName(RHS))) Then
                code = RemoveQuotedArg(All(LCase(TypeName(RHS))), "<" & TypeName(RHS) & ">", "</" & TypeName(RHS) & ">")
                All.Remove LCase(TypeName(RHS))
                Do Until code = ""
                    line = URLDecode(RemoveNextArg(code, vbCrLf))
                    frmMain.ExecuteStatement line
                Loop
            End If
        Case "molecules", "brilliants", "planets", "billboards"
            If All.Exists(LCase(TypeName(RHS))) Then
                Set o = NTNodes10.Deserialize(All(LCase(TypeName(RHS))), AppPath & App.EXEName & ".ocx")
                All.Remove LCase(TypeName(RHS))
                If o.Count > 0 Then
                    Dim m As Object
                    For Each m In o
                        m.Serialize = True
                        If RHS.Exists(m.Key) Then RHS.Remove m.Key
                        RHS.Add m, m.Key
                        If All.Exists(m.Key) Then All.Remove m.Key
                        All.Add m, m.Key
                        ParseSetupObject m, LCase(TypeName(RHS))
                    Next
                End If
                Set o = Nothing
            End If
    End Select
End Property

Public Property Get Serialize()
Attribute Serialize.VB_MemberFlags = "40"

    Serialize = "<DateTime>" & Include.URLEncode(FileDateTime(ScriptRoot & "\Index.vbx")) & "</DateTime>" & vbCrLf
    Serialize = Serialize & Replace(Replace(NTNodes10.Serialize(Billboards, AppPath & App.EXEName & ".ocx"), "<?xml version=""1.0""?>" & vbCrLf & "<Serial>" & vbCrLf, ""), "</Serial>" & vbCrLf, "")
    Serialize = Serialize & Replace(Replace(NTNodes10.Serialize(Planets, AppPath & App.EXEName & ".ocx"), "<?xml version=""1.0""?>" & vbCrLf & "<Serial>" & vbCrLf, ""), "</Serial>" & vbCrLf, "")
    Serialize = Serialize & Replace(Replace(NTNodes10.Serialize(Brilliants, AppPath & App.EXEName & ".ocx"), "<?xml version=""1.0""?>" & vbCrLf & "<Serial>" & vbCrLf, ""), "</Serial>" & vbCrLf, "")
    Serialize = Serialize & Replace(Replace(NTNodes10.Serialize(Molecules, AppPath & App.EXEName & ".ocx"), "<?xml version=""1.0""?>" & vbCrLf & "<Serial>" & vbCrLf, ""), "</Serial>" & vbCrLf, "")

End Property

Public Function Point(ByVal X As Single, ByVal Y As Single, ByVal z As Single) As Point
    Set Point = New Point
    Point.X = X
    Point.Y = Y
    Point.z = z
End Function

Public Function URLDecode(ByVal encodedString As String) As String

    Dim ReturnString As String
    Dim currentChar As String
    
    Dim i As Long
    i = 1

    Do Until i > Len(encodedString)
        currentChar = Mid(encodedString, i, 1)

        If currentChar = "+" Then
            ReturnString = ReturnString + " "
            i = i + 1
        ElseIf currentChar = "%" Then
            currentChar = Mid(encodedString, i + 1, 2)
            ReturnString = ReturnString + Chr(Val("&H" & currentChar))
            i = i + 3
        Else
            ReturnString = ReturnString + currentChar
            i = i + 1
        End If
    Loop
    
    URLDecode = ReturnString

End Function

Public Function URLEncode(ByVal encodeString As String) As String
    Dim ReturnString As String
    Dim currentChar As String
    
    Dim i As Long

    For i = 1 To Len(encodeString)
        currentChar = Mid(encodeString, i, 1)

        If Asc(currentChar) < 91 And Asc(currentChar) > 64 Then
            ReturnString = ReturnString + currentChar
        ElseIf Asc(currentChar) < 123 And Asc(currentChar) > 96 Then
            ReturnString = ReturnString + currentChar
        ElseIf Asc(currentChar) < 58 And Asc(currentChar) > 47 Then
            ReturnString = ReturnString + currentChar
        ElseIf Asc(currentChar) = 32 Then
            ReturnString = ReturnString + "+"
        Else
            If Len(Hex(Asc(currentChar))) = 1 Then
                ReturnString = ReturnString + "%0" + Hex(Asc(currentChar))
            Else
                ReturnString = ReturnString + "%" + Hex(Asc(currentChar))
            End If
        End If
    Next

    URLEncode = ReturnString

End Function

Public Function SafeKey(ByVal mKey As Variant) As Variant
    If IsNumeric(mKey) Then
        SafeKey = CLng(mKey)
    Else
        Dim Ret As String
        Ret = Trim(mKey)
        Ret = Replace(Ret, "/", CStr(Asc("/")))
        Ret = Replace(Ret, "\", CStr(Asc("\")))
        Ret = Replace(Ret, "|", CStr(Asc("|")))
        Ret = Replace(Ret, "[", CStr(Asc("[")))
        Ret = Replace(Ret, "]", CStr(Asc("]")))
        Ret = Replace(Ret, " ", CStr(Asc(" ")))
        Ret = Replace(Ret, "&", CStr(Asc("&")))
        Ret = Replace(Ret, "%", CStr(Asc("%")))
        Ret = Replace(Ret, "$", CStr(Asc("$")))
        Ret = Replace(Ret, ",", CStr(Asc(",")))
        Ret = Replace(Ret, ".", CStr(Asc(".")))
        Ret = Replace(Ret, "-", CStr(Asc("-")))
        Ret = Replace(Ret, "_", CStr(Asc("_")))
        Ret = Replace(Ret, "+", CStr(Asc("+")))
        Ret = Replace(Ret, "=", CStr(Asc("=")))
        Ret = Replace(Ret, "!", CStr(Asc("!")))
        Ret = Replace(Ret, "@", CStr(Asc("@")))
        Ret = Replace(Ret, "#", CStr(Asc("#")))
        Ret = Replace(Ret, "^", CStr(Asc("^")))
        Ret = Replace(Ret, "*", CStr(Asc("*")))
        Ret = Replace(Ret, "(", CStr(Asc(")")))
        Ret = Replace(Ret, "`", CStr(Asc("`")))
        Ret = Replace(Ret, "~", CStr(Asc("~")))
        Ret = Replace(Ret, "<", CStr(Asc("<")))
        Ret = Replace(Ret, ">", CStr(Asc(">")))
        Ret = Replace(Ret, "'", CStr(Asc("'")))
        Ret = Replace(Ret, """", CStr(Asc("""")))
        SafeKey = LCase(Ret)
    End If
End Function

'###############################################################
'### Code base of Sub/Function's added to ScriptControl are  ###
'### also exposed in modScript private to modFactory.Execute ###
'###############################################################
    
Public Property Get NotLoaded() As BillboardTypes
    NotLoaded = BillboardTypes.NotLoaded
End Property
Public Property Get CacheOnly() As BillboardTypes
    CacheOnly = BillboardTypes.CacheOnly
End Property
Public Property Get HudPanel() As BillboardTypes
    HudPanel = BillboardTypes.HudPanel
End Property
Public Property Get Beacon() As BillboardTypes
    Beacon = BillboardTypes.Beacon
End Property

Public Property Get Directed() As BrilliantTypes
    Directed = BrilliantTypes.Directed
End Property
Public Property Get Spot() As BrilliantTypes
    Spot = BrilliantTypes.Spot
End Property
Public Property Get Omni() As BrilliantTypes
    Omni = BrilliantTypes.Omni
End Property

Public Property Get Direct() As MotionTypes
    Direct = MotionTypes.Direct
End Property
Public Property Get Rotate() As MotionTypes
    Rotate = MotionTypes.Rotate
End Property
Public Property Get Scalar() As MotionTypes
    Scalar = MotionTypes.Scalar
End Property
Public Property Get Statue() As MotionTypes
    Statue = MotionTypes.Statue
End Property

Public Property Get Shade() As PlanetTypes
    Shade = PlanetTypes.Shade
End Property
Public Property Get World() As PlanetTypes
    World = PlanetTypes.World
End Property
Public Property Get Plateau() As PlanetTypes
    Plateau = PlanetTypes.Plateau
End Property
Public Property Get Screen() As PlanetTypes
    Screen = PlanetTypes.Screen
End Property


Public Property Get None() As CollisionTypes
    None = CollisionTypes.None
End Property
Public Property Get Ranged() As CollisionTypes
    Ranged = CollisionTypes.Ranged
End Property
Public Property Get Through() As CollisionTypes
    Through = CollisionTypes.Through
End Property
Public Property Get Freely() As CollisionTypes
    Freely = CollisionTypes.Freely
End Property
Public Property Get Gravity() As CollisionTypes
    Gravity = CollisionTypes.Gravity
End Property
Public Property Get Liquid() As CollisionTypes
    Liquid = CollisionTypes.Liquid
End Property
Public Property Get Curbing() As CollisionTypes
    Curbing = CollisionTypes.Curbing
End Property
Public Property Get Coupling() As CollisionTypes
    Coupling = CollisionTypes.Coupling
End Property

Public Property Get Visual() As ControllerModes
    Visual = ControllerModes.Visual
End Property
Public Property Get Hidden() As ControllerModes
    Hidden = ControllerModes.Hidden
End Property
Public Property Get Trapping() As ControllerModes
    Trapping = ControllerModes.Trapping
End Property

Public Property Get Unspecified() As CoordinateTypes
    Unspecified = CoordinateTypes.Unspecified
End Property
Public Property Get Relative() As CoordinateTypes
    Relative = CoordinateTypes.Relative
End Property
Public Property Get Absolute() As CoordinateTypes
    Absolute = CoordinateTypes.Absolute
End Property

'#####################################################
'########## Collections ##############################
'#####################################################
Public Property Get All() As NTNodes10.Collection
    Set All = modParse.All
End Property

Public Property Get Camera() As Camera
    Set Camera = modParse.Camera
End Property

Public Property Get Motions() As Motions
    Set Motions = modParse.Motions
End Property

Public Property Get Brilliants() As Brilliants
    Set Brilliants = modParse.Brilliants
End Property

Public Property Get Molecules() As Molecules
    Set Molecules = modParse.Molecules
End Property

Public Property Get Billboards() As Billboards
    Set Billboards = modParse.Billboards
End Property

Public Property Get Bindings() As Bindings
    Set Bindings = modParse.Bindings
End Property

Public Property Get Planets() As Planets
    Set Planets = modParse.Planets
End Property

'###################################################
'########## Functions ##############################
'###################################################

Public Function CreateObjectPrivate(ByVal Class As String)
    Set CreateObjectPrivate = modFactory.CreateObjectPrivate(Class)
End Function

Public Property Get MILE() As Single
    MILE = modGeometry.MILE
End Property
Public Property Get PI() As Single
    PI = modGeometry.PI
End Property

Public Sub DebugPrint(ByVal Text As String)
    Debug.Print Text
End Sub

Public Sub Echo(ByVal Text As String)
    Process "echo " & frmMain.Evaluate(Text)
End Sub
