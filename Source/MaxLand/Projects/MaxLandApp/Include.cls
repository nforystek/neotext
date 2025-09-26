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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'Public Const PI As Single = 3.14159265359
''Public Const PI As Double = 3.14159265358979            '653589 '79 'leaving out 65358979
'                    'because the planes mess up and wobble
'
'Public Const Epsilon As Double = 0.999999999999999 ' 0.0001 f'
'Public Const DEGREE As Single = (180 / PI)
'Public Const RADIAN As Single = (PI / 180)
'Public Const FOOT As Single = 0.1
'Public Const MILE As Single = 5280 * FOOT
''Public Const FOVY As Single = (FOOT * 8) '(FOOT * 8) '4 feet left, and 4 feet right = 0.8
'Public Const SKYFOVY As Single = (MILE * 4)
'Public Const Far  As Single = 900000000
'Public Const Near As Single = 0 '0.05 'one millimeter (308.4 per foor) or greater
'
'
'
'Public Const FOVY As Single = 1.047198 '2.3561946
'
'Public Const D90 As Single = (PI / 4)
'Public Const D180 As Single = (PI / 2)
'Public Const D360 As Single = PI
'Public Const D720 As Single = (PI * 2)

Public Property Get MouseTrapped() As Boolean
    MouseTrapped = modMain.MouseTrapped
End Property

Public Property Get FullScreen() As Boolean
    FullScreen = modMain.FullScreen
End Property
Public Property Get ConsoleVisible() As Boolean
    ConsoleVisible = modCmds.ConsoleVisible
End Property
Public Sub ConsoleToggle()
    modCmds.ConsoleToggle
End Sub

Public Property Get PI() As Single
    PI = modGeometry.PI
End Property
Public Property Get Epsilon() As Single
    Epsilon = modGeometry.Epsilon
End Property
Public Property Get DEGREE() As Single
    DEGREE = modGeometry.DEGREE
End Property
Public Property Get RADIAN() As Single
    RADIAN = modGeometry.RADIAN
End Property
Public Property Get FOOT() As Single
    FOOT = modGeometry.FOOT
End Property
Public Property Get MILE() As Single
    MILE = modGeometry.MILE
End Property
Public Property Get SKYFOV() As Single
    SKYFOV = modGeometry.SKYFOVY
End Property
Public Property Get NEAR() As Single
    NEAR = modGeometry.NEAR
End Property
Public Property Get FAR() As Single
    FAR = modGeometry.FAR
End Property
Public Property Get FOVY() As Single
    FOVY = modGeometry.FOVY
End Property

'Public Enum Collides
'    InDoor = 1
'    Ground = 2
'    Liquid = 3
'    Ladder = 4
'End Enum

Public Function GetBindingText(ByVal BindIndex As Integer) As String
    GetBindingText = modCmds.GetBindingText(BindIndex)
End Function
Public Function GetBindingIndex(ByVal BindText As String) As Integer
    GetBindingIndex = modCmds.GetBindingIndex(BindText)
End Function

Public Function MakePoint(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As Point
    Set MakePoint = New Point
    MakePoint.X = X
    MakePoint.Y = Y
    MakePoint.Z = Z
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
        Dim ret As String
        ret = Trim(mKey)
        ret = Replace(ret, "/", CStr(Asc("/")))
        ret = Replace(ret, "\", CStr(Asc("\")))
        ret = Replace(ret, "|", CStr(Asc("|")))
        ret = Replace(ret, "[", CStr(Asc("[")))
        ret = Replace(ret, "]", CStr(Asc("]")))
        ret = Replace(ret, " ", CStr(Asc(" ")))
        ret = Replace(ret, "&", CStr(Asc("&")))
        ret = Replace(ret, "%", CStr(Asc("%")))
        ret = Replace(ret, "$", CStr(Asc("$")))
        ret = Replace(ret, ",", CStr(Asc(",")))
        ret = Replace(ret, ".", CStr(Asc(".")))
        ret = Replace(ret, "-", CStr(Asc("-")))
        ret = Replace(ret, "_", CStr(Asc("_")))
        ret = Replace(ret, "+", CStr(Asc("+")))
        ret = Replace(ret, "=", CStr(Asc("=")))
        ret = Replace(ret, "!", CStr(Asc("!")))
        ret = Replace(ret, "@", CStr(Asc("@")))
        ret = Replace(ret, "#", CStr(Asc("#")))
        ret = Replace(ret, "^", CStr(Asc("^")))
        ret = Replace(ret, "*", CStr(Asc("*")))
        ret = Replace(ret, "(", CStr(Asc(")")))
        ret = Replace(ret, "`", CStr(Asc("`")))
        ret = Replace(ret, "~", CStr(Asc("~")))
        ret = Replace(ret, "<", CStr(Asc("<")))
        ret = Replace(ret, ">", CStr(Asc(">")))
        ret = Replace(ret, "'", CStr(Asc("'")))
        ret = Replace(ret, """", CStr(Asc("""")))
        SafeKey = LCase(ret)
    End If
End Function

'###############################################################
'### Code base of Sub/Function's added to ScriptControl are  ###
'### also exposed in modScript private to modFactory.Execute ###
'###############################################################
    


Public Property Get None()
    None = 0
End Property


Public Property Get Directed() As Lighting
    Directed = Lighting.Directed
End Property
Public Property Get Spot() As Lighting
    Spot = Lighting.Spot
End Property
Public Property Get Omni() As Lighting
    Omni = Lighting.Omni
End Property

'Public Enum Actions
'    NotDefined = 0
'    Directing = 1
'    Rotating = 2
'    Scaling = 4
'    Method = 8
'End Enum

Public Property Get Directing() As Actions
    Directing = Actions.Directing
End Property
Public Property Get Rotating() As Actions
    Rotating = Actions.Rotating
End Property
Public Property Get Scaling() As Actions
    Scaling = Actions.Scaling
End Property
Public Property Get Method() As Actions
    Method = Actions.Method
End Property

'Public Enum Collides
'    Normal = 0
'    InDoor = 1
'    Ground = 2
'    Liquid = 3
'    Ladder = 4
'End Enum
Public Property Get Normal() As Long
    Normal = 0
End Property
Public Property Get InDoor() As Long
    InDoor = 1
End Property
Public Property Get Ground() As Long
    Ground = 2
End Property
Public Property Get Liquid() As Long
    Liquid = 3
End Property
Public Property Get Ladder() As Long
    Ladder = 4
End Property

'Public Enum Moving
'    None = 0
'    Level = 1
'    Flying = 2
'    Falling = 4
'    Stepping = 8
'End Enum

Public Property Get Level() As Moving
    Level = Moving.Level
End Property
Public Property Get Flying() As Moving
    Flying = Moving.Flying
End Property
Public Property Get Falling() As Moving
    Falling = Moving.Falling
End Property
Public Property Get Stepping() As Moving
    Stepping = Moving.Stepping
End Property

'Public Enum Playmode
'    Spectator = 0
'    ThirdPerson = 1
'    FirstPerson = 2
'    CameraMode = 3
'End Enum

Public Property Get Spectator() As Playmode
    Spectator = Playmode.Spectator
End Property
Public Property Get ThirdPerson() As Playmode
    ThirdPerson = Playmode.ThirdPerson
End Property
Public Property Get FirstPerson() As Playmode
    FirstPerson = Playmode.FirstPerson
End Property
Public Property Get CameraMode() As Playmode
    CameraMode = Playmode.CameraMode
End Property


'Public Enum Controller
'    Icon = 0
'    Hide = 1
'    Trap = 2
'End Enum


Public Property Get Icon() As MouseControl
    Icon = MouseControl.Icon
End Property
Public Property Get Hide() As MouseControl
    Hide = MouseControl.Hide
End Property
Public Property Get Trap() As MouseControl
    Trap = MouseControl.Trap
End Property

'Public Enum EventBehavior
'    Rapid = 0
'    Press = 1
'    Locks = 2
'End Enum

Public Property Get Rapid() As EventBehavior
    Rapid = EventBehavior.Rapid
End Property

Public Property Get Press() As EventBehavior
    Press = EventBehavior.Press
End Property

Public Property Get Locks() As EventBehavior
    Locks = EventBehavior.Locks
End Property


'###################################################
'########## Functions ##############################
'###################################################

Public Function NewObject(ByVal Class As String)
    Select Case LCase(Class)
        Case "motions", "boards", "cameras", "elements", "lights", "portals", "screens", "sounds", "tracks"
            Set NewObject = New NTNodes10.Collection
            
        Case "bindings"
            Err.Raise 8, , "Bindings object is not creatable."
        Case "database"
            Err.Raise 8, , "Database object is not creatable."
        Case "include"
            Err.Raise 8, , "Include object is not creatable."
            
        Case "beacon"
            Set NewObject = New Beacon
        Case "board"
            Set NewObject = New Board
        Case "camera"
            Set NewObject = New Camera
        Case "color"
            Set NewObject = New Color
       Case "element"
            Set NewObject = New Element
        Case "light"
            Set NewObject = New Light
        Case "motion"
            Set NewObject = New Motion
        Case "onevent"
            Set NewObject = New OnEvent
        Case "plane"
            Set NewObject = New Plane
        Case "player"
            Set NewObject = New Player
        Case "plot"
            Set NewObject = New Plot
        Case "point"
            Set NewObject = New Point
        Case "portal"
            Set NewObject = New Portal
        Case "screen"
            Set NewObject = New Screen
        Case "sound"
            Set NewObject = New Sound
        Case "space"
            Set NewObject = New Space
        Case "track"
            Set NewObject = New Track
            
        Case Else
            Set NewObject = modFactory.NewObject(Class)
    End Select
End Function

Public Sub DebugPrint(ByVal Text As String)
    Debug.Print Text
End Sub

Public Sub Console(ByVal Text As String)
    ConsoleCommand Text
End Sub


Private Sub Class_Initialize()
    Stats_Include_Count = Stats_Include_Count + 1
End Sub

Private Sub Class_Terminate()
    Stats_Include_Count = Stats_Include_Count - 1
End Sub
