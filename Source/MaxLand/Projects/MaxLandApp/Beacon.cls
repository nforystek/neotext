VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Beacon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pFileNames As VBA.Collection
Private pBeaconSkin() As Direct3DTexture8

Private pBeaconPlaq(0 To 5) As MyVertex
Private pBeaconVBuf As Direct3DVertexBuffer8

Private pOrigins As VBA.Collection

Private pDimensions As ImageDimensions
Private pPercentXY As Coord

Private pVisible As Boolean

Private pHorizontalLock As Boolean
Private pVerticalLock As Boolean
Private pVerticalSkew As Boolean
Private pRoundingCut As Integer

Private pTranslucent As Boolean
Private pAlphablend As Boolean

Private pConsumable As Boolean
Private pRandomize As Boolean
Private pAllowance As Long

Private pBeaconAnim As Double
Private pBeaconText As Long
Private pBeaconLight As Long

Private pKey As String
Private pSerialize As Boolean

Public Property Get Origins() As VBA.Collection
    Set Origins = pOrigins
End Property
Public Property Get Origin(Optional ByVal Index As Integer = 0) As Point
    If pOrigins.Count > 0 And Index <= pOrigins.Count And Index > 0 Then
        Set Origin = pOrigins(Index)
    ElseIf pOrigins.Count > 0 Then
        Set Origin = pOrigins(1)
    End If
End Property
Public Property Set Origin(Optional ByVal Index As Integer = 0, ByRef RHS As Point)
    If pOrigins.Count > 0 And Index <= pOrigins.Count And Index > 0 Then
        pOrigins.Add RHS, , , Index
        pOrigins.Remove Index
    Else
        pOrigins.Add RHS
    End If
End Property
Public Property Let Origin(Optional ByVal Index As Integer = 0, ByVal RHS)
    Dim o As New Point
    o = RHS
    If pOrigins.Count > 0 And Index <= pOrigins.Count And Index > 0 Then
        pOrigins.Add o, , , Index
        pOrigins.Remove Index
    Else
        pOrigins.Add o
    End If
End Property

Public Property Get FileNames() As VBA.Collection
    Set FileNames = pFileNames
End Property
Public Property Get FileName(Optional ByVal Index As Long = 0) As String
    If pFileNames.Count > 0 And Index <= pFileNames.Count And Index > 0 Then
        FileName = pFileNames.Item(Index)
    ElseIf pFileNames.Count > 0 Then
        FileName = pFileNames.Item(1)
    End If
End Property
Public Property Let FileName(Optional ByVal Index As Long = 0, ByVal RHS As String)
    If Index = 0 Then
        If PathExists(ScriptRoot & "\Models\" & RHS, True) Then
            pFileNames.Add RHS
            ReDim Preserve pBeaconSkin(1 To pFileNames.Count) As Direct3DTexture8
            Set pBeaconSkin(pFileNames.Count) = LoadTexture(ScriptRoot & "\Models\" & RHS)
            Recreate
        Else
            Err.Raise 8, App.EXEName, "File not found."
        End If
    ElseIf pFileNames.Count > 0 And Index <= pFileNames.Count And Index > 0 Then
        If PathExists(ScriptRoot & "\Models\" & RHS, True) Then
            Set pBeaconSkin(Index) = Nothing
            Set pBeaconSkin(Index) = LoadTexture(ScriptRoot & "\Models\" & RHS)
            pFileNames.Add RHS, , , Index
            pFileNames.Remove Index
            Recreate
        Else
            Err.Raise 8, App.EXEName, "File not found."
        End If
    End If
End Property

Public Property Get PercentXY() As Coord
    Set PercentXY = pPercentXY
End Property

Public Property Set PercentXY(ByRef RHS As Coord)
    Set pPercentXY = RHS
    Recreate
End Property

Public Property Get PercentX() As Single
    PercentX = pPercentXY.X
End Property
Public Property Let PercentX(ByVal RHS As Single)
    pPercentXY.X = RHS
    Recreate
End Property

Public Property Get PercentY() As Single
    PercentY = pPercentXY.Y
End Property
Public Property Let PercentY(ByVal RHS As Single)
    pPercentXY.Y = RHS
    Recreate
End Property

Public Property Get Dimensions() As Coord
    Set Dimensions = New Coord
    Dimensions.X = pDimensions.Width
    Dimensions.Y = pDimensions.Height
End Property

Public Property Set Dimensions(ByRef RHS As Coord)
    pDimensions.Width = RHS.X
    pDimensions.Height = RHS.Y
    Recreate
End Property

Public Property Get Height() As Single
    Height = pDimensions.Height
End Property
Public Property Let Height(ByVal RHS As Single)
    pDimensions.Height = RHS
    Recreate
End Property

Public Property Get Width() As Single
    Width = pDimensions.Width
End Property
Public Property Let Width(ByVal RHS As Single)
    pDimensions.Width = RHS
    Recreate
End Property




Public Property Get Visible() As Boolean
    Visible = pVisible
End Property
Public Property Let Visible(ByVal RHS As Boolean)
    pVisible = RHS
End Property

Public Property Get HorizontalLock() As Boolean
    HorizontalLock = pHorizontalLock
End Property
Public Property Let HorizontalLock(ByVal RHS As Boolean)
    pHorizontalLock = RHS
End Property

Public Property Get VerticalLock() As Boolean
    VerticalLock = pVerticalLock
End Property
Public Property Let VerticalLock(ByVal RHS As Boolean)
    pVerticalLock = RHS
End Property

Public Property Get VerticalSkew() As Boolean
    VerticalSkew = pVerticalSkew
End Property
Public Property Let VerticalSkew(ByVal RHS As Boolean)
    pVerticalSkew = RHS
End Property

Public Property Get RoundingCut() As Integer
    RoundingCut = pRoundingCut
End Property
Public Property Let RoundingCut(ByVal RHS As Integer)
    pRoundingCut = RHS
End Property

Friend Property Get BeaconAnim() As Double
    BeaconAnim = pBeaconAnim
End Property
Friend Property Let BeaconAnim(ByVal RHS As Double)
    pBeaconAnim = RHS
End Property

Friend Property Get BeaconText() As Long
    BeaconText = pBeaconText
End Property
Friend Property Let BeaconText(ByVal RHS As Long)
    pBeaconText = RHS
End Property

Public Property Get BeaconLight() As Long
    BeaconLight = pBeaconLight
End Property
Public Property Let BeaconLight(ByVal RHS As Long)
    pBeaconLight = RHS
End Property


Public Property Get Randomize() As Boolean
    Randomize = pRandomize
End Property
Public Property Let Randomize(ByVal RHS As Boolean)
    pRandomize = RHS
End Property

Public Property Get Consumable() As Boolean
    Consumable = pConsumable
End Property
Public Property Let Consumable(ByVal RHS As Boolean)
    pConsumable = RHS
End Property

Public Property Get Allowance() As Long
    Allowance = pAllowance
End Property
Public Property Let Allowance(ByVal RHS As Long)
    pAllowance = RHS
End Property

Public Property Get Translucent() As Boolean
    Translucent = pTranslucent
End Property
Public Property Let Translucent(ByVal RHS As Boolean)
    pTranslucent = RHS
    If pTranslucent Then pAlphablend = False
End Property

Public Property Get Alphablend() As Boolean
    Alphablend = pAlphablend
End Property
Public Property Let Alphablend(ByVal RHS As Boolean)
    pAlphablend = RHS
    If pAlphablend Then pTranslucent = False
End Property

Public Property Get Serialize() As Boolean
    Serialize = True
End Property
Public Property Let Serialize(ByVal RHS As Boolean)
    pSerialize = RHS
End Property

Public Property Get Key() As String
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

Private Sub Class_Initialize()
        
    Stats_Beacon_Count = Stats_Beacon_Count + 1
    Set pFileNames = New VBA.Collection
    Set pOrigins = New VBA.Collection
    Set pPercentXY = New Coord
    
    Set pBeaconVBuf = DDevice.CreateVertexBuffer(Len(pBeaconPlaq(0)) * 6, 0, FVF_RENDER, D3DPOOL_DEFAULT)
    BeaconText = 1
    BeaconLight = -1
    Width = 1
    Height = 1
    PercentX = 100
    PercentY = 100
    Allowance = 1

    Key = "K" & ObjPtr(Me)
End Sub

Private Sub Class_Terminate()
    If Not pOrigins Is Nothing Then
        Do Until pOrigins.Count = 0
            pOrigins.Remove 1
        Loop
        Set pOrigins = Nothing
    End If
    If Not pFileNames Is Nothing Then
        Do Until pFileNames.Count = 0
            pFileNames.Remove 1
        Loop
        Set pFileNames = Nothing
    End If
    Set pBeaconVBuf = Nothing
    Set pPercentXY = Nothing
    Stats_Beacon_Count = Stats_Beacon_Count - 1
End Sub

Public Sub SetTexture(ByVal Stage As Long, ByVal Index As Long)
Attribute SetTexture.VB_MemberFlags = "40"
    DDevice.SetTexture Stage, pBeaconSkin(Index)
End Sub
Public Sub Render()
Attribute Render.VB_MemberFlags = "40"
    DDevice.SetStreamSource 0, pBeaconVBuf, Len(pBeaconPlaq(0))
    DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
End Sub
Private Sub Recreate()
    CreateSquare pBeaconPlaq, 0, _
        MakeVector(((1 * (PercentXY.X / 100)) / 2), 0, 0), _
        MakeVector(-((1 * (PercentXY.X / 100)) / 2), 0, 0), _
        MakeVector(-((1 * (PercentXY.X / 100)) / 2), (1 * (PercentXY.Y / 100)), 0), _
        MakeVector(((1 * (PercentXY.X / 100)) / 2), (1 * (PercentXY.Y / 100)), 0)
    D3DVertexBuffer8SetData pBeaconVBuf, 0, Len(pBeaconPlaq(0)) * 6, 0, pBeaconPlaq(0)
End Sub
