VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Screen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pVisible As Boolean
Private pKey As String
Private pSerialize As Boolean

Private pFileName As String
Private pImage As Direct3DTexture8
Private pVerticies(0 To 4) As MyScreen

Private pTranslucent As Boolean
Private pAlphablend As Boolean
Private pCoordinate As Plot
Private pDimensions As ImageDimensions

Private pPadding As Long

Friend Sub Render()
    DDevice.SetTexture 0, pImage
    DDevice.SetTexture 1, pImage
    DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, pVerticies(0), LenB(pVerticies(0))
End Sub

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

Public Property Get Visible() As Boolean
    Visible = pVisible
End Property
Public Property Let Visible(ByVal RHS As Boolean)
    pVisible = RHS
End Property

Public Property Get Alphablend() As Boolean
    Alphablend = pAlphablend
End Property
Public Property Let Alphablend(ByVal RHS As Boolean)
    pAlphablend = RHS
End Property

Public Property Get Translucent() As Boolean
    Translucent = pTranslucent
End Property
Public Property Let Translucent(ByVal RHS As Boolean)
    pTranslucent = RHS
End Property


Public Property Get Height() As Single
    Height = pDimensions.Height
End Property
Public Property Let Height(ByVal RHS As Single)
    pDimensions.Height = RHS
    ResetVerticies
End Property

Public Property Get Width() As Single
    Width = pDimensions.Width
End Property
Public Property Let Width(ByVal RHS As Single)
    pDimensions.Width = RHS
    ResetVerticies
End Property

Public Property Get Top() As Single
    Top = pCoordinate.Y
End Property
Public Property Let Top(ByVal RHS As Single)
    pCoordinate.Y = RHS
    ResetVerticies
End Property

Public Property Get Left() As Single
    Left = pCoordinate.X
End Property
Public Property Let Left(ByVal RHS As Single)
    pCoordinate.X = RHS
    ResetVerticies
End Property
'                                                        .Verticies(0).X = .Padding
'                                                        .Verticies(2).X = .Padding
'                                                .Verticies(1).X = .Verticies(0).X + .Dimensions.Width
'                                                .Verticies(3).X = .Verticies(2).X + .Dimensions.Width
Public Property Get Padding() As Long
    Padding = pPadding
End Property
Public Property Let Padding(ByVal RHS As Long)
    pPadding = RHS
    ResetVerticies
End Property

Public Property Get FileName() As String
    FileName = pFileName
End Property
Public Property Let FileName(ByVal RHS As String)
    If PathExists(ScriptRoot & "Models\" & RHS, True) Then
        pFileName = RHS
        Set pImage = LoadTextureEx(ScriptRoot & "Models\" & RHS, pDimensions)
        If pImage Is Nothing Then
            Debug.Print "Image Dimensions Error"
        End If
        ResetVerticies
    Else
        Debug.Print "Image Not Found"
    End If
End Property

Public Property Get Coordinate() As Plot
    Set Coordinate = pCoordinate
End Property
Public Property Set Coordinate(ByRef RHS As Plot)
    Set pCoordinate = RHS
    ResetVerticies
End Property

Public Property Get Dimensions() As Plot
    Set Dimensions = New Plot
    Dimensions.X = pDimensions.Width
    Dimensions.Y = pDimensions.Height
End Property

Public Property Set Dimensions(ByRef RHS As Plot)
    pDimensions.Width = RHS.X
    pDimensions.Height = RHS.Y
    ResetVerticies
End Property

Private Sub ResetVerticies()
    pVerticies(0).X = pPadding + CLng(pCoordinate.X)
    pVerticies(2).X = pPadding + CLng(pCoordinate.X)
    pVerticies(1).X = pVerticies(0).X + pDimensions.Width
    pVerticies(3).X = pVerticies(2).X + pDimensions.Width
    pVerticies(0).Y = pPadding + CLng(pCoordinate.Y)
    pVerticies(1).Y = pPadding + CLng(pCoordinate.Y)
    pVerticies(2).Y = pVerticies(0).Y + pDimensions.Height
    pVerticies(3).Y = pVerticies(1).Y + pDimensions.Height
End Sub
Public Sub AlignLeft()
    pVerticies(0).X = pPadding
    pVerticies(2).X = pPadding
    pVerticies(1).X = pVerticies(0).X + pDimensions.Width
    pVerticies(3).X = pVerticies(2).X + pDimensions.Width
End Sub
Public Sub AlignCenter()
    pVerticies(0).X = ((frmMain.Width / Screen.TwipsPerPixelX) / 2) - (pDimensions.Width / 2)
    pVerticies(2).X = ((frmMain.Width / Screen.TwipsPerPixelX) / 2) - (pDimensions.Width / 2)
    pVerticies(1).X = pVerticies(0).X + pDimensions.Width
    pVerticies(3).X = pVerticies(2).X + pDimensions.Width
End Sub
Public Sub ALignRight()
    pVerticies(0).X = (frmMain.Width / Screen.TwipsPerPixelX) - pPadding - pDimensions.Width
    pVerticies(2).X = (frmMain.Width / Screen.TwipsPerPixelX) - pPadding - pDimensions.Width
    pVerticies(1).X = pVerticies(0).X + pDimensions.Width
    pVerticies(3).X = pVerticies(2).X + pDimensions.Width
End Sub
Public Sub AlignTop()
    pVerticies(0).Y = pPadding
    pVerticies(1).Y = pPadding
    pVerticies(2).Y = pVerticies(0).Y + pDimensions.Height
    pVerticies(3).Y = pVerticies(1).Y + pDimensions.Height
End Sub
Public Sub AlignMiddle()
    pVerticies(0).Y = ((frmMain.Height / Screen.TwipsPerPixelY) / 2) - (pDimensions.Height / 2)
    pVerticies(1).Y = ((frmMain.Height / Screen.TwipsPerPixelY) / 2) - (pDimensions.Height / 2)
    pVerticies(2).Y = pVerticies(0).Y + pDimensions.Height
    pVerticies(3).Y = pVerticies(1).Y + pDimensions.Height
End Sub
Public Sub AlignBottom()
    pVerticies(0).Y = (frmMain.Height / Screen.TwipsPerPixelY) - pPadding - pDimensions.Height
    pVerticies(1).Y = (frmMain.Height / Screen.TwipsPerPixelY) - pPadding - pDimensions.Height
    pVerticies(2).Y = pVerticies(0).Y + pDimensions.Height
    pVerticies(3).Y = pVerticies(1).Y + pDimensions.Height
End Sub


Private Sub Class_Initialize()
    Stats_Image_Count = Stats_Image_Count + 1
    Set pCoordinate = New Plot
    pVerticies(0) = MakeScreen(0, 0, -1, 0, 0)
    pVerticies(1) = MakeScreen(0, 0, -1, 1, 0)
    pVerticies(2) = MakeScreen(0, 0, -1, 0, 1)
    pVerticies(3) = MakeScreen(0, 0, -1, 1, 1)
    
    Key = "K" & ObjPtr(Me)
End Sub

Private Sub Class_Terminate()
    Set pImage = Nothing
    Set pCoordinate = Nothing
    Stats_Image_Count = Stats_Image_Count - 1
End Sub
