VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.UserControl BrowseButton 
   ClientHeight    =   1425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1770
   HitBehavior     =   2  'Use Paint
   ScaleHeight     =   1425
   ScaleWidth      =   1770
   ToolboxBitmap   =   "BrowseButton.ctx":0000
   Begin VB.CommandButton Command1 
      Height          =   405
      Left            =   1200
      TabIndex        =   0
      Top             =   825
      Visible         =   0   'False
      Width           =   480
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   795
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer buttonOver 
      Enabled         =   0   'False
      Left            =   690
      Top             =   810
   End
   Begin VB.Image imgDisabled 
      Height          =   240
      Left            =   1410
      Picture         =   "BrowseButton.ctx":0312
      Top             =   105
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Dot
      Height          =   270
      Left            =   1440
      Top             =   435
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      X1              =   135
      X2              =   555
      Y1              =   510
      Y2              =   480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      X1              =   660
      X2              =   675
      Y1              =   135
      Y2              =   555
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   135
      X2              =   705
      Y1              =   60
      Y2              =   60
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      X1              =   60
      X2              =   75
      Y1              =   30
      Y2              =   495
   End
   Begin VB.Image imgOut 
      Height          =   240
      Left            =   1095
      Picture         =   "BrowseButton.ctx":069F
      Top             =   60
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgOver 
      Height          =   240
      Left            =   825
      Picture         =   "BrowseButton.ctx":0A2C
      Top             =   75
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image buttonImg 
      Height          =   240
      Left            =   1005
      Picture         =   "BrowseButton.ctx":0DC0
      Top             =   420
      Width           =   240
   End
End
Attribute VB_Name = "BrowseButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'TOP DOWN

Option Compare Binary


Public Enum BorderStyles
    FixedSingle = 1
    MouseOver = 2
    None = 4
End Enum

Public Enum BrowseActions
    Cacneled = -2
    None = -1
    
    Dialog = 0
    Desktop = 1
    
    ProgramFolder = 2
    Network = 14
    
    ControlPanel = 3
    Printers = 4
    
    MyDocuments = 5
    Favorites = 6
    
    RecyclingBin = 10
    MyComputer = 13
End Enum

Private pBorderStyle As BorderStyles

Private pBrowseTitle As String
Private pBrowseAction As BrowseActions
Private pCurrentAction As BrowseActions

Private pBrowseReturn As String

Private pFileFilter As String
Private pFileFilterIndex As Integer

Private pImageOver As Object
Private pImageOut As Object
Private pImageDisabled As Object

Private pToolTipText As String
Private pEnabled As Boolean

Private CancelResize As Boolean
Private IsMouseDown As Boolean

Private Const BorderHighlight = &H80000014
Private Const BorderShadow = &H80000010
Private BorderWidth As Integer

Public Event ButtonClick(ByVal BrowseReturn As String)

Private Type SHITEMID
    cb As Long
    abID() As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" _
                              (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" _
                              (ByVal hwndOwner As Long, ByVal nFolder As Long, _
                              pidl As ITEMIDLIST) As Long

Private Const NOERROR = 0
Private Const CSIDL_DESKTOP = &H0
Private Const CSIDL_PROGRAMS = &H2
Private Const CSIDL_CONTROLS = &H3
Private Const CSIDL_PRINTERS = &H4
Private Const CSIDL_PERSONAL = &H5
Private Const CSIDL_FAVORITES = &H6
Private Const CSIDL_STARTUP = &H7
Private Const CSIDL_RECENT = &H8
Private Const CSIDL_SENDTO = &H9
Private Const CSIDL_BITBUCKET = &HA
Private Const CSIDL_STARTMENU = &HB
Private Const CSIDL_DESKTOPDIRECTORY = &H10
Private Const CSIDL_DRIVES = &H11
Private Const CSIDL_NETWORK = &H12
Private Const CSIDL_NETHOOD = &H13
Private Const CSIDL_FONTS = &H14
Private Const CSIDL_TEMPLATES = &H15
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" _
                              (lpBrowseInfo As BROWSEINFO) As Long

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000

Private Const MAX_PATH = 260
Private Const SHGFI_DISPLAYNAME = &H200

Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400

Private Const ILD_TRANSPARENT = &H1

Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHFILEINFO
   hIcon As Long
   iIcon As Long
   dwAttributes As Long
   szDisplayName As String * MAX_PATH
   szTypeName As String * 80
End Type

Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long

Private MousePos As POINTAPI


Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Public Property Get ImageOver() As Object
    Set ImageOver = pImageOver
End Property
Public Property Set ImageOver(ByRef newVal As Object)
    Set pImageOver = newVal
End Property

Public Property Get ImageOut() As Object
    Set ImageOut = pImageOut
End Property
Public Property Set ImageOut(ByRef newVal As Object)
    Set pImageOut = newVal
End Property

Public Property Get ImageDisabled() As Object
    Set ImageDisabled = pImageDisabled
End Property
Public Property Set ImageDisabled(ByRef newVal As Object)
    Set pImageDisabled = newVal
End Property

Public Property Get TabStop() As Boolean
    TabStop = Command1.TabStop
End Property
Public Property Let TabStop(ByVal newVal As Boolean)
    Command1.TabStop = newVal
End Property
Public Property Get TabIndex() As Integer
    TabIndex = Command1.TabIndex
End Property
Public Property Let TabIndex(ByVal newVal As Integer)
    Command1.TabIndex = newVal
End Property

Public Property Get BorderStyle() As BorderStyles
    BorderStyle = pBorderStyle
End Property
Public Property Let BorderStyle(ByVal newValue As BorderStyles)
    pBorderStyle = newValue
    SetBorder
End Property

Public Property Get BrowseTitle() As String
    BrowseTitle = pBrowseTitle
End Property
Public Property Let BrowseTitle(ByVal newValue As String)
    pBrowseTitle = newValue
End Property
Public Property Get BrowseAction() As BrowseActions
    BrowseAction = pBrowseAction
End Property
Public Property Let BrowseAction(ByVal newValue As BrowseActions)
    pBrowseAction = newValue
End Property

Public Property Get CurrentAction() As BrowseActions
    CurrentAction = pCurrentAction
End Property

Public Property Get BrowseReturn() As String
    BrowseReturn = pBrowseReturn
End Property

Public Property Get FileFilter() As String
    FileFilter = pFileFilter
End Property
Public Property Let FileFilter(ByVal newValue As String)
    pFileFilter = newValue
End Property
Public Property Get FileFilterIndex() As Integer
    FileFilterIndex = pFileFilterIndex
End Property
Public Property Let FileFilterIndex(ByVal newValue As Integer)
    pFileFilterIndex = newValue
End Property
Public Property Get FilterPath() As String
    FilterPath = CommonDialog1.FileName
End Property
Public Property Let FilterPath(ByVal newValue As String)
    CommonDialog1.FileName = newValue
End Property


Public Property Get ToolTipText() As String
    ToolTipText = pToolTipText
End Property
Public Property Let ToolTipText(ByVal newValue As String)
    pToolTipText = newValue
    buttonImg.ToolTipText = newValue
End Property
Public Property Get Enabled() As Boolean
    Enabled = pEnabled
End Property
Public Property Let Enabled(ByVal newValue As Boolean)
    pEnabled = newValue
    UserControl.Enabled = newValue
End Property

Private Sub BrowseClicked()
    If (Not pBrowseAction = BrowseActions.None) And pEnabled Then
        RaiseEvent ButtonClick(Browse())
    Else
        RaiseEvent ButtonClick("")
    End If
End Sub

Private Function IsOverButton() As Boolean
    Dim pt As POINTAPI, hwnd As Long
    
    GetCursorPos pt

    hwnd = WindowFromPoint(pt.X, pt.Y)

    If hwnd = UserControl.hwnd Then
        IsOverButton = True
    Else
        IsOverButton = False
    End If
End Function

Private Sub EventMouseDown()
    buttonOver.Enabled = True
    IsMouseDown = True
    setMouseDown
End Sub

Private Sub EventMouseMove()
    buttonOver.Enabled = True
End Sub

Private Sub EventMouseUp()
    IsMouseDown = False
    
    If IsOverButton Then BrowseClicked
    
    Select Case pBorderStyle
        Case BorderStyles.None
            buttonOver.Enabled = False
            
            setMouseOut
            
        Case BorderStyles.FixedSingle
            buttonOver.Enabled = False
       
            setMouseOver
        
        Case BorderStyles.MouseOver
            buttonOver.Enabled = False
            
            setMouseOut
            
    End Select
End Sub

Private Sub buttonImg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventMouseDown
End Sub

Private Sub buttonImg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventMouseMove
End Sub

Private Sub buttonImg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    EventMouseUp
End Sub

Private Sub Command1_Click()
    If Command1.TabStop And Shape1.Visible Then
        BrowseClicked
    End If
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 32) And Command1.TabStop Then

        Shape1.Top = (Screen.TwipsPerPixelY * 2)
        Shape1.Left = (Screen.TwipsPerPixelX * 2)

        Shape1.Width = UserControl.Width - (Screen.TwipsPerPixelX * 4)
        Shape1.Height = UserControl.Height - (Screen.TwipsPerPixelY * 4)
    End If
End Sub

Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 32) And Command1.TabStop Then

        Shape1.Top = Screen.TwipsPerPixelY
        Shape1.Left = Screen.TwipsPerPixelX

        Shape1.Width = UserControl.Width - (Screen.TwipsPerPixelX * 2)
        Shape1.Height = UserControl.Height - (Screen.TwipsPerPixelY * 2)
    
    End If
End Sub

Private Sub Command1_GotFocus()
    If Command1.TabStop Then
        Shape1.Visible = True
    Else
        Shape1.Visible = False
    End If
End Sub
Private Sub Command1_LostFocus()
    If Command1.TabStop Then
        Shape1.Visible = False
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    buttonImg_MouseDown Button, Shift, X, Y
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    buttonImg_MouseMove Button, Shift, X, Y
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    buttonImg_MouseUp Button, Shift, X, Y
End Sub

Private Sub buttonOver_Timer()

    If IsOverButton Then
        
        If IsMouseDown Then
            setMouseDown
        Else
            setMouseOver
        End If
    Else
        buttonOver.Enabled = False
        If pBorderStyle <> BorderStyles.FixedSingle Then
            setMouseOut
        Else
            setMouseOver
        End If
    End If
    
End Sub

Private Sub setMouseDown()
    If Me.Enabled Then
        buttonImg.Picture = imgOver.Picture
    Else
        buttonImg.Picture = imgDisabled.Picture
    End If
    SetDimensions
    If pBorderStyle = BorderStyles.None Then
        Line1.Visible = False
        Line2.Visible = False
        Line3.Visible = False
        Line4.Visible = False
    Else
        Line1.Visible = True
        Line2.Visible = True
        Line3.Visible = True
        Line4.Visible = True
    End If
    
    Line1.BorderColor = BorderShadow
    Line2.BorderColor = BorderShadow
    Line3.BorderColor = BorderHighlight
    Line4.BorderColor = BorderHighlight
    
    Dim newL As Long
    Dim newT As Long
    newL = BorderWidth + Screen.TwipsPerPixelX
    newT = BorderWidth + Screen.TwipsPerPixelY
    If buttonImg.Left <> newL Then buttonImg.Left = newL
    If buttonImg.Top <> newT Then buttonImg.Top = newT

End Sub
Private Sub setMouseOut()
    If Me.Enabled Then
        Select Case pBorderStyle
            Case BorderStyles.None
                buttonImg.Picture = imgOver.Picture
            Case BorderStyles.FixedSingle
                buttonImg.Picture = imgOver.Picture
            Case BorderStyles.MouseOver
                buttonImg.Picture = imgOut.Picture
        End Select
    Else
        buttonImg.Picture = imgDisabled.Picture
    End If
    SetDimensions
    
    Line1.Visible = False
    Line2.Visible = False
    Line3.Visible = False
    Line4.Visible = False
    
    Line1.BorderColor = BorderHighlight
    Line2.BorderColor = BorderHighlight
    Line3.BorderColor = BorderShadow
    Line4.BorderColor = BorderShadow
    
    Dim newL As Long
    Dim newT As Long
    newL = BorderWidth
    newT = BorderWidth
    If buttonImg.Left <> newL Then buttonImg.Left = newL
    If buttonImg.Top <> newT Then buttonImg.Top = newT

End Sub

Private Sub setMouseOver()
    If Me.Enabled Then
        buttonImg.Picture = imgOver.Picture
    Else
        buttonImg.Picture = imgOver.Picture
    End If
    SetDimensions
    
    If pBorderStyle = BorderStyles.None Then
        Line1.Visible = False
        Line2.Visible = False
        Line3.Visible = False
        Line4.Visible = False
    Else
        Line1.Visible = True
        Line2.Visible = True
        Line3.Visible = True
        Line4.Visible = True
    End If
    
    Line1.BorderColor = BorderHighlight
    Line2.BorderColor = BorderHighlight
    Line3.BorderColor = BorderShadow
    Line4.BorderColor = BorderShadow
    
    Dim newL As Long
    Dim newT As Long
    newL = BorderWidth
    newT = BorderWidth
    If buttonImg.Left <> newL Then buttonImg.Left = newL
    If buttonImg.Top <> newT Then buttonImg.Top = newT

End Sub
Private Sub SetBorder()
    BorderWidth = 2 * Screen.TwipsPerPixelY
    buttonImg.BorderStyle = 0
    
    Select Case pBorderStyle
        Case BorderStyles.None
            buttonOver.Enabled = False
            setMouseOut
            
        Case BorderStyles.FixedSingle
            buttonOver.Enabled = False
            setMouseOver
        
        Case BorderStyles.MouseOver
            buttonOver.Enabled = True
            setMouseOut
        
    End Select

End Sub

Private Function GetFolderValue(tAction As BrowseActions) As Long
    
    Dim wIdx As Integer
    wIdx = CInt(tAction)
    
    If wIdx < 2 Then
        GetFolderValue = 0
    
    ElseIf wIdx < 12 Then
        GetFolderValue = wIdx

    Else
        GetFolderValue = wIdx + 4
    End If

End Function

Private Function GetReturnType() As Long
    Dim dwRtn As Long
    dwRtn = dwRtn Or BIF_RETURNONLYFSDIRS

    GetReturnType = dwRtn
End Function
Public Function GetFolderByAction(ByVal BrowseAction As BrowseActions) As String
    Dim BI As BROWSEINFO
    Dim nFolder As Long
    Dim IDL As ITEMIDLIST
    Dim sPath As String
    With BI
        .hOwner = UserControl.Extender.Parent.hwnd
    
        nFolder = GetFolderValue(pBrowseAction)
    
        If SHGetSpecialFolderLocation(ByVal UserControl.Extender.Parent.hwnd, ByVal nFolder, IDL) = NOERROR Then
            .pidlRoot = IDL.mkid.cb
        End If
    
        .pszDisplayName = String$(MAX_PATH, 0)
    
        .lpszTitle = pBrowseTitle
    
        .ulFlags = GetReturnType()
    
    End With
     
    sPath = String$(MAX_PATH, 0)
    SHGetPathFromIDList ByVal IDL.mkid.cb, ByVal sPath

    sPath = Left(sPath, InStr(sPath, vbNullChar) - 1)

    GetFolderByAction = sPath
End Function

Public Function Browse() As String
    pCurrentAction = pBrowseAction
    
    Dim sPath As String
    
    If pBrowseAction = BrowseActions.Dialog Then
        CommonDialog1.InitDir = CurDir
        CommonDialog1.DialogTitle = pBrowseTitle
        CommonDialog1.Filter = pFileFilter
        CommonDialog1.FilterIndex = pFileFilterIndex
        If Not ((CommonDialog1.Flags And &H4) = &H4) Then CommonDialog1.Flags = CommonDialog1.Flags + &H4
        CommonDialog1.CancelError = True
        
        On Error Resume Next
        CommonDialog1.Action = 1
        If Not Err Then
            sPath = CommonDialog1.FileName
            If InStrRev(CommonDialog1.FileName, "\") > 0 Then
                ChDir Left(CommonDialog1.FileName, InStrRev(CommonDialog1.FileName, "\") - 1)
            End If
        End If
        If Err Then
            sPath = ""
            Err.Clear
        End If
        On Error GoTo 0
        
    Else
    
        Dim BI As BROWSEINFO
        Dim nFolder As Long
        Dim IDL As ITEMIDLIST
        Dim pidl As Long
        With BI
            On Error Resume Next
            .hOwner = UserControl.Extender.Parent.hwnd
            If Err Then .hOwner = UserControl.hwnd
            On Error GoTo 0
        
            nFolder = GetFolderValue(pBrowseAction)
        
            If SHGetSpecialFolderLocation(ByVal .hOwner, ByVal nFolder, IDL) = NOERROR Then
                .pidlRoot = IDL.mkid.cb
            End If
        
            .pszDisplayName = String$(MAX_PATH, 0)
        
            .lpszTitle = pBrowseTitle
        
            .ulFlags = GetReturnType()
        
        End With
     
        pidl = SHBrowseForFolder(BI)
      
        If pidl = 0 Then
            sPath = ""
        Else
        
            sPath = String$(MAX_PATH, 0)
            SHGetPathFromIDList ByVal pidl, ByVal sPath
    
            sPath = Left(sPath, InStr(sPath, vbNullChar) - 1)
    
        End If
       
    End If
    
    pBrowseReturn = sPath
    pCurrentAction = IIf(sPath = "", BrowseActions.Cacneled, BrowseActions.None)
    
    Browse = sPath
End Function


Private Sub UserControl_Initialize()

    Command1.Top = -Command1.Height
    Command1.Left = -Command1.Width
    Command1.Visible = True
    
    buttonOver.Interval = 50
    
    IsMouseDown = False
    CancelResize = False
    BorderWidth = 3 * Screen.TwipsPerPixelY
    
End Sub

Private Sub UserControl_InitProperties()
    
    BorderStyle = BorderStyles.MouseOver
    
    BrowseTitle = "Browse"
    BrowseAction = BrowseActions.Desktop
    
    FileFilter = "All Files|*.*"
    FileFilterIndex = 1
    
    ToolTipText = ""
    
    Enabled = True
    
    Set ImageOver = imgOver.Picture
    Set ImageOut = imgOut.Picture
    Set ImageDisabled = imgDisabled.Picture

End Sub

Private Sub UserControl_Paint()
    On Error Resume Next
   ' BackColor = UserControl.Parent.BackColor
    If Err Then Err.Clear
    On Error GoTo 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        
        BorderStyle = .ReadProperty("pBorderStyle", 2)
        
        BrowseTitle = .ReadProperty("bBrowseTitle", "Browse")
        
        BrowseAction = .ReadProperty("bBrowseAction", 1)
        
        FileFilter = .ReadProperty("bFileFilter", "All Files|*.*")
        
        FileFilterIndex = .ReadProperty("bFileFilterIndex", 1)
       
        ToolTipText = .ReadProperty("bToolTipText", "")
        
        Enabled = .ReadProperty("bEnabled", True)
        
        Set ImageOver = .ReadProperty("pImageOver", imgOver.Picture)
        Set ImageOut = .ReadProperty("pImageOut", imgOut.Picture)
        Set ImageDisabled = .ReadProperty("pImageDisabled", imgDisabled.Picture)
       
    End With
End Sub

Private Sub SetDimensions()
    On Error Resume Next
    
    CancelResize = True

    UserControl.Height = (21 * Screen.TwipsPerPixelY)
    UserControl.Width = (21 * Screen.TwipsPerPixelX)
    
    buttonImg.Left = BorderWidth
    buttonImg.Top = BorderWidth
    
    buttonImg.Height = UserControl.Height - BorderWidth
    buttonImg.Width = UserControl.Width - BorderWidth
    
    Line1.X1 = 0
    Line1.X2 = 0
    Line1.Y1 = 0
    Line1.Y2 = UserControl.Height
    
    Line2.X1 = 0
    Line2.X2 = UserControl.Width
    Line2.Y1 = 0
    Line2.Y2 = 0
    
    Line3.X1 = 0
    Line3.X2 = UserControl.Width
    Line3.Y1 = UserControl.Height - Screen.TwipsPerPixelY
    Line3.Y2 = UserControl.Height - Screen.TwipsPerPixelY
    
    Line4.X1 = UserControl.Width - Screen.TwipsPerPixelX
    Line4.X2 = UserControl.Width - Screen.TwipsPerPixelX
    Line4.Y1 = 0
    Line4.Y2 = UserControl.Height
    
    Shape1.Top = Screen.TwipsPerPixelY
    Shape1.Left = Screen.TwipsPerPixelX
    
    Shape1.Width = UserControl.Width - (Screen.TwipsPerPixelX * 2)
    Shape1.Height = UserControl.Height - (Screen.TwipsPerPixelY * 2)
    
    CancelResize = False
    
    If Err Then Err.Clear
    On Error GoTo 0
End Sub

Private Sub UserControl_Resize()

    If Not CancelResize Then
        SetDimensions
    End If

End Sub

Private Sub UserControl_Terminate()
    buttonOver.Enabled = False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag

        .WriteProperty "pBorderStyle", pBorderStyle, BorderStyles.MouseOver

        .WriteProperty "bBrowseTitle", pBrowseTitle, "Browse"
        .WriteProperty "bBrowseAction", pBrowseAction, BrowseActions.Desktop
        .WriteProperty "bFileFilter", pFileFilter, "All Files|*.*"
        .WriteProperty "bFileFilterIndex", pFileFilterIndex, 1

        .WriteProperty "bToolTipText", pToolTipText, ""

        .WriteProperty "bEnabled", pEnabled, True

        .WriteProperty "pImageOver", pImageOver, imgOver.Picture
        .WriteProperty "pImageOut", pImageOut, imgOut.Picture
        .WriteProperty "pImageDisabled", pImageDisabled, imgDisabled.Picture

    End With
End Sub




