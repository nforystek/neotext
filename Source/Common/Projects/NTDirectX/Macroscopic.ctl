VERSION 5.00
Begin VB.UserControl Macroscopic 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5160
   ClipControls    =   0   'False
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   Enabled         =   0   'False
   LockControls    =   -1  'True
   ScaleHeight     =   4350
   ScaleWidth      =   5160
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1245
      Top             =   3045
   End
   Begin VB.PictureBox Renderview 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   2400
      Left            =   0
      ScaleHeight     =   2400
      ScaleWidth      =   2820
      TabIndex        =   0
      Top             =   0
      Width           =   2820
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   480
         Left            =   690
         Top             =   585
         Width           =   1425
      End
      Begin VB.Image Image2 
         Height          =   4515
         Left            =   1155
         Picture         =   "Macroscopic.ctx":0000
         Top             =   1575
         Visible         =   0   'False
         Width           =   5955
      End
   End
End
Attribute VB_Name = "Macroscopic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function LoadFolder(ByVal PathName As String) As Boolean
    On Error GoTo failload

    'If frmMain.SerialStack = False Then
    '    frmMain.SerialStack = True
     '   frmMain.Serialize ParseScript(PathName & "\Index.vbx")
    '    frmMain.SerialStack = False
    'End If

    'Debug.Print Planets.Count; Molecules.Count; All.Count

    Exit Function
failload:
    MsgBox "Unable to """ & PathName & "\Index.vbx""" & vbCrLf & Err.Description, vbCritical
    Err.Clear
End Function
Friend Property Get Parent() As Object
    On Error GoTo noclient
    
    Set Parent = UserControl.Parent
    
    Exit Property
noclient:
    Set Parent = frmMain
    StopGame = True
End Property
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property
Friend Property Get State() As Boolean
    State = Not PauseGame
End Property
Public Property Get Width() As Long
    Width = UserControl.ScaleWidth / VB.Screen.TwipsPerPixelX
End Property
Public Property Get Height() As Long
    Height = UserControl.ScaleHeight / VB.Screen.TwipsPerPixelY
End Property
Public Property Get Top() As Long
    Top = UserControl.ScaleTop / VB.Screen.TwipsPerPixelY
End Property
Public Property Get Left() As Long
    Left = UserControl.ScaleLeft / VB.Screen.TwipsPerPixelX
End Property

Private Sub Timer1_Timer()
    Timer1.Enabled = False

    RenderFrame Me
        
    If StopGame Then
        Unload Parent
    Else
        Timer1.Enabled = True
    End If
    
End Sub

Friend Sub PauseRendering()
    If (Not PauseGame) Then
        
        frmMain.Visible = False
        Image2.Visible = True
        Shape1.Visible = True
        
        PauseGame = True
        TermGameData Me
        TermDirectX Me
    
        If TrapMouse Or FullScreen Then
            VB.Screen.MousePointer = 0
        End If
        
        Timer1.Enabled = False
    
    End If
        
End Sub

Friend Sub ResumeRendering()
    On Error GoTo fault
    If PauseGame Then
        
        Image2.Visible = False
        Shape1.Visible = False
        frmMain.Visible = True
        
        UserControl_Resize
        
        PauseGame = False
        
        On Error GoTo fault
        InitDirectX Me
        InitGameData Me
        On Error GoTo 0
        
        If (TrapMouse Or FullScreen) And (Bindings.Controller = Trapping Or Bindings.Controller = Hidden) Then
            VB.Screen.MousePointer = 99
        End If
        
        Timer1.Enabled = True
                
    End If

    On Error GoTo 0
Exit Sub
fault:
    TermDirectX Me
    Err.Clear
End Sub

Private Sub UserControl_Initialize()

    FPSCount = 36
    PauseGame = False
    TrapMouse = False
    
    Resolution = GetSetting(AppEXE(True, True), "System", "Resolution", "1024x768")
    FullScreen = GetSetting(AppEXE(True, True), "System", "FullScreen", False)
    SilentMode = GetSetting(AppEXE(True, True), "System", "SilentMode", False)

    Set VB.Screen.MouseIcon = PictureFromByteStream(LoadResData(2, "CUSTOM"))
    
    BackColor = ConvertColor(SystemColorConstants.vbButtonFace)
        
    UserControl_Resize

    If Not (IsDesignMode Or IsRunningMode) Then
                
        PauseGame = True
        ResumeRendering
        Timer1.Enabled = True
        
    End If

End Sub
Friend Property Get Viewport() As Control
    Set Viewport = Renderview
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 112 Then ShowSetup = True
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next

    Renderview.Top = 0
    Renderview.Left = 0
    Renderview.Width = UserControl.Width
    Renderview.Height = UserControl.Height
    Image2.Left = (UserControl.Width / 2) - (Image2.Width / 2)
    Image2.Top = (UserControl.Height / 2) - (Image2.Height / 2)
    Shape1.Left = 0
    Shape1.Top = 0
    Shape1.Width = UserControl.Width
    Shape1.Height = UserControl.Height
    
    If Not FullScreen Then

        SetParent frmMain.hwnd, Renderview.hwnd
        frmMain.Move 0, 0, Renderview.Width, Renderview.Height

    Else
        frmMain.Width = CSng(NextArg(Resolution, "x")) * VB.Screen.TwipsPerPixelX
        frmMain.Height = CSng(RemoveArg(Resolution, "x")) * VB.Screen.TwipsPerPixelY
        SetParent frmMain.hwnd, 0
    End If

End Sub

Private Sub UserControl_Show()
    UserControl_Resize
End Sub

Private Sub UserControl_Terminate()
    If Not (IsDesignMode Or IsRunningMode) Then
        PauseRendering
    End If
    StopGame = True
    
    Unload frmMain
    
End Sub
