VERSION 5.00
Object = "{BA98913A-7219-4720-8E5D-F3D8E058DF1B}#388.0#0"; "NTImaging10.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmStudio 
   Caption         =   "adPatch"
   ClientHeight    =   10605
   ClientLeft      =   135
   ClientTop       =   -1635
   ClientWidth     =   21750
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "frmStudio.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10605
   ScaleWidth      =   21750
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer UpdateTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   12675
      Top             =   6135
   End
   Begin VB.PictureBox HelperImage 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   705
      Left            =   5055
      ScaleHeight     =   705
      ScaleWidth      =   840
      TabIndex        =   15
      Top             =   5235
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.PictureBox Designer 
      Align           =   4  'Align Right
      BackColor       =   &H00000000&
      Height          =   6105
      Left            =   20490
      ScaleHeight     =   6045
      ScaleWidth      =   1200
      TabIndex        =   5
      Top             =   4500
      Width           =   1260
   End
   Begin VB.PictureBox Symbols 
      Align           =   4  'Align Right
      Height          =   6105
      Left            =   15360
      ScaleHeight     =   6045
      ScaleWidth      =   5070
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4500
      Visible         =   0   'False
      Width           =   5130
      Begin VB.PictureBox SymbolBitmap 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   720
         Left            =   4095
         ScaleHeight     =   48
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   48
         TabIndex        =   14
         Top             =   1830
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.PictureBox SymbolEdit 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1680
         Left            =   516
         ScaleHeight     =   1650
         ScaleWidth      =   3000
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1632
         Width           =   3024
      End
   End
   Begin VB.PictureBox Help 
      Align           =   1  'Align Top
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4140
      Left            =   0
      ScaleHeight     =   4080
      ScaleWidth      =   21690
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   21750
      Begin VB.Image Image1 
         Height          =   3630
         Left            =   2580
         Picture         =   "frmStudio.frx":23D2
         Stretch         =   -1  'True
         Top             =   195
         Width           =   14340
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3276
      Top             =   1344
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.PictureBox ModePanel 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      Height          =   360
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   21690
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4140
      Width           =   21750
      Begin VB.PictureBox SymbolButtons 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5610
         ScaleHeight     =   300
         ScaleWidth      =   2970
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   2970
         Begin VB.CommandButton Command4 
            Caption         =   "Cut"
            Height          =   300
            Left            =   -15
            TabIndex        =   21
            Top             =   0
            Width           =   555
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Copy"
            Height          =   300
            Left            =   555
            TabIndex        =   20
            Top             =   0
            Width           =   630
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Paste"
            Height          =   300
            Left            =   1200
            TabIndex        =   19
            Top             =   0
            Width           =   690
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Clear"
            Height          =   300
            Left            =   2235
            TabIndex        =   18
            Top             =   0
            Width           =   690
         End
      End
      Begin VB.PictureBox DrawStyles 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   10095
         ScaleHeight     =   300
         ScaleWidth      =   4170
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   4170
         Begin VB.CheckBox DoubleThick 
            Caption         =   "Double"
            Height          =   195
            Left            =   3300
            TabIndex        =   16
            Top             =   60
            Width           =   825
         End
         Begin VB.PictureBox CrossStitches 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2685
            ScaleHeight     =   300
            ScaleWidth      =   315
            TabIndex        =   12
            Top             =   -15
            Width           =   315
            Begin VB.Line Line4 
               Visible         =   0   'False
               X1              =   240
               X2              =   60
               Y1              =   240
               Y2              =   60
            End
            Begin VB.Line Line2 
               Visible         =   0   'False
               X1              =   240
               X2              =   60
               Y1              =   150
               Y2              =   150
            End
            Begin VB.Line Line3 
               Visible         =   0   'False
               X1              =   60
               X2              =   240
               Y1              =   240
               Y2              =   60
            End
            Begin VB.Line Line1 
               BorderColor     =   &H80000012&
               Visible         =   0   'False
               X1              =   150
               X2              =   150
               Y1              =   240
               Y2              =   60
            End
         End
         Begin VB.CheckBox RemovalTool 
            Caption         =   "Eraser"
            Height          =   195
            Left            =   135
            TabIndex        =   1
            Top             =   60
            Width           =   765
         End
         Begin VB.CheckBox Precision 
            Caption         =   "Drawing"
            Height          =   195
            Left            =   990
            TabIndex        =   2
            Top             =   60
            Width           =   930
         End
         Begin VB.VScrollBar VScroll2 
            Enabled         =   0   'False
            Height          =   252
            Left            =   3000
            Max             =   0
            Min             =   -5
            TabIndex        =   3
            Top             =   15
            Value           =   -1
            Width           =   252
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Stencil:"
            Height          =   225
            Left            =   2070
            TabIndex        =   13
            Top             =   60
            Width           =   600
         End
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   855
         Left            =   30
         TabIndex        =   0
         Top             =   0
         Width           =   3150
         _ExtentX        =   5556
         _ExtentY        =   1508
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Designer"
               Key             =   "designer"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Symbols"
               Key             =   "symbols"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Pattern"
               Key             =   "pattern"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox Browser 
      Align           =   3  'Align Left
      ClipControls    =   0   'False
      Height          =   6105
      Left            =   0
      ScaleHeight     =   6045
      ScaleWidth      =   2370
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4500
      Width           =   2436
      Begin NTImaging10.Gallery Gallery1 
         Height          =   1335
         Left            =   375
         TabIndex        =   4
         Top             =   855
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   2355
         FilePath        =   ""
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuDash23 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuDash892qw7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "E&xport..."
      End
      Begin VB.Menu mnuDash3902 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Umdo"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "Re&do"
         Enabled         =   0   'False
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuDash27834 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add Material..."
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "&Remove Material"
      End
      Begin VB.Menu mnuDash327832 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t Pattern"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy Pattern"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste Pattern"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDash27382 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "C&lear Pattern"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDash1234 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAdvanced 
         Caption         =   "Ad&vanced Mode"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuSymbols 
         Caption         =   "&Symbols"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuPattern 
         Caption         =   "&Pattern"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuOnScreenHelp 
         Caption         =   "&On Screen Quick Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuDash3287 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About Splash Screen"
      End
   End
End
Attribute VB_Name = "frmStudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CancelUpdate As Boolean

Private SymbolPen As Boolean

Public LastX As Single
Public LastY As Single

Public LastWidth As Single
Public LastHeight As Single

Public UndoRecord As Long
Public UndoBuffer As New NTNodes10.Collection
Public UndoAction As String

Public ExportCapture As Long


Public Sub UpdateGallery()

    Dim backSel As Long
    backSel = Gallery1.ListIndex
    
    Gallery1.Clear
    Gallery1.Stretch = False


startover:
    Dim rs As New ADODB.Recordset
    
    db.rsQuery rs, "SELECT * FROM Materials;"
    
    If Not db.rsEnd(rs) Then
        rs.MoveFirst
        Do


            StringToSymbol rs("Symbol")

            Gallery1.AddImageByPicture SymbolBitmap, rs("ID")
            Gallery1.BackgroundColors(Gallery1.count - 1) = rs("Color")
            

            If Not PathExists(GetSymbolFile(rs("ID")), True) Then
                SavePicture SymbolBitmap.Image, GetSymbolFile(rs("ID"))
            End If
            
            If Not PathExists(GetColorFile(rs("Color")), True) Then
                HelperImage.Cls
                HelperImage.Line (1, 1)-(HelperImage.Width, HelperImage.Height), rs("Color"), BF
                SavePicture HelperImage.Image, GetColorFile(rs("Color"))
            End If
            
            rs.MoveNext
        Loop Until db.rsEnd(rs)
    Else
        CreateColor 0, False
        GoTo startover
    End If

    db.rsClose rs
    If backSel > -1 And backSel < Gallery1.count Then
        Gallery1.ListIndex = backSel
    End If
    
    Gallery1_Click
End Sub
Friend Sub MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 And (TrapMouse = 0) Then
        If ((Abs(LastX - X) <> 0) And (Abs(LastY - Y) <> 0)) Then

            Player.Location.X = ((Player.Location.X + (LastX - X)))
            Player.Location.Y = ((Player.Location.Y - (LastY - Y)))
            
        End If
    Else
        TrapMouse = 0
        If Not (frmMain.MousePointer = 99) And ((GetActiveWindow = frmStudio.hwnd) Or (GetActiveWindow = frmMain.hwnd)) Then
            frmMain.MousePointer = 99
            frmMain.MouseIcon = LoadPicture(AppPath & "Base\mouse.cur")
        ElseIf (Not (frmMain.MousePointer = 0)) And ((GetActiveWindow <> frmStudio.hwnd) And (GetActiveWindow <> frmMain.hwnd)) Then
            frmMain.MousePointer = 0
        End If

    End If
    
    LastX = X
    LastY = Y

End Sub

Private Sub Command1_Click()
    If MsgBox("Are you sure you want to clear the current pattern?", vbQuestion + vbYesNo) = vbYes Then
 
        UndoCOmmit
        
        Dim BlockX As Single
        Dim BlockY As Single
        For BlockX = 0 To SymbolWidth - 1
            For BlockY = 0 To SymbolHeight - 1
                If SymbolBitmap.Point(BlockX, BlockY) = vbWhite Then
                   UndoAction = UndoAction & "B" & Gallery1.BackgroundColors(Gallery1.ListIndex) & "," & BlockX & "," & BlockY & vbCrLf
                Else
                   UndoAction = UndoAction & "W" & Gallery1.BackgroundColors(Gallery1.ListIndex) & "," & BlockX & "," & BlockY & vbCrLf
                End If
                
                SymbolBitmap.PSet (BlockX, BlockY), vbWhite
            Next
        Next
        
        UndoCOmmit
        
        PaintSymbol
        UpdateTimer.Enabled = True
        
    End If
End Sub

Private Sub Command2_Click()
    If Clipboard.GetFormat(vbCFBitmap) Then
        
       If MsgBox("Are you sure you want to paste the clipboard contents over the current pattern?", vbQuestion + vbYesNo) = vbYes Then
    
           UndoCOmmit
    
           Dim BlockX As Single
           Dim BlockY As Single
           For BlockX = 0 To SymbolWidth - 1
               For BlockY = 0 To SymbolHeight - 1
                If SymbolBitmap.Point(BlockX, BlockY) = vbWhite Then
                   UndoAction = UndoAction & "B" & Gallery1.BackgroundColors(Gallery1.ListIndex) & "," & BlockX & "," & BlockY & vbCrLf
                Else
                   UndoAction = UndoAction & "W" & Gallery1.BackgroundColors(Gallery1.ListIndex) & "," & BlockX & "," & BlockY & vbCrLf
                End If
               Next
           Next
           
           UndoCOmmit
           
           SymbolBitmap.Picture = Clipboard.GetData(vbCFBitmap)


           For BlockX = 0 To SymbolWidth - 1
               For BlockY = 0 To SymbolHeight - 1
                If SymbolBitmap.Point(BlockX, BlockY) = vbWhite Then
                   UndoAction = UndoAction & "B" & Gallery1.BackgroundColors(Gallery1.ListIndex) & "," & BlockX & "," & BlockY & vbCrLf
                Else
                   UndoAction = UndoAction & "W" & Gallery1.BackgroundColors(Gallery1.ListIndex) & "," & BlockX & "," & BlockY & vbCrLf
                End If
               Next
           Next
                      
           UndoCOmmit
           
           PaintSymbol
           UpdateTimer.Enabled = True
           
       End If
    Else
        MsgBox "The clipbard does not contain a valid pattern or bitmap image.", vbExclamation
        
    
    End If
End Sub

Private Sub Command3_Click()
 
    UndoCOmmit
        
    Clipboard.Clear
    
    Clipboard.SetData SymbolBitmap.Picture, vbCFBitmap
        
        
End Sub

Private Sub Command4_Click()
    If MsgBox("Are you sure you want to cut the current pattern and place it in the clipboard?", vbQuestion + vbYesNo) = vbYes Then
 
        UndoCOmmit
        
        Clipboard.SetData SymbolBitmap.Picture, vbCFBitmap
        
        
 
        Dim BlockX As Single
        Dim BlockY As Single
        For BlockX = 0 To SymbolWidth - 1
            For BlockY = 0 To SymbolHeight - 1
            
                SymbolBitmap.PSet (BlockX, BlockY), vbWhite
            
                UndoAction = UndoAction & "W" & Gallery1.BackgroundColors(Gallery1.ListIndex) & "," & BlockX & "," & BlockY & vbCrLf
            Next
        Next
        
        UndoCOmmit
        
        PaintSymbol
        UpdateTimer.Enabled = True
        
    End If
End Sub

Private Sub Designer_Click()
    Me.Hide
End Sub

Friend Sub KeyDown(KeyCode As Integer, Shift As Integer)

End Sub

Friend Sub KeyPress(KeyAscii As Integer)

End Sub

Private Sub HitMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseMove Button, Shift, X, Y
End Sub

Private Sub Designer_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyDown KeyCode, Shift
End Sub

Private Sub Designer_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Designer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseMove Button, Shift, X, Y
End Sub

Private Sub Browser_Resize()
    Gallery1.Top = 0
    Gallery1.Left = 0
    Gallery1.Width = Browser.ScaleWidth
    Gallery1.Height = Browser.ScaleHeight
    
End Sub

Private Sub Designer_Resize()
'    frmMain.Picture1.Width = Designer.ScaleWidth
'    frmMain.Picture1.Height = Designer.ScaleHeight
End Sub

Private Sub Form_Load()
    
    ChDir AppPath & "Base\Patterns"
    
    ThatchIndex = 1
    Load frmMain
    VScroll2.Tag = 0
    
    Precision.Value = 1
    
    AdvancedGUI
    frmMain.Width = Screen.Width ' * Screen.TwipsPerPixelY
    frmMain.Height = Screen.Height ' * Screen.TwipsPerPixelX

    frmMain.Picture1.Width = Screen.Width ' * Screen.TwipsPerPixelY
    frmMain.Picture1.Height = Screen.Height ' * Screen.TwipsPerPixelX
    
    
    SetParent frmMain.hwnd, Designer.hwnd
    
    frmMain.Show
            
    On Error GoTo fault
    InitDirectX
    InitGameData
    On Error GoTo 0

    TabStrip1.SelectedItem = TabStrip1.Tabs("designer")
    VScroll2.Value = -5
    SymbolBitmap.Width = SymbolWidth * Screen.TwipsPerPixelX
    SymbolBitmap.Height = SymbolHeight * Screen.TwipsPerPixelY

    DrawSytleUIEnable
    TabStrip1_Click
    VScroll2_Change
        
    ResizeMultidimArray

    
    HelperImage.Width = (Screen.TwipsPerPixelX * SymbolWidth)
    HelperImage.Height = (Screen.TwipsPerPixelY * SymbolHeight)
  
    UpdateGallery
    
    modProj.ProjPath = "Untitled"
    modProj.Dirty = False

    modProj.DecalWidth = HelperImage.Width
    modProj.DecalHeight = HelperImage.Height
    modProj.CleanUpProj
    modProj.CreateProj

    Me.Show
    DoEvents
    GotoCenter
Exit Sub
fault:

    TermDirectX
    
    MsgBox "There was an error initializing the game.  Please try reinstalling it or contact support." & vbCrLf & "Error Infromation: " & Err.number & ", " & Err.Description, vbOKOnly + vbInformation, App.Title
    Err.Clear
    End

End Sub

Private Sub Form_LostFocus()
    DoNotFocused
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If CloseProject Then

        Call SetWindowWord(frmMain.hwnd, SWW_HPARENT, 0&)
    Else
        Cancel = True
    End If
End Sub

Private Sub Form_Resize()
    If Me.ScaleWidth - Browser.Width > 0 Then
        Designer.Width = Me.ScaleWidth - Browser.Width
        Symbols.Width = Designer.Width

    End If
    Symbols_Resize

End Sub

Private Sub Form_Unload(Cancel As Integer)

    TermGameData
    TermDirectX
        
    Unload frmMain
    
    CleanupProjFiles
            
    Set db = Nothing
    
    End
End Sub

Public Function SaveProject(Optional ByVal Prompt As Boolean = False) As Boolean
    
    If (Not PathExists(ProjPath, True)) Or ProjPath = "" Or Prompt Then
        On Error Resume Next
        CommonDialog1.CancelError = True
        CommonDialog1.DefaultExt = ".csp"
        CommonDialog1.DialogTitle = "Save Cross Stitch Project"
        CommonDialog1.Filter = "Cross Stitch Project (*.csp)|*.csp|All Files (*.*)|*.*"
        If CommonDialog1.FileName <> "" Then CommonDialog1.FileName = Replace(CommonDialog1.FileName, ".pdf", ".csp", , , vbTextCompare)
        CommonDialog1.FilterIndex = 1
        CommonDialog1.Flags = &H4 And &H200000 And &H8000 And &H2
        CommonDialog1.ShowSave
        If Err Then
            Err.Clear
            SaveProject = False
        Else
            SaveProject = True
        End If
        On Error GoTo 0
        If SaveProject Then
            UndoCOmmit
            
            TabStrip1.SelectedItem = TabStrip1.Tabs("designer")
            TabStrip1_Click
            
            ProjPath = CommonDialog1.FileName
            SaveProject = modProj.WriteToDisk
            modProj.Dirty = False
        End If
    Else
        UndoCOmmit
        TabStrip1.SelectedItem = TabStrip1.Tabs("designer")
        TabStrip1_Click
            
        SaveProject = modProj.WriteToDisk
        modProj.Dirty = False
        SaveProject = True
    End If
End Function

Private Function CloseProject() As Boolean

    If modProj.Dirty Then
        Select Case MsgBox("Save changes before closing this project?", vbYesNoCancel)
            Case vbCancel
                CloseProject = False
            Case vbYes
                CloseProject = SaveProject
            Case vbNo
                CloseProject = True
        End Select
    Else
        CloseProject = True
    End If

End Function

Private Sub DrawSytleUIEnable()
    
    If RemovalTool.Value Then
        VScroll2.Tag = VScroll2.Value
        VScroll2.Value = 0
    ElseIf VScroll2.Value = 0 And VScroll2.Tag <> 0 Then
        VScroll2.Value = VScroll2.Tag
        VScroll2.Tag = 0
    End If
    
    Line1.BorderColor = IIf(Precision.Value, SystemColorConstants.vbWindowText, SystemColorConstants.vbGrayText)
    Line2.BorderColor = Line1.BorderColor
    Line3.BorderColor = Line2.BorderColor
    Line4.BorderColor = Line3.BorderColor
    
    VScroll2.Enabled = (Precision.Value)

    VScroll2_Change
End Sub

Private Sub Gallery1_Click()

    
    Select Case TabStrip1.SelectedItem.Key
            
        Case "symbols"

            SymbolEdit.Cls

            If PathExists(GetSymbolFile, True) Then
                SymbolBitmap.Picture = LoadPicture(GetSymbolFile)
            Else
                SymbolBitmap.Cls
            End If
            PaintSymbol
            
            
        Case "pattern"
        Case "designer"
    End Select
End Sub

Public Sub LinkedLines_Click()
    If CancelUpdate = False Then
        CancelUpdate = True
            
        Precision.Value = 0
        RemovalTool.Value = 0

        DrawSytleUIEnable
        CancelUpdate = False
    End If
End Sub

Private Sub Gallery1_Selected(ByVal ListIndex As Integer, Cancel As Boolean)
    If UpdateTimer.Enabled Then
        UpdateTimer.Enabled = False
        UpdateTimer_Timer
    End If
    
End Sub

Private Sub Help_Resize()
    Image1.Left = (Help.ScaleWidth / 2) - (Image1.Width / 2)
    Image1.Top = (Help.ScaleHeight / 2) - (Image1.Height / 2)
    
End Sub

Public Function GetMaterialID(ByVal Color As Long) As String
    Dim cnt As Long
    If Gallery1.count > 0 Then
        For cnt = 0 To Gallery1.count - 1
            If Gallery1.BackgroundColors(cnt) = Color Then
                GetMaterialID = ColorToHex(frmStudio.Gallery1.Info(cnt))
            End If
        Next
    End If
End Function

Private Function GetColorFile(Optional ByVal Color As Long = -1) As String
    If Color = -1 Then
        GetColorFile = AppPath & "Base\Stitchings\FlossThreads\" & ColorToHex(Gallery1.BackgroundColors(Gallery1.ListIndex)) & ".bmp"
    Else
        GetColorFile = AppPath & "Base\Stitchings\FlossThreads\" & ColorToHex(Color) & ".bmp"
    End If
End Function

Private Function GetSymbolFile(Optional ByVal id As Long = -1) As String
    If id = -1 Then
        If Gallery1.count > 0 Then
            GetSymbolFile = AppPath & "Base\Stitchings\LegendKeys\" & ColorToHex(Gallery1.Info(Gallery1.ListIndex)) & ".bmp"
        End If
    Else
        GetSymbolFile = AppPath & "Base\Stitchings\LegendKeys\" & ColorToHex(id) & ".bmp"
    End If
End Function

Public Function CreateColor(ByVal Color As Long, ByVal OnlyIfMissing As Boolean, Optional ByVal Symbol As String) As Boolean
    Dim rs As New ADODB.Recordset
    If Not PathExists(GetColorFile(Color), True) Then
        HelperImage.Line (1, 1)-(HelperImage.Width, HelperImage.Height), Color, BF
        SavePicture HelperImage.Image, GetColorFile(Color)
        OnlyIfMissing = False
        CreateColor = True
    End If

    db.rsQuery rs, "SELECT * FROM Materials WHERE Color=" & Color & ";"
    If db.rsEnd(rs) Or (Not OnlyIfMissing) Or (Symbol <> "") Then
        
        If db.rsEnd(rs) Then
            db.dbQuery "INSERT INTO Materials (Color, Symbol) VALUES (" & Color & ", 'NEWRECORD');"
            db.rsQuery rs, "SELECT * FROM Materials WHERE Symbol='NEWRECORD' AND Color=" & Color & ";"
        End If

        SymbolBitmap.Cls
        
        If (Symbol <> "") Then StringToSymbol Symbol

        SavePicture SymbolBitmap.Image, GetSymbolFile(rs("ID"))
                
        db.dbQuery "UPDATE Materials SET Symbol='" & SymbolToString & "' WHERE ID=" & rs("ID") & ";"
        CreateColor = True
    End If
    db.rsClose rs
End Function

Private Function RemoveMaterial(ByVal Color As Long) As String
    RemoveMaterial = "R " & Color & ","
    RemoveColorUsed Color
    Dim rs As New ADODB.Recordset
    db.rsQuery rs, "SELECT * FROM Materials WHERE Color = " & Color & ";"
    If Not db.rsEnd(rs) Then
        If PathExists(GetSymbolFile(rs("ID")), True) Then Kill GetSymbolFile(rs("ID"))
        RemoveMaterial = RemoveMaterial & rs("Symbol")
        db.dbQuery "DELETE FROM Materials WHERE ID=" & rs("ID") & ";"
    End If
    If PathExists(GetColorFile(Color), True) Then Kill GetColorFile(Color)
    db.rsClose rs
End Function

Private Sub mnuAdd_Click()

    Load frmColor
    frmColor.Show 1, frmStudio

    If frmColor.Color <> -1 Then

        CreateColor frmColor.Color, False

        UndoAction = UndoAction & "M" & frmColor.Color & "," & SymbolToString & vbCrLf
    
        UndoEnables
        
        UpdateGallery
        Gallery1.ListIndex = Gallery1.count - 1
        Gallery1_Click
        
        modProj.CleanUpProj
        modProj.CreateProj
        
    End If
    Unload frmColor

End Sub

Private Sub mnuExport_Click()
'    Load frmThatch
'    frmThatch.SetupExport
'
'    frmThatch.Show 1
'
'    If frmThatch.Tag = "OK" Then
    
        On Error Resume Next
        CommonDialog1.CancelError = True
        CommonDialog1.DefaultExt = ".pdf"
        CommonDialog1.DialogTitle = "Export Cross Stitch Project View"
        CommonDialog1.Filter = "Windows Bitmap (*.bmp)|*.bmp|All Files (*.*)|*.*"
        If CommonDialog1.FileName <> "" Then CommonDialog1.FileName = Replace(CommonDialog1.FileName, ".csp", ".pdf", , , vbTextCompare)
        CommonDialog1.FilterIndex = 1
        CommonDialog1.Flags = &H4 And &H200000 And &H8000 And &H2
        CommonDialog1.ShowSave
        If Err Then
            Err.Clear
        Else
            UndoCOmmit
            
            On Error GoTo 0

            
            Set TabStrip1.SelectedItem = TabStrip1.Tabs("designer")
            TabStrip1_Click
'            ExportCapture = 1
            
            
            Dim dViewPort As D3DVIEWPORT8
            Dim DSurface As D3DXRenderToSurface
            Dim dm As D3DDISPLAYMODE
            Dim pal As PALETTEENTRY
            Dim rct As DxVBLibA.RECT
            Dim BufferedTexture As Direct3DTexture8
            Dim ReflectRenderTarget As Direct3DSurface8
            Dim ReflectFrontBuffer As Direct3DSurface8

            DDevice.GetViewport dViewPort
            
            Set DSurface = D3DX.CreateRenderToSurface(DDevice, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY, Display.Format, False, D3DFMT_D16)

            Set ReflectRenderTarget = DDevice.CreateRenderTarget((frmMain.Width / Screen.TwipsPerPixelX), (frmMain.Height / Screen.TwipsPerPixelY), CONST_D3DFORMAT.D3DFMT_A8R8G8B8, D3DMULTISAMPLE_NONE, True)
            Set BufferedTexture = DDevice.CreateTexture((frmMain.Width / Screen.TwipsPerPixelX), (frmMain.Height / Screen.TwipsPerPixelY), 1, CONST_D3DUSAGEFLAGS.D3DUSAGE_RENDERTARGET, CONST_D3DFORMAT.D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
            Set ReflectFrontBuffer = BufferedTexture.GetSurfaceLevel(0)

            
            dViewPort.Width = Screen.Width / Screen.TwipsPerPixelX
            dViewPort.Height = Screen.Height / Screen.TwipsPerPixelY

            DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, BackColor, 1, 0
            
            DSurface.BeginScene DDevice.GetRenderTarget, dViewPort
            
            SetupWorld

            RenderView False, True
        

            DSurface.EndScene

            DDevice.GetDisplayMode dm

            rct.Top = 0
            rct.Left = 0

            rct.Right = dViewPort.Width
            rct.Bottom = dViewPort.Height
            

            D3DX.SaveSurfaceToFile Replace(CommonDialog1.FileName, ".pdf", " Display.bmp", , , vbTextCompare), D3DXIFF_BMP, DDevice.GetRenderTarget, pal, rct
       
            Set TabStrip1.SelectedItem = TabStrip1.Tabs("pattern")
            TabStrip1_Click


            DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, BackColor, 1, 0
            
            DSurface.BeginScene DDevice.GetRenderTarget, dViewPort
            
            SetupWorld

            RenderView False, True
        

            DSurface.EndScene

            DDevice.GetDisplayMode dm

            rct.Top = 0
            rct.Left = 0

            rct.Right = dViewPort.Width
            rct.Bottom = dViewPort.Height
            

            D3DX.SaveSurfaceToFile Replace(CommonDialog1.FileName, ".pdf", " Pattern.bmp", , , vbTextCompare), D3DXIFF_BMP, DDevice.GetRenderTarget, pal, rct
            
'            Mirrors.Add D3DX.CreateTextureFromFileEx(DDevice, GetTemporaryFolder & "\" & Electrons.Key(i) & ".bmp", _
'                DViewPort.Width, DViewPort.Height, D3DX_FILTER_NONE, 0, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, _
'                D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, Transparent, ByVal 0, ByVal 0), Electrons.Key(i)


            

            
          '  Set DSurface = Nothing
            
        End If
    
'    End If
'    Unload frmThatch
End Sub
Public Sub FinishCapture()

        Dim dm As D3DDISPLAYMODE
        Dim pal As PALETTEENTRY
        Dim rct As DxVBLibA.RECT
        
    If ExportCapture = 3 Then
    

        DDevice.GetDisplayMode dm
        rct.Top = 3
        rct.Left = 5
        rct.Right = (frmMain.Picture1.Width / Screen.TwipsPerPixelX)
        rct.Bottom = (frmMain.Picture1.Height / Screen.TwipsPerPixelY)
        D3DX.SaveSurfaceToFile Replace(CommonDialog1.FileName, ".pdf", "1.bmp", , , vbTextCompare), D3DXIFF_BMP, DDevice.GetRenderTarget, pal, rct
        frmMain.Picture1.Picture = LoadPicture(Replace(CommonDialog1.FileName, ".pdf", "1.bmp", , , vbTextCompare))
        SaveJPG frmMain.Picture1.Image, Replace(CommonDialog1.FileName, ".pdf", "1.jpg", , , vbTextCompare), 100
        Kill Replace(CommonDialog1.FileName, ".pdf", "1.bmp", , , vbTextCompare)
        frmMain.Picture1.Picture = LoadPicture("")
        Set TabStrip1.SelectedItem = TabStrip1.Tabs("pattern")
            
        TabStrip1_Click
        
    ElseIf ExportCapture = 6 Then

        DDevice.GetDisplayMode dm
        rct.Top = 3
        rct.Left = 5
        rct.Right = (frmMain.Picture1.Width / Screen.TwipsPerPixelX)
        rct.Bottom = (frmMain.Picture1.Height / Screen.TwipsPerPixelY)
        D3DX.SaveSurfaceToFile Replace(CommonDialog1.FileName, ".pdf", "2.bmp", , , vbTextCompare), D3DXIFF_BMP, DDevice.GetRenderTarget, pal, rct
        frmMain.Picture1.Picture = LoadPicture(Replace(CommonDialog1.FileName, ".pdf", "2.bmp", , , vbTextCompare))
        SaveJPG frmMain.Picture1.Image, Replace(CommonDialog1.FileName, ".pdf", "2.jpg", , , vbTextCompare), 100
        Kill Replace(CommonDialog1.FileName, ".pdf", "2.bmp", , , vbTextCompare)
        frmMain.Picture1.Picture = LoadPicture("")
        Set TabStrip1.SelectedItem = TabStrip1.Tabs("designer")
            
        TabStrip1_Click
                
        ExportCapture = 0
        
        If PathExists(CommonDialog1.FileName, True) Then Kill CommonDialog1.FileName

        Dim pdf As New NTImaging10.PDFCompiler
        pdf.QueueFile Replace(CommonDialog1.FileName, ".pdf", "1.jpg", , , vbTextCompare)
        pdf.QueueFile Replace(CommonDialog1.FileName, ".pdf", "2.jpg", , , vbTextCompare)
        pdf.FitImageTopage = False
        pdf.PageWidth = (((frmMain.Picture1.Width / Screen.TwipsPerPixelX) / PixelPerPoint) * (PixelPerPoint / GetMonitorDPI.Width)) + 4

        pdf.PageHeight = (((frmMain.Picture1.Height / Screen.TwipsPerPixelY) / PixelPerPoint) * (PixelPerPoint / GetMonitorDPI.Height)) + 4

        pdf.MarginBottom = 0
        pdf.MarginLeft = 0
        pdf.MarginTop = 0
        pdf.MarginRight = 0
        pdf.PageFooter = False
        pdf.ChangeQuality = False
        pdf.Exhibit = False
        pdf.CompilePDF CommonDialog1.FileName
        Kill Replace(CommonDialog1.FileName, ".pdf", "1.jpg", , , vbTextCompare)
        Kill Replace(CommonDialog1.FileName, ".pdf", "2.jpg", , , vbTextCompare)
        
        
        
         
    End If
    
    
'        frmMain.Picture1.Height = LastHeight
'        frmMain.Picture1.Width = LastWidth
            
End Sub

Private Sub mnuRemove_Click()

    If MsgBox("Are you sure you want to remove the selected material?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If

    UndoAction = UndoAction & RemoveMaterial(frmStudio.Gallery1.BackgroundColors(Gallery1.ListIndex)) & vbCrLf
    
    UndoEnables
    
    UpdateGallery
    
End Sub

Public Function SymbolToString() As String
    Dim X As Long
    Dim Y As Long
    For X = 0 To (SymbolWidth - 1)
        For Y = 0 To (SymbolHeight - 1)
            SymbolToString = SymbolToString & IIf(SymbolBitmap.Point(X, Y) = vbBlack, 1, 0)
        Next
    Next
End Function

Public Sub StringToSymbol(ByVal txt As String)
    SymbolBitmap.Cls
    Dim X As Long
    Dim Y As Long
    X = 0
    Y = 0
    Do Until txt = ""
        SymbolBitmap.PSet (X, Y), IIf(Left(txt, 1) = "1", vbBlack, vbWhite)
        txt = Mid(txt, 2)
        Y = Y + 1
        If Y = SymbolHeight Then
            Y = 0
            X = X + 1
            If X = SymbolWidth Then
                X = 0
            End If
        End If
    Loop
End Sub

Private Sub mnuAdvanced_Click()
    mnuAdvanced.Checked = Not mnuAdvanced.Checked
    AdvancedGUI
End Sub

Private Sub AdvancedGUI()

    If mnuAdvanced.Checked Then
        DrawStyles.Width = 4200
        DoubleThick.Visible = True
    Else
        DrawStyles.Width = 3315
        DoubleThick.Visible = False
    End If
    ModePanel_Resize
End Sub



Private Sub mnuOpen_Click()

    If CloseProject Then
    
        On Error Resume Next
        CommonDialog1.CancelError = True
        CommonDialog1.DefaultExt = ".csp"
        CommonDialog1.DialogTitle = "Open Cross Stitch Project"
        CommonDialog1.InitDir = CurDir
        CommonDialog1.Filter = "Cross Stitch Project (*.csp)|*.csp|All Files (*.*)|*.*"
        CommonDialog1.FilterIndex = 1
        CommonDialog1.Flags = &H1000 And &H200000 And &H1000
        CommonDialog1.ShowOpen
        If Err Then
            Err.Clear

        Else
            On Error GoTo 0

            TabStrip1.SelectedItem = TabStrip1.Tabs("designer")
        
            TabStrip1_Click
    
            ProjPath = CommonDialog1.FileName
            modProj.ReadFromDisk
            modProj.CleanUpProj
            modProj.CreateProj
            UndoReset
        
        
        End If

    End If
End Sub



Private Sub mnuSave_Click()
    SaveProject
End Sub

Private Sub mnuSaveAs_Click()
    SaveProject True
End Sub

Private Sub UndoCOmmit()
    If UndoAction <> "" Then
        If UndoRecord < UndoBuffer.count And UndoBuffer.count > 0 Then
            Do Until UndoRecord = UndoBuffer.count
                UndoBuffer.Remove UndoBuffer.count
            Loop
        End If
        UndoBuffer.Add UndoAction
        UndoRecord = UndoRecord + 1
        UndoAction = ""
        mnuUndo.Enabled = True
    End If
End Sub
Private Sub SelectColor(ByVal Color As Long)
    Dim cnt As Long
    If Gallery1.count > 0 Then
        For cnt = 0 To Gallery1.count - 1
            If Gallery1.BackgroundColors(cnt) = Color Then
                Gallery1.ListIndex = cnt
                Exit Sub
            End If
        Next
    End If
End Sub
Public Sub mnuUndo_Click()
    UndoCOmmit

    If (UndoRecord > 0 And UndoBuffer.count > 0) Then

        Dim undotext As String
        Dim undoline As String
        
        Dim undoColor As Long
        Dim undobit As Long
        Dim undoX As Long
        Dim undoY As Long
        Dim undoVal As Boolean
        
        Dim changed As Boolean
        
        undotext = UndoBuffer(UndoRecord)
        
        Do Until undotext = ""
            undoline = RemoveNextArg(undotext, vbCrLf)
            If Left(undoline, 1) = "C" Then
                undoline = Mid(undoline, 2)
                
                undoColor = CLng(RemoveNextArg(undoline, ","))
                undobit = CLng(RemoveNextArg(undoline, ","))
                undoX = CLng(RemoveNextArg(undoline, ","))
                undoY = CLng(RemoveNextArg(undoline, ","))
                undoVal = Not CBool(RemoveNextArg(undoline, ","))
            
                CheckStitch(undobit, undoX, undoY, undoColor) = undoVal
            ElseIf Left(undoline, 1) = "M" Then
                undoline = Mid(undoline, 2)
                undoColor = CLng(RemoveNextArg(undoline, ","))
                RemoveMaterial undoColor
                UpdateGallery
            ElseIf Left(undoline, 1) = "R" Then
                undoline = Mid(undoline, 2)
                undoColor = CLng(RemoveNextArg(undoline, ","))
                If CreateColor(undoColor, False, undoline) Then UpdateGallery
                
            ElseIf Left(undoline, 1) = "B" Then
                undoline = Mid(undoline, 2)
                If Not (Gallery1.BackgroundColors(Gallery1.ListIndex) = CLng(NextArg(undoline, ","))) Then
                    SelectColor CLng(NextArg(undoline, ","))
                    Set TabStrip1.SelectedItem = TabStrip1.Tabs("symbols")
                    Gallery1_Click
                End If
                
                SymbolBitmap.PSet (CLng(NextArg(RemoveArg(undoline, ","), ",")), CLng(RemoveArg(RemoveArg(undoline, ","), ","))), vbWhite
                'HelperImage.PSet (Screen.TwipsPerPixelX * (CLng(NextArg(RemoveArg(undoline, ","), ",")) - 1), Screen.TwipsPerPixelY * (CLng(RemoveArg(RemoveArg(undoline, ","), ",")) - 1)), vbWhite
                changed = True

            ElseIf Left(undoline, 1) = "W" Then
                undoline = Mid(undoline, 2)
                If Not (Gallery1.BackgroundColors(Gallery1.ListIndex) = CLng(NextArg(undoline, ","))) Then
                    SelectColor CLng(NextArg(undoline, ","))
                    Set TabStrip1.SelectedItem = TabStrip1.Tabs("symbols")
                    Gallery1_Click
                End If
                SymbolBitmap.PSet (CLng(NextArg(RemoveArg(undoline, ","), ",")), CLng(RemoveArg(RemoveArg(undoline, ","), ","))), vbBlack
                'HelperImage.PSet (Screen.TwipsPerPixelX * (CLng(NextArg(RemoveArg(undoline, ","), ",")) - 1), Screen.TwipsPerPixelY * (CLng(RemoveArg(RemoveArg(undoline, ","), ",")) - 1)), vbBlack
                changed = True

            End If
            
        Loop
        
        If changed Then
                SymbolEdit_Resize
                UpdateSymbol
        End If

    
        UndoRecord = UndoRecord - 1
    End If
    UndoEnables
End Sub

Public Sub UndoReset()
    UndoRecord = 0
    UndoBuffer.Clear
    UndoAction = ""
    UndoEnables
End Sub

Public Function GetSymbol(ByVal Color As Long) As String
    Dim rs As New ADODB.Recordset
    db.rsQuery rs, "SELECT * FROM Materials WHERE Color=" & Color & ";"
    If Not db.rsEnd(rs) Then
        GetSymbol = rs("Symbol")
    Else
        GetSymbol = String(SymbolWidth * SymbolHeight, "0")
    End If
    db.rsClose rs

End Function

Public Sub SetSymbol(ByVal Color As Long, ByVal Symbol As String)

    db.dbQuery "UPDATE Materials SET Symbol='" & Symbol & "' WHERE Color=" & Color & ";"
    
    Dim rs As New ADODB.Recordset
    db.rsQuery rs, "SELECT * FROM Materials WHERE Color=" & Color & ";"
    If Not db.rsEnd(rs) Then
        
        StringToSymbol Symbol

        SavePicture SymbolBitmap.Image, GetSymbolFile(rs("ID"))
    End If
    db.rsClose rs
    
End Sub

Public Sub mnuRedo_Click()

    If (UndoRecord < UndoBuffer.count And UndoBuffer.count > 0) And (UndoAction = "") Then
    
        UndoRecord = UndoRecord + 1
        
        Dim undotext As String
        Dim undoline As String
        
        Dim undoColor As Long
        Dim undobit As Long
        Dim undoX As Long
        Dim undoY As Long
        Dim undoVal As Boolean
        
        Dim changed As Boolean
        
        undotext = UndoBuffer(UndoRecord)
        Do Until undotext = ""
            undoline = RemoveNextArg(undotext, vbCrLf)
            If Left(undoline, 1) = "C" Then
                undoline = Mid(undoline, 2)
                undoColor = CLng(RemoveNextArg(undoline, ","))
                undobit = CLng(RemoveNextArg(undoline, ","))
                undoX = CLng(RemoveNextArg(undoline, ","))
                undoY = CLng(RemoveNextArg(undoline, ","))
                undoVal = CBool(RemoveNextArg(undoline, ","))

                CheckStitch(undobit, undoX, undoY, undoColor) = undoVal
            ElseIf Left(undoline, 1) = "M" Then
                undoline = Mid(undoline, 2)
                undoColor = CLng(RemoveNextArg(undoline, ","))
                If CreateColor(undoColor, False, undoline) Then UpdateGallery
            ElseIf Left(undoline, 1) = "R" Then
                undoline = Mid(undoline, 2)
                undoColor = CLng(RemoveNextArg(undoline, ","))
                RemoveMaterial undoColor
                UpdateGallery
            ElseIf Left(undoline, 1) = "B" Then
                undoline = Mid(undoline, 2)
                If Not (Gallery1.BackgroundColors(Gallery1.ListIndex) = CLng(NextArg(undoline, ","))) Then
                    SelectColor CLng(NextArg(undoline, ","))
                    Set TabStrip1.SelectedItem = TabStrip1.Tabs("symbols")
                    Gallery1_Click
                End If
                SymbolBitmap.PSet (CLng(NextArg(RemoveArg(undoline, ","), ",")), CLng(RemoveArg(RemoveArg(undoline, ","), ","))), vbBlack
                'HelperImage.PSet (Screen.TwipsPerPixelX * (CLng(NextArg(RemoveArg(undoline, ","), ",")) - 1), Screen.TwipsPerPixelY * (CLng(RemoveArg(RemoveArg(undoline, ","), ",")) - 1)), vbBlack
                changed = True
            ElseIf Left(undoline, 1) = "W" Then
                undoline = Mid(undoline, 2)
                If Not (Gallery1.BackgroundColors(Gallery1.ListIndex) = CLng(NextArg(undoline, ","))) Then
                    SelectColor CLng(NextArg(undoline, ","))
                    Set TabStrip1.SelectedItem = TabStrip1.Tabs("symbols")
                    Gallery1_Click
                End If
                SymbolBitmap.PSet (CLng(NextArg(RemoveArg(undoline, ","), ",")), CLng(RemoveArg(RemoveArg(undoline, ","), ","))), vbWhite
                'HelperImage.PSet (Screen.TwipsPerPixelX * (CLng(NextArg(RemoveArg(undoline, ","), ",")) - 1), Screen.TwipsPerPixelY * (CLng(RemoveArg(RemoveArg(undoline, ","), ",")) - 1)), vbWhite
                changed = True
                
            End If
            
        Loop
        If changed Then
        SymbolEdit_Resize
                UpdateSymbol
        End If

        mnuRedo.Enabled = (UndoRecord < UndoBuffer.count And UndoBuffer.count > 0) And (UndoAction = "")
    End If

End Sub

Public Sub UndoEnables()
    mnuUndo.Enabled = (UndoRecord > 0 And UndoBuffer.count > 0) Or (UndoAction <> "")
    mnuRedo.Enabled = (UndoRecord <= UndoBuffer.count And UndoBuffer.count > 0) And (UndoAction = "")
End Sub



Public Sub Precision_Click()
    If CancelUpdate = False Then
        CancelUpdate = True
        RemovalTool.Value = 0
        Precision.Value = 1
        DrawSytleUIEnable
        CancelUpdate = False
    End If
End Sub

Public Sub RemovalTool_Click()
    If CancelUpdate = False Then
        CancelUpdate = True
        Precision.Value = 0

        RemovalTool.Value = 1
        DrawSytleUIEnable
        CancelUpdate = False
    End If
End Sub

Private Property Get GridWidth() As Single
    GridWidth = Round(SymbolEdit.ScaleWidth / SymbolWidth, 2)
End Property
Private Property Get GridHeight() As Single
    GridHeight = Round(SymbolEdit.ScaleHeight / SymbolHeight, 2)
End Property
Private Sub PaintSymbol()
    If SymbolEdit.Visible Then
    
        Dim BlockX As Single
        Dim BlockY As Single
                
        For BlockX = SymbolEdit.ScaleWidth - (Screen.TwipsPerPixelX * 2) To 0 Step -GridWidth
            SymbolEdit.Line (BlockX, SymbolEdit.ScaleHeight)-(BlockX, 0), &H404040
        Next
        For BlockY = SymbolEdit.ScaleHeight - (Screen.TwipsPerPixelY * 2) To 0 Step -GridHeight
            SymbolEdit.Line (0, BlockY)-(SymbolEdit.ScaleWidth, BlockY), &H404040
        Next
       
        For BlockX = 0 To SymbolWidth - 1
            For BlockY = 0 To SymbolHeight - 1

                If SymbolBitmap.Point(BlockX, BlockY) = vbBlack Then
                    SymbolEdit.Line (((BlockX * GridWidth)), ((BlockY * GridHeight)))-(((BlockX + 1) * GridWidth) - (Screen.TwipsPerPixelX * 2), ((BlockY + 1) * GridHeight) - (Screen.TwipsPerPixelY * 2)), vbBlack, BF
                Else
                    SymbolEdit.Line (((BlockX * GridWidth)), ((BlockY * GridHeight)))-(((BlockX + 1) * GridWidth) - (Screen.TwipsPerPixelX * 2), ((BlockY + 1) * GridHeight) - (Screen.TwipsPerPixelY * 2)), vbWhite, BF
                End If
            Next
        Next

    End If
End Sub

Private Sub SymbolEdit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim BlockX As Long
    Dim BlockY As Long
    BlockX = (X \ GridWidth)
    BlockY = (Y \ GridHeight)
        
    If Button = 1 Then
        UpdateTimer.Enabled = False

        Debug.Print BlockX; BlockY
        If SymbolBitmap.Point(BlockX, BlockY) = vbBlack Then
            SymbolBitmap.PSet (BlockX, BlockY), vbWhite
            
            SymbolEdit.Line (((BlockX * GridWidth)), ((BlockY * GridHeight)))-(((BlockX + 1) * GridWidth) - (Screen.TwipsPerPixelX * 2), ((BlockY + 1) * GridHeight) - (Screen.TwipsPerPixelY * 2)), vbWhite, BF
            
            'HelperImage.PSet (Screen.TwipsPerPixelX * (BlockX - 1), Screen.TwipsPerPixelY * (BlockY - 1)), vbWhite
            SymbolPen = False
            If InStr(UndoAction, "W" & Gallery1.BackgroundColors(Gallery1.ListIndex) & "," & BlockX & "," & BlockY & vbCrLf) = 0 Then
                UndoAction = UndoAction & "W" & Gallery1.BackgroundColors(Gallery1.ListIndex) & "," & BlockX & "," & BlockY & vbCrLf
            End If
        Else
            SymbolBitmap.PSet (BlockX, BlockY), vbBlack
            SymbolEdit.Line (((BlockX * GridWidth)), ((BlockY * GridHeight)))-(((BlockX + 1) * GridWidth) - (Screen.TwipsPerPixelX * 2), ((BlockY + 1) * GridHeight) - (Screen.TwipsPerPixelY * 2)), vbBlack, BF
            
            'HelperImage.PSet (Screen.TwipsPerPixelX * (BlockX - 1), Screen.TwipsPerPixelY * (BlockY - 1)), vbBlack
            SymbolPen = True
            If InStr(UndoAction, "B" & Gallery1.BackgroundColors(Gallery1.ListIndex) & "," & BlockX & "," & BlockY & vbCrLf) = 0 Then
                UndoAction = UndoAction & "B" & Gallery1.BackgroundColors(Gallery1.ListIndex) & "," & BlockX & "," & BlockY & vbCrLf
            End If
        End If

        UpdateTimer.Enabled = True
       ' SymbolEdit_Resize
    Else
        SymbolPen = (SymbolBitmap.Point(BlockX, BlockY) = vbBlack)
    End If
End Sub

Private Sub SymbolEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim BlockX As Long
    Dim BlockY As Long
    BlockX = (X \ GridWidth)
    BlockY = (Y \ GridHeight)
        
    If Button = 1 Then
        UpdateTimer.Enabled = False
        
        If Not SymbolPen Then
        
            SymbolBitmap.PSet (BlockX, BlockY), vbWhite
            SymbolEdit.Line (((BlockX * GridWidth)), ((BlockY * GridHeight)))-(((BlockX + 1) * GridWidth) - (Screen.TwipsPerPixelX * 2), ((BlockY + 1) * GridHeight) - (Screen.TwipsPerPixelY * 2)), vbWhite, BF

            'HelperImage.PSet (Screen.TwipsPerPixelX * (BlockX - 1), Screen.TwipsPerPixelY * (BlockY - 1)), vbWhite
            If InStr(UndoAction, "W" & Gallery1.BackgroundColors(Gallery1.ListIndex) & "," & BlockX & "," & BlockY & vbCrLf) = 0 Then
                UndoAction = UndoAction & "W" & Gallery1.BackgroundColors(Gallery1.ListIndex) & "," & BlockX & "," & BlockY & vbCrLf
            End If
        Else
            SymbolBitmap.PSet (BlockX, BlockY), vbBlack
            SymbolEdit.Line (((BlockX * GridWidth)), ((BlockY * GridHeight)))-(((BlockX + 1) * GridWidth) - (Screen.TwipsPerPixelX * 2), ((BlockY + 1) * GridHeight) - (Screen.TwipsPerPixelY * 2)), vbBlack, BF

            'HelperImage.PSet (Screen.TwipsPerPixelX * (BlockX - 1), Screen.TwipsPerPixelY * (BlockY - 1)), vbBlack
            If InStr(UndoAction, "B" & Gallery1.BackgroundColors(Gallery1.ListIndex) & "," & BlockX & "," & BlockY & vbCrLf) = 0 Then
                UndoAction = UndoAction & "B" & Gallery1.BackgroundColors(Gallery1.ListIndex) & "," & BlockX & "," & BlockY & vbCrLf
            End If
        End If

        UpdateTimer.Enabled = True
       ' SymbolEdit_Resize
    Else
        SymbolPen = (SymbolBitmap.Point(BlockX, BlockY) = vbBlack)
    End If
End Sub

Private Sub SymbolEdit_Resize()
    SymbolEdit.Cls
    PaintSymbol
    
End Sub

Private Sub Symbols_Resize()

    SymbolEdit.Width = Symbols.ScaleWidth
    SymbolEdit.Left = 0
    SymbolEdit.Top = 0
    SymbolEdit.Height = Symbols.ScaleHeight
    PaintSymbol
End Sub

Public Sub AdjustImportGallery(ByRef Owner As PictureBox, ByRef Import As Import, ByRef Gallery As Gallery)
    Gallery.Left = 0
    Gallery.Top = 0
    If Owner.Width - Import.Width > 0 Then
        Gallery.Width = Owner.Width - Import.Width
    End If
    Import.Left = Gallery.Width
    Gallery.Top = 0
    Gallery.Height = Import.Height
    
End Sub
Private Sub mnuAbout_Click()
    frmSplash.ShowAbout
End Sub

Private Sub mnuClose_Click()

    If modProj.Dirty Then
        If MsgBox("Save changes before closing this project?", vbYesNoCancel) Then
        End If
    End If

End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Function ThatchProperties() As Boolean
    frmThatch.SetupThatch
    frmThatch.Text3.Text = modProj.ThatchBlock
    frmThatch.Text1.Text = modProj.ThatchYUnits
    frmThatch.Text2.Text = modProj.ThatchXUnits

    frmThatch.Show 1

    If frmThatch.Tag = "OK" Then

        modProj.ThatchYUnits = CSng(frmThatch.Text1.Text)
        modProj.ThatchXUnits = CSng(frmThatch.Text2.Text)
        modProj.ThatchBlock = CSng(frmThatch.Text3.Text)

        modProj.ThatchIndex = frmThatch.Gallery1.ListIndex + 1
        ResizeMultidimArray
        GotoCenter

        
        ThatchProperties = True
    End If
    Unload frmThatch
End Function

Private Sub mnuNew_Click()

    
    
    If ThatchProperties() Then
    
        CloseProject
        Dim X As Long
        Dim Y As Long
        Dim i As Byte
        For X = 1 To ThatchXUnits
            For Y = 1 To ThatchYUnits
                If ProjGrid(X, Y).count > 0 Then
                    Erase ProjGrid(X, Y).Details
                    ProjGrid(X, Y).count = 0
                End If
            Next
        Next
        
        UndoReset
        
    End If
    
    

End Sub

Private Sub mnuOnScreenHelp_Click()
    Help.Visible = Not Help.Visible
    mnuOnScreenHelp.Checked = Help.Visible
    
End Sub

Private Sub ModePanel_Resize()
    
    DrawStyles.Left = ModePanel.Width - DrawStyles.Width
    SymbolButtons.Left = ModePanel.Width - SymbolButtons.Width
    
End Sub

Private Sub TabStrip1_Click()
    
    If UpdateTimer.Enabled Then
        UpdateTimer.Enabled = False
        UpdateTimer_Timer
    End If
    
    DrawStyles.Visible = TabStrip1.SelectedItem.Key = "designer"

    SymbolButtons.Visible = TabStrip1.SelectedItem.Key = "symbols"
    
    mnuCut.Enabled = TabStrip1.SelectedItem.Key = "symbols"
    mnuCopy.Enabled = TabStrip1.SelectedItem.Key = "symbols"
    mnuPaste.Enabled = TabStrip1.SelectedItem.Key = "symbols"
    mnuClear.Enabled = TabStrip1.SelectedItem.Key = "symbols"


    Symbols.Visible = (TabStrip1.SelectedItem.Key = "symbols") And (Gallery1.count > 0)
    Designer.Visible = TabStrip1.SelectedItem.Key <> "symbols"


    modProj.CleanUpProj
    modProj.CreateProj

    Gallery1_Click


End Sub
Private Sub UpdateSymbol()
    If PathExists(GetSymbolFile, True) Then

        SavePicture SymbolBitmap.Image, GetSymbolFile
'        Debug.Print Gallery1.Info(Gallery1.ListIndex)
        db.dbQuery "UPDATE Materials SET Symbol='" & SymbolToString & "' WHERE ID=" & Gallery1.Info(Gallery1.ListIndex) & ";"
        
        Set Gallery1.Images(Gallery1.ListIndex) = SymbolBitmap.Image
        
    End If
End Sub
Private Sub UpdateTimer_Timer()
    UpdateTimer.Enabled = False
   
    
    
    If SymbolEdit.Visible Then

        UpdateSymbol


    End If
    
    UndoCOmmit
End Sub

Private Sub VScroll2_Change()
    Line1.Visible = (VScroll2.Value = -1)
    Line2.Visible = (VScroll2.Value = -2)
    Line3.Visible = (VScroll2.Value = -3) Or (VScroll2.Value = -5)
    Line4.Visible = (VScroll2.Value = -4) Or (VScroll2.Value = -5)
    
End Sub

Private Sub VScroll2_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyDown KeyCode, Shift
End Sub

Private Sub VScroll2_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub
