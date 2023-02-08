VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9465
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   23415
   _ExtentX        =   41301
   _ExtentY        =   16695
   _Version        =   393216
   Description     =   "VB 6 Neotext Basic - Enhancements for Visual Basic 6.0"
   DisplayName     =   "VB 6 Neotext Basic"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

Implements IConnect

Public Enum UIStates
    None = -1
    Design = 0
    Run = 1
    Break = 2
    Running = 4
End Enum

Private VBInstance As VBIDE.VBE
Private VBWindow As VBIDE.Window

Public CmdBar As CommandBar 'compiler toolbar

'##################Tool button Objects and events either new or existing built in + some built mocked of built in)

Private cmdButton1 As CommandBarButton 'sign
Private WithEvents cmdMenuButton1 As CommandBarEvents
Attribute cmdMenuButton1.VB_VarHelpID = -1

Private cmdButton2 As CommandBarButton 'start with full compile
Private WithEvents cmdMenuButton2 As CommandBarEvents
Attribute cmdMenuButton2.VB_VarHelpID = -1

Private cmdButton3 As CommandBarButton 'remake executable
Private WithEvents cmdMenuButton3 As CommandBarEvents
Attribute cmdMenuButton3.VB_VarHelpID = -1

Private cmdButton4 As CommandBarButton 'remake release
Private WithEvents cmdMenuButton4 As CommandBarEvents
Attribute cmdMenuButton4.VB_VarHelpID = -1

Private cmdButton5 As CommandBarButton 'start the executable
Private WithEvents cmdMenuButton5 As CommandBarEvents
Attribute cmdMenuButton5.VB_VarHelpID = -1

Private cmdButton6 As CommandBarButton 'make...
Private WithEvents cmdMenuButton6 As CommandBarEvents
Attribute cmdMenuButton6.VB_VarHelpID = -1

Private cmdButton7 As CommandBarButton 'make project group
Private WithEvents cmdMenuButton7 As CommandBarEvents
Attribute cmdMenuButton7.VB_VarHelpID = -1

Private cmdButton8 As CommandBarButton 'stop the executable
Private WithEvents cmdMenuButton8 As CommandBarEvents
Attribute cmdMenuButton8.VB_VarHelpID = -1

'##################Menu Objects and events (either new or existing built in + some built mocked of built in)


Private cbMenuCommandBar1 As Office.CommandBarControl 'start the executable
Private WithEvents MenuHandler4 As CommandBarEvents
Attribute MenuHandler4.VB_VarHelpID = -1

Private cbMenuCommandBar2 As Office.CommandBarControl 'build project release
Private WithEvents MenuHandler3 As CommandBarEvents
Attribute MenuHandler3.VB_VarHelpID = -1

Private cbMenuCommandBar3 As Office.CommandBarControl 'remake project build
Private WithEvents MenuHandler5 As CommandBarEvents
Attribute MenuHandler5.VB_VarHelpID = -1

Private cbMenuCommandBar4 As Office.CommandBarControl 'make...
Private WithEvents MenuHandler2 As CommandBarEvents
Attribute MenuHandler2.VB_VarHelpID = -1

Private cbMenuCommandBar6 As Office.CommandBarControl 'properties...
Private WithEvents MenuHandler9 As CommandBarEvents
Attribute MenuHandler9.VB_VarHelpID = -1

Private cbMenuCommandBar5 As Office.CommandBarControl
Private WithEvents MenuHandler8 As CommandBarEvents
Attribute MenuHandler8.VB_VarHelpID = -1

Private cbMenuCommandBar7 As Office.CommandBarControl 'Procedure Attributes...
Private WithEvents MenuHandler10 As CommandBarEvents
Attribute MenuHandler10.VB_VarHelpID = -1

'#########Single event objects below from existing built-ins
Private WithEvents cmdBarBtnEvents1 As CommandBarEvents 'start
Attribute cmdBarBtnEvents1.VB_VarHelpID = -1
Private WithEvents cmdBarBtnEvents2 As CommandBarEvents 'break
Attribute cmdBarBtnEvents2.VB_VarHelpID = -1
Private WithEvents cmdBarBtnEvents3 As CommandBarEvents 'end
Attribute cmdBarBtnEvents3.VB_VarHelpID = -1
Private WithEvents MenuHandler1 As CommandBarEvents 'start with full compile
Attribute MenuHandler1.VB_VarHelpID = -1
Private WithEvents MenuHandler6 As CommandBarEvents 'make project group
Attribute MenuHandler6.VB_VarHelpID = -1


'#########Mock created barbuttons event objects below to catch & place existing built-ins


Private cmdBarBtn1 As CommandBarButton 'Start
Private WithEvents cmdBarBtnEvents4 As CommandBarEvents
Attribute cmdBarBtnEvents4.VB_VarHelpID = -1

Private cmdBarBtn2 As CommandBarButton 'Break
Private WithEvents cmdBarBtnEvents5 As CommandBarEvents
Attribute cmdBarBtnEvents5.VB_VarHelpID = -1

Private cmdBarBtn3 As CommandBarButton 'End
Private WithEvents cmdBarBtnEvents6 As CommandBarEvents
Attribute cmdBarBtnEvents6.VB_VarHelpID = -1

Private cmdBarBtn4 As CommandBarButton 'options
Private WithEvents cmdBarBtnEvents7 As CommandBarEvents
Attribute cmdBarBtnEvents7.VB_VarHelpID = -1

Private cmdBarBtn5 As CommandBarButton 'sign
Private WithEvents MenuHandler7 As CommandBarEvents
Attribute MenuHandler7.VB_VarHelpID = -1

Private cmdBarBtn8 As CommandBarButton 'Start
Private WithEvents cmdBarBtnEvents8 As CommandBarEvents
Attribute cmdBarBtnEvents8.VB_VarHelpID = -1

Private cmdBarBtn9 As CommandBarButton 'Break
Private WithEvents cmdBarBtnEvents9 As CommandBarEvents
Attribute cmdBarBtnEvents9.VB_VarHelpID = -1

Private cmdBarBtn10 As CommandBarButton 'End
Private WithEvents cmdBarBtnEvents10 As CommandBarEvents
Attribute cmdBarBtnEvents10.VB_VarHelpID = -1

Private cmdBarBtn6 As CommandBarButton 'sign
Private cmdBarBtn7 As CommandBarButton 'sign

Private docSettings As Settings

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Const guidPos = "{99179999-A05E-4A21-A9DB-C5614C53F992}"

Public MSWHeelObject As VBIDE.AddIn

Public WithEvents FCE As FileControlEvents
Attribute FCE.VB_VarHelpID = -1
'Public VBInstance As VBIDE.VBE


Private Sub FCE_AfterWriteFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String, ByVal Result As Integer)
    On Error Resume Next
    On Local Error Resume Next
    If CLng(GetSetting("BasicNeotext", "Options", "ProcedureDesc", 0)) = 1 Then
        If PathExists(FileName, True) Then
            'BuildComments CommentsToAttribute, GetCodeModule(VBInstance.VBProjects, VBProject.Name, GetModuleName(FileName))
            BuildFileDescriptions FileName, False
        End If
    End If
    If GetFileExt(FileName, True, True) = "vbp" Then
        BuildProject FileName
    End If

End Sub

Private Sub FCE_BeforeLoadFile(ByVal VBProject As VBIDE.VBProject, FileNames() As String)
    On Error Resume Next
    On Local Error Resume Next
    Dim cnt As Long
    
    For cnt = LBound(FileNames) To UBound(FileNames)
        If PathExists(FileNames(cnt), True) Then
            If GetFileExt(FileNames(cnt), True, True) = "vbp" Then
                BuildProject FileNames(cnt)
               ' Stop
            End If
            If CLng(GetSetting("BasicNeotext", "Options", "ProcedureDesc", 0)) = 1 Then
                BuildFileDescriptions FileNames(cnt), True
            End If
            'BuildComments AttributeToComments, GetCodeModule(VBInstance.VBProjects, VBProject.Name, GetModuleName(FileNames(cnt)))
        End If
    Next

End Sub

Private Function GetUIState() As UIStates
    If InStr(VBInstance.MainWindow.Caption, "Microsoft Visual Basic [running]") > 0 Then
        GetUIState = UIStates.Running
    ElseIf InStr(VBInstance.MainWindow.Caption, "Microsoft Visual Basic [break]") > 0 Then
        GetUIState = UIStates.Break
    ElseIf InStr(VBInstance.MainWindow.Caption, "Microsoft Visual Basic [run]") > 0 Then
        GetUIState = UIStates.Run
    Else
        GetUIState = UIStates.Design
    End If
End Function

Public Sub AdjustHeaders()
    'MsgBox "up"
End Sub
Public Sub UnAdjustHeaders()
   ' MsgBox "out"
End Sub

Public Sub SetUIState(Optional ByVal NewState As UIStates = UIStates.None)
    
    If NewState = None Then NewState = GetUIState
    If NewState = Design Then UnAdjustHeaders
    IterateCommandObj VBInstance.CommandBars, NewState
End Sub

Private Function IterateCommandObj(ByVal cmdObj As Object, ByVal UIState As UIStates) As Boolean
    Dim obj As Object
    Select Case TypeName(cmdObj)
        Case "CommandBar", "CommandBarControl"
        Case "CommandBarButton"
            IterateCommandObj = CheckCommandObj(cmdObj, UIState)
        Case "CommandBarPopup"
            For Each obj In cmdObj.Controls
                IterateCommandObj = IterateCommandObj(obj, UIState)
            Next
        Case "CommandBars"
            For Each obj In cmdObj
                IterateCommandObj = IterateCommandObj(obj.Controls, UIState)
            Next
        Case "CommandBarControls"
            For Each obj In cmdObj
                IterateCommandObj = IterateCommandObj(obj, UIState)
            Next
    End Select
End Function
Private Function CheckCommandObj(ByVal cmdObj As Object, ByVal UIState As UIStates) As Boolean
    Dim val As Boolean
    Dim chkval As Variant
    
    If TypeName(cmdObj) = "CommandBarButton" Then
        chkval = cmdObj.faceid
    Else
        chkval = cmdObj.Caption
    End If
    
    Select Case chkval
'        Case "&Start", 186
'            val = (((UIState And UIStates.Design) = UIStates.Design) Or ((UIState And UIStates.Break) = UIStates.Break))
'            If Not (cmdObj.Enabled = val) Then
'                cmdObj.Enabled = val
'                CheckCommandObj = (cmdObj.Enabled <> val)
'            End If
'        Case "Brea&k", 189
'            val = (((UIState And UIStates.Run) = UIStates.Run) Or ((UIState And UIStates.Running) = UIStates.Running))
'            If Not (cmdObj.Enabled = val) Then
'                cmdObj.Enabled = val
'                CheckCommandObj = (cmdObj.Enabled <> val)
'            End If
'        Case "&End", 228
'            val = (((UIState And UIStates.Run) = UIStates.Run) Or ((UIState And UIStates.Break) = UIStates.Break) Or ((UIState And UIStates.Running) = UIStates.Running))
'            If Not (cmdObj.Enabled = val) Then
'                cmdObj.Enabled = val
'                CheckCommandObj = (cmdObj.Enabled <> val)
'            End If
        Case "Sign", 30
            If Not (VBInstance.ActiveVBProject Is Nothing) Then
                val = (((UIState And UIStates.Design) = UIStates.Design) And PathExists(VBInstance.ActiveVBProject.BuildFileName, True))
            Else
                val = ((UIState And UIStates.Design) = UIStates.Design)
            End If
            If Not (cmdObj.Enabled = val) Then
                cmdObj.Enabled = val
                CheckCommandObj = (cmdObj.Enabled <> val)
            End If
        Case "Start &The Executable", 459
            If Not (VBInstance.ActiveVBProject Is Nothing) Then
                val = (((UIState And UIStates.Design) = UIStates.Design) And PathExists(VBInstance.ActiveVBProject.BuildFileName, True))
            Else
                val = ((UIState And UIStates.Design) = UIStates.Design)
            End If
            If Not (cmdObj.Enabled = val) Then
                cmdObj.Enabled = val
                CheckCommandObj = (cmdObj.Enabled <> val)
            End If
        Case "Start With &Full Compile", 539
            val = ((UIState And UIStates.Design) = UIStates.Design)
            If Not (cmdObj.Enabled = val) Then
                cmdObj.Enabled = val
                CheckCommandObj = (cmdObj.Enabled <> val)
            End If
        Case "Remake Pro&ject Build", 37
            If Not (VBInstance.ActiveVBProject Is Nothing) Then
                val = (((UIState And UIStates.Design) = UIStates.Design) And PathExists(VBInstance.ActiveVBProject.BuildFileName, True))
             Else
                val = ((UIState And UIStates.Design) = UIStates.Design)
            End If
            If Not (cmdObj.Enabled = val) Then
                cmdObj.Enabled = val
                CheckCommandObj = (cmdObj.Enabled <> val)
            End If
        Case "Make...", 215
            val = ((UIState And UIStates.Design) = UIStates.Design)
            If Not (cmdObj.Enabled = val) Then
                cmdObj.Enabled = val
                CheckCommandObj = (cmdObj.Enabled <> val)
            End If
        Case "Make Project &Group..."
            val = ((UIState And UIStates.Design) = UIStates.Design)
            If Not (cmdObj.Enabled = val) Then
                cmdObj.Enabled = val
                CheckCommandObj = (cmdObj.Enabled <> val)
            End If
        Case Else
    End Select
    
'''    Select Case cmdObj.Parent.Name
'''        Case "Run", "Compiler"
'''            Exit Function
'''    End Select
'''    Select Case cmdObj.faceid
'''        Case 186, 189, 228
'''            If Not (cmdObj.Visible = False) Then cmdObj.Visible = False
'''    End Select

End Function


Private Sub AddinInstance_OnAddInsUpdate(custom() As Variant)
   
    SetUIState

    
End Sub

Private Sub AddinInstance_OnBeginShutdown(custom() As Variant)
    On Error GoTo exitthis
    On Local Error GoTo exitthis

    MSVBRedraw True
    
'    Dim cP As Window
'    For Each cP In VBInstance.Windows
'
'        Select Case StrReverse(NextArg(StrReverse(cP.Caption), " "))
'            Case "(UserControl)", "(Form)", "(UserDocument)", "(AddInDesigner)", "(Code)"
'                cP.Close
'        End Select
'    Next
    
exitthis:
    If Err Then Err.Clear
    On Error GoTo 0
    On Local Error GoTo 0
End Sub

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error Resume Next
    On Local Error Resume Next
    
    Set VBInstance = Application

    Set CmdBar = VBInstance.CommandBars.Add("Compiler", msoBarTop)
    
    CmdBar.Position = GetSetting("BasicNeotext", "Options", "ToolBar_Position", CmdBar.Position)
    CmdBar.RowIndex = GetSetting("BasicNeotext", "Options", "ToolBar_RowIndex", CmdBar.RowIndex)
    CmdBar.Left = GetSetting("BasicNeotext", "Options", "ToolBar_Left", CmdBar.Left)
    CmdBar.Top = GetSetting("BasicNeotext", "Options", "ToolBar_Top", CmdBar.Top)
    CmdBar.Visible = GetSetting("BasicNeotext", "Options", "ToolBar_Visible", True)
    
    Set cmdButton1 = CmdBar.Controls.Add(msoControlButton)
    cmdButton1.Caption = "Sign"
    cmdButton1.ToolTipText = "Code Sign Executable"
    cmdButton1.Style = msoButtonIcon
    cmdButton1.faceid = 30 '20
    Set cmdBarBtn4 = CmdBar.Controls.Add(msoControlButton)
    cmdBarBtn4.Caption = "Options"
    cmdBarBtn4.ToolTipText = "Options"
    cmdBarBtn4.Style = msoButtonIcon
    cmdBarBtn4.faceid = 162
'    Set cmdBarBtn5 = CmdBar.Controls.Add(msoControlButton)
'    cmdBarBtn5.Caption = "Sign"
'    cmdBarBtn5.ToolTipText = "Sign Executable"
'    cmdBarBtn5.Style = msoButtonIcon
'    cmdBarBtn5.faceid = 30 '20
    Set cmdButton6 = CmdBar.Controls.Add(msoControlButton)
    cmdButton6.BeginGroup = True
    cmdButton6.Caption = "Remake Pro&ject Build"
    cmdButton6.ToolTipText = "Remake Project Executable"
    cmdButton6.Style = msoButtonIcon
    cmdButton6.faceid = 37
    Set cmdButton5 = CmdBar.Controls.Add(msoControlButton)
    cmdButton5.Caption = "Start &The Executable"
    cmdButton5.ToolTipText = "Start the Executable Only"
    cmdButton5.Style = msoButtonIcon
    cmdButton5.faceid = 459
    Set cmdButton3 = CmdBar.Controls.Add(msoControlButton)
    cmdButton3.Caption = "Make..."
    cmdButton3.ToolTipText = "Make Project Dialog"
    cmdButton3.Style = msoButtonIcon
    cmdButton3.faceid = 215
    Set cmdButton2 = CmdBar.Controls.Add(msoControlButton)
    cmdButton2.Caption = "Start With &Full Compile"
    cmdButton2.ToolTipText = "Start With Full Compile"
    cmdButton2.Style = msoButtonIcon
    cmdButton2.faceid = 539
    Set cmdButton8 = CmdBar.Controls.Add(msoControlButton)
    cmdButton8.Caption = "Stop the E&xecutable"
    cmdButton8.ToolTipText = "Stop the E&xecutable"
    cmdButton8.Style = msoButtonIcon
    cmdButton8.faceid = 348
        
'    Set cmdBarBtn1 = VBInstance.CommandBars("Run").Controls("&Start").Copy(CmdBar)
'    cmdBarBtn1.BeginGroup = True
'    Set cmdBarBtn2 = VBInstance.CommandBars("Run").Controls("Brea&k").Copy(CmdBar)
'    Set cmdBarBtn3 = VBInstance.CommandBars("Run").Controls("&End").Copy(CmdBar)

'    Set cmdButton4 = CmdBar.Controls.Add(msoControlButton)
'    cmdButton4.BeginGroup = True
'    cmdButton4.Caption = "&Build Project Release"
'    cmdButton4.ToolTipText = "Build With Release Options"
'    cmdButton4.Style = msoButtonIcon
'    cmdButton4.faceid = 184
'
'    Set cmdButton7 = CmdBar.Controls.Add(msoControlButton)
'    cmdButton7.Caption = "Make Project Group..."
'    cmdButton7.ToolTipText = "Make Project Group Dialog"
'    cmdButton7.Style = msoButtonIcon
'    cmdButton7.faceid = 185
    
'    Set cmdBarBtnEvents4 = VBInstance.Events.CommandBarEvents(cmdBarBtn1)
'    Set cmdBarBtnEvents5 = VBInstance.Events.CommandBarEvents(cmdBarBtn2)
'    Set cmdBarBtnEvents6 = VBInstance.Events.CommandBarEvents(cmdBarBtn3)
    Set cmdBarBtnEvents7 = VBInstance.Events.CommandBarEvents(cmdBarBtn4)



'    Set MenuHandler7 = VBInstance.Events.CommandBarEvents(cmdBarBtn5)

    Set cmdMenuButton1 = VBInstance.Events.CommandBarEvents(cmdButton1)
    Set cmdMenuButton2 = VBInstance.Events.CommandBarEvents(cmdButton2)
    Set cmdMenuButton3 = VBInstance.Events.CommandBarEvents(cmdButton3)
'    Set cmdMenuButton4 = VBInstance.Events.CommandBarEvents(cmdButton4)
    Set cmdMenuButton5 = VBInstance.Events.CommandBarEvents(cmdButton5)
    Set cmdMenuButton6 = VBInstance.Events.CommandBarEvents(cmdButton6)
'    Set cmdMenuButton7 = VBInstance.Events.CommandBarEvents(cmdButton7)
    Set cmdMenuButton8 = VBInstance.Events.CommandBarEvents(cmdButton8)
    
    Dim cbNextCommand As Office.CommandBarControl
    Dim cbMenu As Office.CommandBar
    Set cbMenu = VBInstance.CommandBars.Item("File")
    If Not cbMenu Is Nothing Then
        For Each cbNextCommand In cbMenu.Controls
            Select Case cbNextCommand.Caption
                Case "Make..."
                    Set cbMenuCommandBar4 = cbNextCommand
                    Set MenuHandler2 = VBInstance.Events.CommandBarEvents(cbMenuCommandBar4)
                Case "Make Project &Group...", "Make Project Group..."
                    Set MenuHandler6 = VBInstance.Events.CommandBarEvents(cbNextCommand)
                    Set cbMenuCommandBar2 = cbMenu.Controls.Add(1, , , cbNextCommand.Index)
                    cbMenuCommandBar2.Caption = "&Build Project Release"
                    cbMenuCommandBar2.Tag = "&Build Project Release"
                    Set MenuHandler3 = VBInstance.Events.CommandBarEvents(cbMenuCommandBar2)
                    Set cbMenuCommandBar3 = cbMenu.Controls.Add(1, , , cbMenuCommandBar2.Index)
                    cbMenuCommandBar3.Caption = "Remake Pro&ject Build"
                    cbMenuCommandBar3.Tag = "Remake Pro&ject Build"
                    Set MenuHandler5 = VBInstance.Events.CommandBarEvents(cbMenuCommandBar3)
            End Select
        Next
    End If
    Set cbMenu = VBInstance.CommandBars.Item("Project")
    If Not cbMenu Is Nothing Then
        For Each cbNextCommand In cbMenu.Controls

            If Right(Replace(cbNextCommand.Caption, "&", ""), 13) = "Properties..." Then
                Set cbMenuCommandBar6 = cbNextCommand
                Set MenuHandler9 = VBInstance.Events.CommandBarEvents(cbMenuCommandBar6)
            End If
        Next
    End If
    Set cbMenu = VBInstance.CommandBars.Item("Tools")
    If Not cbMenu Is Nothing Then
        For Each cbNextCommand In cbMenu.Controls

            If Right(Replace(cbNextCommand.Caption, "&", ""), 23) = "Procedure Attributes..." Then
                Set cbMenuCommandBar7 = cbNextCommand
                Set MenuHandler10 = VBInstance.Events.CommandBarEvents(cbMenuCommandBar7)
            End If
        Next
    End If
    
    Set cbMenu = VBInstance.CommandBars.Item("Run")
    If Not cbMenu Is Nothing Then
        For Each cbNextCommand In cbMenu.Controls
            Select Case cbNextCommand.Caption
                Case "Start With &Full Compile"
                    Set MenuHandler1 = VBInstance.Events.CommandBarEvents(cbNextCommand)
                    Set cbMenuCommandBar1 = cbMenu.Controls.Add(1, , , cbNextCommand.Index, False)
                    cbMenuCommandBar1.Caption = "Start &The Executable"
                    cbMenuCommandBar1.Tag = "Start &The Executable"
                    Set MenuHandler4 = VBInstance.Events.CommandBarEvents(cbMenuCommandBar1)
                Case "&Start"
                    Set cmdBarBtn8 = cbNextCommand
                    Set cmdBarBtnEvents8 = VBInstance.Events.CommandBarEvents(cmdBarBtn8)
                Case "Brea&k"
                    Set cmdBarBtn9 = cbNextCommand
                    Set cmdBarBtnEvents9 = VBInstance.Events.CommandBarEvents(cmdBarBtn9)
                Case "&End"
                    Set cmdBarBtn10 = cbNextCommand
                    Set cmdBarBtnEvents10 = VBInstance.Events.CommandBarEvents(cmdBarBtn10)
            End Select
        Next
    End If
    Set cbNextCommand = Nothing
    Set cbMenu = Nothing
    
    Dim NxtPos As Boolean
    Set cbMenu = VBInstance.CommandBars.Item("Window")
    If Not cbMenu Is Nothing Then
        For Each cbNextCommand In cbMenu.Controls

            Select Case cbNextCommand.Caption
                Case "&Arrange Icons"
                    NxtPos = True
                Case Else
                    If NxtPos And cbNextCommand.BeginGroup Then
                        Set cbMenuCommandBar5 = cbMenu.Controls.Add(1, , , cbNextCommand.Index, False)
                        cbMenuCommandBar5.BeginGroup = True
                        cbMenuCommandBar5.Caption = LayoutCaption
                        Set MenuHandler8 = VBInstance.Events.CommandBarEvents(cbMenuCommandBar5)
                    End If
            End Select
        Next
    End If
    Set cbNextCommand = Nothing
    Set cbMenu = Nothing


'    Set cmdBarBtnEvents1 = VBInstance.Events.CommandBarEvents(VBInstance.CommandBars("Run").Controls("Start"))
'    Set cmdBarBtnEvents2 = VBInstance.Events.CommandBarEvents(VBInstance.CommandBars("Run").Controls("Brea&k"))
'    Set cmdBarBtnEvents3 = VBInstance.Events.CommandBarEvents(VBInstance.CommandBars("Run").Controls("&End"))

    
'    Set cmdBarBtnEvents8 = VBInstance.Events.CommandBarEvents(VBInstance.CommandBars("Run").Controls("Start"))
'    Set cmdBarBtnEvents9 = VBInstance.Events.CommandBarEvents(VBInstance.CommandBars("Run").Controls("Brea&k"))
'    Set cmdBarBtnEvents10 = VBInstance.Events.CommandBarEvents(VBInstance.CommandBars("Run").Controls("&End"))
    
    
    Set VBWindow = VBInstance.Windows.CreateToolWindow(AddInInst, "BasicNeotext.Settings", "Compiler Settings", guidPos, docSettings)
    VBWindow.Visible = GetSetting("BasicNeotext", "Options", "Settings_Visible", VBWindow.Visible)
    
    Set docSettings.VBInstance = VBInstance
    Set FCE = VBInstance.Events.FileControlEvents(Nothing)
    
    docSettings.StartTimer
    ''uncomment the following two lines to hook F5
    
    
End Sub


Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error GoTo exitthis
    On Local Error GoTo exitthis
    
    docSettings.StopTimer
    
    If Not CmdBar Is Nothing Then
        SaveSetting "BasicNeotext", "Options", "ToolBar_Position", CmdBar.Position
        SaveSetting "BasicNeotext", "Options", "ToolBar_RowIndex", CmdBar.RowIndex
        SaveSetting "BasicNeotext", "Options", "ToolBar_Left", CmdBar.Left
        SaveSetting "BasicNeotext", "Options", "ToolBar_Top", CmdBar.Top
        SaveSetting "BasicNeotext", "Options", "ToolBar_Visible", CmdBar.Visible
    End If
  
    On Error Resume Next
    If Not cmdMenuButton1 Is Nothing Then Set cmdMenuButton1 = Nothing
    If Not cmdMenuButton2 Is Nothing Then Set cmdMenuButton2 = Nothing
    If Not cmdMenuButton3 Is Nothing Then Set cmdMenuButton3 = Nothing
    If Not cmdMenuButton4 Is Nothing Then Set cmdMenuButton4 = Nothing
    If Not cmdMenuButton5 Is Nothing Then Set cmdMenuButton5 = Nothing
    If Not cmdMenuButton6 Is Nothing Then Set cmdMenuButton6 = Nothing
    If Not cmdMenuButton7 Is Nothing Then Set cmdMenuButton7 = Nothing
    If Not cmdMenuButton8 Is Nothing Then Set cmdMenuButton8 = Nothing
    
    If Not MenuHandler10 Is Nothing Then Set MenuHandler10 = Nothing
    If Not MenuHandler9 Is Nothing Then Set MenuHandler9 = Nothing
    If Not MenuHandler8 Is Nothing Then Set MenuHandler8 = Nothing
    If Not MenuHandler7 Is Nothing Then Set MenuHandler7 = Nothing
    If Not MenuHandler6 Is Nothing Then Set MenuHandler6 = Nothing
    If Not MenuHandler5 Is Nothing Then Set MenuHandler5 = Nothing
    If Not MenuHandler4 Is Nothing Then Set MenuHandler4 = Nothing
    If Not MenuHandler3 Is Nothing Then Set MenuHandler3 = Nothing
    If Not MenuHandler2 Is Nothing Then Set MenuHandler2 = Nothing
    If Not MenuHandler1 Is Nothing Then Set MenuHandler1 = Nothing
    
    If Not cbMenuCommandBar1 Is Nothing Then
        If cbMenuCommandBar1.Tag = "Start &The Executable" Then
            cbMenuCommandBar1.Delete
            Set cbMenuCommandBar1 = Nothing
        End If
    End If
    If Not cbMenuCommandBar2 Is Nothing Then
        If cbMenuCommandBar2.Tag = "&Build Project Release" Then
            cbMenuCommandBar2.Delete
            Set cbMenuCommandBar2 = Nothing
        End If
    End If
    If Not cbMenuCommandBar3 Is Nothing Then
        If cbMenuCommandBar3.Tag = "Remake Pro&ject Build" Then
            cbMenuCommandBar3.Delete
            Set cbMenuCommandBar3 = Nothing
        End If
    End If
    If Not cbMenuCommandBar4 Is Nothing Then
        Set cbMenuCommandBar4 = Nothing
    End If
    If Not cbMenuCommandBar5 Is Nothing Then
        cbMenuCommandBar5.Delete
        Set cbMenuCommandBar5 = Nothing
    End If
    If Not cbMenuCommandBar6 Is Nothing Then
        Set cbMenuCommandBar6 = Nothing
    End If
    If Not cbMenuCommandBar7 Is Nothing Then
        Set cbMenuCommandBar7 = Nothing
    End If
    
    If Not cmdButton1 Is Nothing Then cmdButton1.Delete
    If Not cmdButton2 Is Nothing Then cmdButton2.Delete
    If Not cmdButton3 Is Nothing Then cmdButton3.Delete
    If Not cmdButton4 Is Nothing Then cmdButton4.Delete
    If Not cmdButton5 Is Nothing Then cmdButton5.Delete
    If Not cmdButton6 Is Nothing Then cmdButton6.Delete
    If Not cmdButton7 Is Nothing Then cmdButton7.Delete
    If Not cmdButton8 Is Nothing Then cmdButton8.Delete
    
    Set cmdButton1 = Nothing
    Set cmdButton2 = Nothing
    Set cmdButton3 = Nothing
    Set cmdButton4 = Nothing
    Set cmdButton5 = Nothing
    Set cmdButton6 = Nothing
    Set cmdButton7 = Nothing
    
    If Not cmdBarBtn1 Is Nothing Then cmdBarBtn1.Delete
    If Not cmdBarBtn2 Is Nothing Then cmdBarBtn2.Delete
    If Not cmdBarBtn3 Is Nothing Then cmdBarBtn3.Delete
    If Not cmdBarBtn4 Is Nothing Then cmdBarBtn4.Delete
    If Not cmdBarBtn5 Is Nothing Then cmdBarBtn5.Delete
    If Not cmdBarBtn6 Is Nothing Then cmdBarBtn6.Delete
    If Not cmdBarBtn7 Is Nothing Then cmdBarBtn7.Delete
    If Not cmdBarBtn8 Is Nothing Then cmdBarBtn8.Delete
    If Not cmdBarBtn9 Is Nothing Then cmdBarBtn9.Delete
    If Not cmdBarBtn10 Is Nothing Then cmdBarBtn10.Delete

    Set cmdBarBtn1 = Nothing
    Set cmdBarBtn2 = Nothing
    Set cmdBarBtn3 = Nothing
    Set cmdBarBtn4 = Nothing
    Set cmdBarBtn5 = Nothing
    
    Set cmdBarBtnEvents1 = Nothing
    Set cmdBarBtnEvents2 = Nothing
    Set cmdBarBtnEvents3 = Nothing
    Set cmdBarBtnEvents4 = Nothing
    Set cmdBarBtnEvents5 = Nothing
    Set cmdBarBtnEvents6 = Nothing
    Set cmdBarBtnEvents7 = Nothing
    Set cmdBarBtnEvents8 = Nothing
    Set cmdBarBtnEvents9 = Nothing
    Set cmdBarBtnEvents10 = Nothing
    
    If Not CmdBar Is Nothing Then
        CmdBar.Delete
    End If
    Set CmdBar = Nothing
    
    SaveSetting "BasicNeotext", "Options", "Settings_Visible", VBWindow.Visible

    VBWindow.Close
    Set VBWindow = Nothing

    Set docSettings.VBInstance = Nothing
    Set docSettings = Nothing
    
exitthis:
    If Err Then Err.Clear
    On Error GoTo 0
    On Local Error GoTo 0
End Sub

Private Sub AddinInstance_OnStartupComplete(custom() As Variant)
    SetVBSettings
    'DescriptionsStartup VBInstance
End Sub

Private Sub AddinInstance_Terminate()
    QuitFail = -1
    Set FCE = Nothing
End Sub

Private Function LayoutCaption() As String
    LayoutCaption = IIf(PathExists(GetFilePath(AppEXE(False)) & "\REG.BAK", True), "&Release Window Layout", "&Preserve Window Layout")
End Function


'Private Sub MenuHandler8_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'    On Error GoTo exitthis
'    On Local Error GoTo exitthis
'
'    If PathExists(GetFilePath(AppEXE(False)) & "\REG.BAK", True) Then
'        Kill GetFilePath(AppEXE(False)) & "\REG.BAK"
'    Else
'        SetBNSettings
'    End If
'
'    CommandBarControl.Caption = LayoutCaption
'
'exitthis:
'    If Err Then Err.Clear
'    On Error GoTo 0
'    On Local Error GoTo 0
'End Sub

Private Sub cmdBarBtnEvents8_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    StartEvent CommandBarControl, handled, CancelDefault
End Sub

Private Sub cmdBarBtnEvents9_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    BreakEvent CommandBarControl, handled, CancelDefault
End Sub

Private Sub cmdBarBtnEvents10_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    EndEvent CommandBarControl, handled, CancelDefault
End Sub

Private Sub cmdBarBtnEvents1_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    StartEvent CommandBarControl, handled, CancelDefault
End Sub

Private Sub cmdBarBtnEvents2_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    BreakEvent CommandBarControl, handled, CancelDefault
End Sub

Private Sub cmdBarBtnEvents3_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    EndEvent CommandBarControl, handled, CancelDefault
End Sub

''Private Sub cmdBarBtnEvents4_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
''    StartEvent CommandBarControl, handled, CancelDefault
''End Sub
''
''Private Sub cmdBarBtnEvents5_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
''    BreakEvent CommandBarControl, handled, CancelDefault
''End Sub
''
''Private Sub cmdBarBtnEvents6_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
''    EndEvent CommandBarControl, handled, CancelDefault
''End Sub

Private Sub cmdBarBtnEvents7_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    SetBNSettings
    VBWindow.Visible = Not VBWindow.Visible
End Sub

Private Sub cmdMenuButton1_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    SignExecutable CommandBarControl, handled, CancelDefault
End Sub

Private Sub cmdMenuButton2_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    StartFullCompile CommandBarControl, handled, CancelDefault
End Sub

Private Sub cmdMenuButton3_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    MakeDialog CommandBarControl, handled, CancelDefault
End Sub

Private Sub cmdMenuButton5_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    StartExecutable CommandBarControl, handled, CancelDefault
End Sub

Private Sub cmdMenuButton6_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    RemakeProject CommandBarControl, handled, CancelDefault
End Sub

Private Sub cmdMenuButton8_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    StopExecutable CommandBarControl, handled, CancelDefault
End Sub

Private Sub MenuHandler1_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    StartFullCompile CommandBarControl, handled, CancelDefault
End Sub

Private Sub MenuHandler10_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    ProcedureAttributes CommandBarControl, handled, CancelDefault
End Sub

Private Sub MenuHandler2_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    MakeDialog CommandBarControl, handled, CancelDefault
End Sub

Private Sub MenuHandler3_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    BuildRelease CommandBarControl, handled, CancelDefault
End Sub

Private Sub MenuHandler4_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    StartExecutable CommandBarControl, handled, CancelDefault
End Sub

Private Sub MenuHandler5_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    RemakeProject CommandBarControl, handled, CancelDefault
End Sub

Private Sub MenuHandler6_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    MakeGroup CommandBarControl, handled, CancelDefault
End Sub

Private Sub MenuHandler7_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    handled = True
    SignExecutable CommandBarControl, handled, CancelDefault
End Sub

Private Sub MenuHandler9_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    ProjectProperties CommandBarControl, handled, CancelDefault
End Sub

Public Sub StartEvent(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'''    '"&Start" from menu and/or toolbar this should be
'''    '   catch all hook for the start of project and
'''    '   VBIDE condcomp should be -1, all spaces out
    AdjustHeaders

End Sub
Public Sub BreakEvent(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'''    '"Brea&k" from menu and/or toolbar this should be
'''    '   catch all hook for the break of project run

End Sub

Public Sub EndEvent(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'''    '"&End" from menu and/or toolbar this should be
'''    '   catch all hook for the stop of project run

End Sub
Private Function StartupProjectEXE(ByVal BinPath As String) As String

    Dim i2 As Project
    Dim i As Project

    If modMain.Projs.Includes.count > 0 Then
        For Each i In modMain.Projs.Includes
            If UCase(Right(i.Compiled, 4)) = ".EXE" Then
                If i.Includes.count > 0 Then
                    For Each i2 In i.Includes
                        If LCase(i2.Compiled) = LCase(BinPath) Then
                            StartupProjectEXE = i.Compiled
                        End If
                    Next
                End If
            End If
        Next
    End If
    If StartupProjectEXE = "" And UCase(Right(BinPath, 4)) = ".EXE" Then
        StartupProjectEXE = BinPath
    End If

End Function

Private Function StartupProject() As VBProject
    Dim vbp
    For Each vbp In VBInstance.VBProjects
    'Debug.Print vbp.StartMode & " " & vbp.BuildFileName
        If vbp.StartMode = 0 Then
            Set StartupProject = vbp
            Exit Function
        End If
    Next
End Function


Private Sub StartFullCompile(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'''    '"Start With &Full Compile"-539
'''    '   Same as the BuiltIn menu item
'''    '   but we want this hook options
                    
    Dim p As Project
    Set p = modMain.Projs
    Dim startexe As String

    startexe = StartupProjectEXE(StartupProject.BuildFileName)
    
    If Not (VBInstance.ActiveVBProject Is Nothing) Then
        If PathExists(startexe, True) Then
            If ProcessMake() Then RunProcessEx startexe, Projs.CmdLine
        End If
    End If
   
End Sub


Private Sub MakeDialog(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'''    '"Make..."-215
'''    '   Same as the BuildIn menu item
'''    '   but we want this hook options

    On Error Resume Next
    On Local Error Resume Next
    
    If StopProcess Then
    
        cbMenuCommandBar4.Execute
    
    End If
    If Err Then Err.Clear
End Sub

Private Function ShowMessage(ByVal Title As String, ByVal Message As String, ByVal IncludeCancel As Boolean) As Long
    Static onetatime As Boolean
    If Not onetatime Then
        onetatime = True
        Dim frm As New frmHelp
        frm.Command2.Caption = "&Yes"
        frm.Command1.Caption = "&No"
        frm.Command3.Caption = "&Cancel"
        If IncludeCancel Then frm.Command3.Visible = True
        frm.Command2.Visible = True
        frm.Label26.Visible = False
        frm.Frame1.Visible = False
        frm.Label3.Visible = False
        frm.Command2.Top = 1020 ' frm.Frame1.Top
        frm.Height = 1935 'frm.Frame1.Top + frm.Command2.Height + (frm.Height - (frm.Command2.Top + frm.Command2.Height))
        frm.Command1.Top = 1020 ' frm.Frame1.Top
        frm.Command3.Top = 1020 ' frm.Frame1.Top
        frm.Label2.Caption = Message
        frm.Image1.Visible = True
        frm.Caption = Title
        frm.Label2.Width = frm.Label2.Width - 2000
        frm.Command1.Left = frm.Command1.Left - 2000
        frm.Command2.Left = frm.Command2.Left - 2000
        frm.Command3.Left = frm.Command3.Left - 2000
        If Not IncludeCancel Then
            frm.Command2.Left = frm.Command1.Left
            frm.Command1.Left = frm.Command3.Left
        End If
        frm.Width = frm.Width - 3000
        frm.Tag = 0
        frm.Show
        TopMostForm frm, True, True
        Do Until frm.Tag <> 0
            DoTasks
        Loop
        ShowMessage = frm.Tag
        Unload frm
        onetatime = False
    End If
End Function
Private Function ProcessMake(Optional ByRef UseProject As VBProject = Nothing) As Boolean

'    Dim killexe As String
'    If StartEXE = "" Then
'        killexe = StartupProjectEXE
'    Else
'        killexe = StartEXE
'    End If
'
'    If IsProccessEXERunning(GetFileName(killexe)) Then
'        Dim tmp As Long
'        If CLng(GetSetting("BasicNeotext", "Options", "KillBeforeMake", 0)) = 0 Then
'
'            tmp = ShowMessage("Permission denied", "The process is already running, do you want to" & vbCrLf & _
'                              "terminate it and continue to make the executable?", True)
'            If tmp = vbCancel Then
'                Exit Function
'            End If
'        Else
'            tmp = vbYes
'        End If
'        If tmp = vbYes Then
'            KillApp GetFileName(killexe)
'        End If
'    End If
    If UseProject Is Nothing Then Set UseProject = StartupProject
    
    On Error Resume Next
    On Local Error Resume Next
    If StopProcess Then
        DoMakeProj UseProject
        ProcessMake = True
    End If
End Function
Private Sub BuildRelease(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'''    '"&Build Project Release"-184
'''    '   Same as Remake but sets the
'''    '   VBIDE condcomp flag to 0 and
'''    '   signs when enabled to do so
    Dim buildexe As String
    buildexe = StartupProjectEXE(VBInstance.ActiveVBProject.BuildFileName)
    
    If Not (VBInstance.ActiveVBProject Is Nothing) Then
        If PathExists(buildexe, True) Then
            If ProcessMake() Then SignTool buildexe, SignAndStamp
        End If
    End If

End Sub

Private Function StopProcess(Optional ByVal startexe As String = "") As Boolean
    Dim stopexe As String
    If startexe = "" Then
        stopexe = StartupProjectEXE(StartupProject.BuildFileName)
    Else
        stopexe = startexe
    End If

    If Not (VBInstance.ActiveVBProject Is Nothing) Then
        If IsProccessEXERunning(GetFileName(stopexe)) Then

            If CLng(GetSetting("BasicNeotext", "Options", "KillBeforeMake", 0)) = 0 Then
                Dim tmp As Long
                tmp = ShowMessage("Process is Running", "Are you sure you want to terminate it?", False)
                If tmp = vbYes Then
                    Do While IsProccessEXERunning(GetFileName(stopexe))
                        KillApp GetFileName(stopexe)
                    Loop
                    
                    StopProcess = True
                End If
            Else
                Do While IsProccessEXERunning(GetFileName(stopexe))
                    KillApp GetFileName(stopexe)
                Loop
                StopProcess = True
            End If
        Else
            StopProcess = True
        End If
    End If
    '(Not IsProccessEXERunning(GetFileName(stopexe)))
    
End Function
Private Sub StopExecutable(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'''    '"Stop The E&xecutable"-348
'''    '   kills the running process
'''    '   (if it is running of course)

    StopProcess
'    Dim stopexe As String
'    stopexe = StartupProjectEXE(VBInstance.ActiveVBProject.BuildFileName)
'
'    If Not (VBInstance.ActiveVBProject Is Nothing) Then
'        If IsProccessEXERunning(GetFileName(VBInstance.ActiveVBProject.BuildFileName)) Then
'
'            If CLng(GetSetting("BasicNeotext", "Options", "KillBeforeMake", 0)) = 0 Then
'                Dim tmp As Long
'                tmp = ShowMessage("Process is Running", "Are you sure you want to terminate it?", False)
'                If tmp = vbYes Then
'                    KillApp GetFileName(VBInstance.ActiveVBProject.BuildFileName)
'                End If
'            Else
'                KillApp GetFileName(VBInstance.ActiveVBProject.BuildFileName)
'            End If
'
'        End If
'    End If

End Sub
Private Sub DoMakeProj(ByRef vbp As VBProject)
    On Error Resume Next
    On Local Error Resume Next
    
    If PathExists(vbp.BuildFileName, True) Then
        ChDir GetFilePath(vbp.BuildFileName)
    End If
    
    vbp.MakeCompiledFile
    
    If Err.Number <> 0 Then

        If Err.Number <> 0 Then
            MsgBox Err.Description, vbCritical, "Error " & Err.Number
        End If
    
        Err.Clear
    End If
    
End Sub
Private Sub StartExecutable(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'''    '"Start &The Executable"-459
'''    '   Run the current executable
'''    '   with out compiling it first
'''    '   (must been already built)
    
    
    If Not (VBInstance.ActiveVBProject Is Nothing) Then
        Dim startexe As String
        startexe = StartupProjectEXE(StartupProject.BuildFileName)
        If startexe = "" Then
            Dim tmp As Long
            tmp = ShowMessage("Process Not Compiled", "This process has not been compiled, do you want to compile it?", False)
            If tmp = vbYes Then
                DoMakeProj StartupProject
            Else
                Exit Sub
            End If
        End If
        If PathExists(startexe, True) Then
            RunProcessEx startexe, Projs.CmdLine
        End If
    End If

End Sub

Private Sub RemakeProject(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'''    '"Remake Pro&ject Build"-37
'''    '   Automatically rebuild the executable with out
'''    '   prompting for location (must been already built)
    If Not (VBInstance.ActiveVBProject Is Nothing) Then
        If StopProcess Then
        
        'If IsProccessEXERunning(GetFileName(VBInstance.ActiveVBProject.BuildFileName)) Then KillApp GetFileName(VBInstance.ActiveVBProject.BuildFileName)
            ProcessMake VBInstance.ActiveVBProject
        End If
    End If

End Sub


Private Sub MakeGroup(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'''    '"Make Project Group..."-185
'''    '   Same as the BuiltIn menu item
'''    '   but we want this hook options
    On Error Resume Next
    On Local Error Resume Next
    
    If StopProcess Then
        If Not handled Then
            VBInstance.CommandBars("File").Controls("Make Project &Group...").Execute
        End If
    End If
    
    If Err Then Err.Clear
End Sub

Private Sub SignExecutable(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'''    '   Signs the project's executable
'''    '   indpendent of command line or
'''    '   build release, must be in list
    
    If Not (VBInstance.ActiveVBProject Is Nothing) Then
        SignTool VBInstance.ActiveVBProject.BuildFileName, SignAndStamp
    End If
    
End Sub

Private Sub ProjectProperties(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'''    '   if the project properties dialog is invoked



   ' QuitCall = True
   ' VBInstance.Quit
    

End Sub

Private Sub ProcedureAttributes(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'''    '   if the procedure attributes dialog is invoked
  ' VBInstance.ActiveCodePane.CodeModule
    'BuildComments CommentsToAttribute, VBInstance.ActiveVBProject.Name, VBInstance.ActiveCodePane.CodeModule
    UpdateCommentToAttributeDescriptions VBInstance
End Sub

Private Sub MenuHandler8_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    
    If PathExists(GetFilePath(AppEXE(False)) & "\REG.BAK", True) Then
        Kill GetFilePath(AppEXE(False)) & "\REG.BAK"
    Else
        WriteFile GetFilePath(AppEXE(False)) & "\REG.BAK", _
            StrConv(GetSettingByte(modRegistry.HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0", "Dock"), vbUnicode) & vbCrLf & _
            StrConv(GetSettingByte(modRegistry.HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0", "Tool"), vbUnicode) & vbCrLf & _
            StrConv(GetSettingByte(modRegistry.HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0", "UI"), vbUnicode)
    End If

    CommandBarControl.Caption = LayoutCaption
    
exitthis:
    If Err Then Err.Clear
    On Error GoTo 0
    On Local Error GoTo 0
End Sub

