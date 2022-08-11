VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   10020
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   23265
   _ExtentX        =   41037
   _ExtentY        =   17674
   _Version        =   393216
   Description     =   "Enhancements for Visual Basic 6.0"
   DisplayName     =   "VB 6 Neotext Basic"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
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
    
    
    modHotKey.UnHook
    
    Dim cP As Window
    For Each cP In VBInstance.Windows
        Select Case StrReverse(NextArg(StrReverse(cP.Caption), " "))
            Case "(UserControl)", "(Form)", "(UserDocument)", "(AddInDesigner)"
                cP.Close
        End Select
    Next
    
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
    
    Set docSettings.FCE = VBInstance.Events.FileControlEvents(Nothing)
    
'    Dim Ret As Long
'    Ret = modHotKey.EnumThreadWindows(modHotKey.GetCurrentThreadId, AddressOf modHotKey.EnumThreadProc, ObjPtr(Me))

'''    docSettings.SetUIEvents Me
    

End Sub


Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error GoTo exitthis
    On Local Error GoTo exitthis
    
    modHotKey.UnHook
    
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
    
    If Not MenuHandler7 Is Nothing Then Set MenuHandler9 = Nothing
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
    
    If Not cmdButton1 Is Nothing Then cmdButton1.Delete
    If Not cmdButton2 Is Nothing Then cmdButton2.Delete
    If Not cmdButton3 Is Nothing Then cmdButton3.Delete
    If Not cmdButton4 Is Nothing Then cmdButton4.Delete
    If Not cmdButton5 Is Nothing Then cmdButton5.Delete
    If Not cmdButton6 Is Nothing Then cmdButton6.Delete
    If Not cmdButton7 Is Nothing Then cmdButton7.Delete
    
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
    
    Set docSettings.FCE = Nothing
    
    Set docSettings = Nothing
    
exitthis:
    If Err Then Err.Clear
    On Error GoTo 0
    On Local Error GoTo 0
End Sub

Private Sub AddinInstance_OnStartupComplete(custom() As Variant)
'
End Sub

Private Sub AddinInstance_Terminate()
    modHotKey.UnHook
    
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

Private Sub MenuHandler1_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    StartFullCompile CommandBarControl, handled, CancelDefault
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


Private Sub StartFullCompile(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'''    '"Start With &Full Compile"-539
'''    '   Same as the BuiltIn menu item
'''    '   but we want this hook options
    
    If Not (VBInstance.ActiveVBProject Is Nothing) Then
        If PathExists(VBInstance.ActiveVBProject.BuildFileName, True) Then
            VBInstance.ActiveVBProject.MakeCompiledFile
            RunProcessEx VBInstance.ActiveVBProject.BuildFileName, Projs.CmdLine
        End If
    End If
   
End Sub


Private Sub MakeDialog(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'''    '"Make..."-215
'''    '   Same as the BuildIn menu item
'''    '   but we want this hook options

    On Error Resume Next
    On Local Error Resume Next
    
    cbMenuCommandBar4.Execute

    If Err Then Err.Clear
End Sub

Private Sub BuildRelease(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'''    '"&Build Project Release"-184
'''    '   Same as Remake but sets the
'''    '   VBIDE condcomp flag to 0 and
'''    '   signs when enabled to do so
    
    If Not (VBInstance.ActiveVBProject Is Nothing) Then
        If PathExists(VBInstance.ActiveVBProject.BuildFileName, True) Then
            VBInstance.ActiveVBProject.MakeCompiledFile
            SignTool VBInstance.ActiveVBProject.BuildFileName, SignAndStamp
        End If
    End If
    

End Sub

Private Sub StartExecutable(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'''    '"Start &The Executable"-459
'''    '   Run the current executable
'''    '   with out compiling it first
'''    '   (must been already built)
    
    If Not (VBInstance.ActiveVBProject Is Nothing) Then
        If PathExists(VBInstance.ActiveVBProject.BuildFileName, True) Then
            RunProcessEx VBInstance.ActiveVBProject.BuildFileName, Projs.CmdLine
        End If
    End If

End Sub

Private Sub RemakeProject(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'''    '"Remake Pro&ject Build"-37
'''    '   Automatically rebuild the executable with out
'''    '   prompting for location (must been already built)
    If Not (VBInstance.ActiveVBProject Is Nothing) Then
        VBInstance.ActiveVBProject.MakeCompiledFile
    End If

End Sub


Private Sub MakeGroup(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'''    '"Make Project Group..."-185
'''    '   Same as the BuiltIn menu item
'''    '   but we want this hook options
    On Error Resume Next
    On Local Error Resume Next
    
    If Not handled Then
        VBInstance.CommandBars("File").Controls("Make Project &Group...").Execute
    End If
    
    If Err Then Err.Clear
End Sub

Private Sub SignExecutable(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'''    '"Sign Executable"-30
'''    '   Signs the project's executable
'''    '   indpendent of command line or
'''    '   build release, must be in list
    
    If Not (VBInstance.ActiveVBProject Is Nothing) Then
        SignTool VBInstance.ActiveVBProject.BuildFileName, SignAndStamp
    End If
    
End Sub

Private Sub ProjectProperties(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'''    '"Prop&erties..." from menu, this should happen
'''    '   if the project properties dialog is invoked


End Sub

Private Sub MenuHandler8_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error GoTo exitthis
    On Local Error GoTo exitthis
    
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

