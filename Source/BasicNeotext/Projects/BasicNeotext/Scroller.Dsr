VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Scroller 
   ClientHeight    =   10020
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   23265
   _ExtentX        =   41037
   _ExtentY        =   17674
   _Version        =   393216
   Description     =   "Enhancements for Visual Basic 6.0"
   DisplayName     =   "VB 6 Mouse Wheel"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Scroller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Option Compare Text

Private Sub AddinInstance_OnBeginShutdown(custom() As Variant)
    modMWheel.UnHook
End Sub

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    Dim ret As Long
    ret = modMWheel.EnumThreadWindows(modMWheel.GetCurrentThreadId, AddressOf modMWheel.EnumThreadProc, 0)
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    modMWheel.UnHook
End Sub

Private Sub AddinInstance_Terminate()
    modMWheel.UnHook
End Sub
