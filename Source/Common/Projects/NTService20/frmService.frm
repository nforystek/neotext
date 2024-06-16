VERSION 5.00
Object = "{E7BC34A0-BA86-11CF-84B1-CBC2DA68BF6C}#1.0#0"; "ntsvc.ocx"
Begin VB.Form frmService 
   BorderStyle     =   0  'None
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1305
   ControlBox      =   0   'False
   Icon            =   "frmService.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   1305
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin NTService.NTService NTService1 
      Left            =   360
      Top             =   180
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      ServiceName     =   "Simple"
      StartMode       =   3
   End
End
Attribute VB_Name = "frmService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Binary

Public Event Control(lEvent As Long)
Public Event ContinueService(Success As Boolean)
Public Event PauseService(Success As Boolean)
Public Event StartService(Success As Boolean)
Public Event StopService()

Private propName As String
Private propID As Long
Public Property Get ProcessName() As String
    ProcessName = propName
End Property
Public Property Let ProcessName(ByVal RHS As String)
    propName = RHS
End Property
Public Property Get ProcessID() As Long
    ProcessID = propID
End Property
Public Property Let ProcessID(ByVal RHS As Long)
    propID = RHS
End Property

Private Sub Form_Load()
    NTService1.ControlsAccepted = 7
   ' NTService1.Debug = (LCase(Trim(GetFileTitle(ProcessRunning(CStr(GetCurrentProcessId))))) = "vb6")
End Sub

Private Sub NTService1_Continue(Success As Boolean)
    RaiseEvent ContinueService(Success)
End Sub

Private Sub NTService1_Control(ByVal lEvent As Long)
    RaiseEvent Control(lEvent)
End Sub

Private Sub NTService1_Pause(Success As Boolean)
    RaiseEvent PauseService(Success)
End Sub

Private Sub NTService1_Start(Success As Boolean)
    RaiseEvent StartService(Success)
End Sub

Private Sub NTService1_Stop()
    RaiseEvent StopService
End Sub

