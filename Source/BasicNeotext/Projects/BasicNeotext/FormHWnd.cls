VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormHWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private phwnd As Long
Private pCM As CodeModule
Private pVis As Boolean

'
'Public Property Get ProjectName() As String
'    ProjectName = pProjName
'End Property
'Public Property Let ProjectName(ByVal RHS As String)
'    pProjName = RHS
'End Property
Public Property Get hWnd() As Long
    hWnd = phwnd
End Property
Public Property Let hWnd(ByVal RHS As Long)
    phwnd = RHS
    SaveVisibility
End Property
Public Sub SaveVisibility()
    pVis = IsWindowVisible(phwnd)
End Sub
Public Property Get CodeModule() As CodeModule
    Set CodeModule = pCM
End Property
Public Property Set CodeModule(ByRef RHS As CodeModule)
    Set pCM = RHS
End Property

Public Property Get Visible() As Boolean
    Visible = pVis
End Property

Public Sub Hide()
    If Not pCM Is Nothing Then
'        pCM.CodePane.Window.Close
    End If
End Sub

Public Sub Show()
    If Not pCM Is Nothing Then
'        pCM.CodePane.Show
    End If
End Sub

