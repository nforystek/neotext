VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IConnect"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Implements IDTExtensibility

Public Event OnAddInsUpdate(custom() As Variant)
Public Event OnConnection(ByVal VBInst As Object, ByVal ConnectMode As VBIDE.vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
Public Event OnDisconnection(ByVal RemoveMode As VBIDE.vbext_DisconnectMode, custom() As Variant)
Public Event OnStartupComplete(custom() As Variant)

Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
    RaiseEvent OnAddInsUpdate(custom)
End Sub

Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As VBIDE.vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
    RaiseEvent OnConnection(VBInst, ConnectMode, AddInInst, custom)
End Sub

Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As VBIDE.vbext_DisconnectMode, custom() As Variant)
    RaiseEvent OnDisconnection(RemoveMode, custom)
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    RaiseEvent OnStartupComplete(custom)
End Sub
'
'Private Sub IDTExtensibility2_OnAddInsUpdate(custom() As Variant)
'
'End Sub
'
'Private Sub IDTExtensibility2_OnBeginShutdown(custom() As Variant)
'
'End Sub
'
'Private Sub IDTExtensibility2_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
'
'End Sub
'
'Private Sub IDTExtensibility2_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
'
'End Sub
'
'Private Sub IDTExtensibility2_OnStartupComplete(custom() As Variant)
'
'End Sub
