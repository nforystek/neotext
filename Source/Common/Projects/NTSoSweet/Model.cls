VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Model"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary

Private sColors As Colors
Private sHoops As Hoops
Private sReams As Reams

Public Property Get Colors() As Colors
    Set Colors = sColors
End Property

Public Property Set Colors(ByRef newval As Colors)
    Set sColors = newval
End Property

Public Property Get Hoops() As Hoops
    Set Hoops = sHoops
End Property

Friend Property Set Hoops(ByRef newval As Hoops)
    Set sHoops = newval
End Property

Public Property Get Reams() As Reams
    Set Reams = sReams
End Property

Friend Property Set Reams(ByRef newval As Reams)
    Set sReams = newval
End Property

'Public Property Get NewEnum() As Object
'    Set NewEnum = Me.[_NewEnum]
'End Property
