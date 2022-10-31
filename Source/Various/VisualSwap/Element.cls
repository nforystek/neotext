VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Element"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private locate As New Vector
Private orig As New Vector
Private label As String
Private first As String
Private clr As Variant


Private coords As Collection

Public Property Get Color() As Variant
    Color = clr
End Property
Public Property Let Color(ByVal RHS As Variant)
    clr = RHS
End Property

Public Property Get Coordinates() As Collection
    Set Coordinates = coords
End Property

Public Property Get Origin() As Vector
   Set Origin = orig
End Property
Public Property Set Origin(ByVal RHS As Vector)
   Set orig = RHS
End Property

Public Property Get Located() As Vector
    Set Located = locate
End Property
Public Property Set Located(ByVal RHS As Vector)
    Set locate = RHS
End Property


Public Property Get Holder() As String
    Holder = first
End Property
Public Property Let Holder(ByVal RHS As String)
    first = RHS
End Property

Public Property Get Display() As String
    Display = label
End Property
Public Property Let Display(ByVal RHS As String)
    label = RHS
End Property

Private Sub Class_Initialize()
    Set coords = New Collection
    clr = vbWhite
End Sub

Private Sub Class_Terminate()
    Do Until coords.Count = 0
        coords.Remove 1
    Loop

    Set coords = Nothing
End Sub
