VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pIndex As Long
Private pRepeat As Single
Private pVolume As Single
Private pOrigin As Point
Private pEnabled As Boolean
Private pKey As String
Private pSerialize As Boolean
Private pRange As Single
Private pFileName As String

Public Property Get Serialize() As Boolean
    Serialize = True
End Property
Public Property Let Serialize(ByVal RHS As Boolean)
    pSerialize = RHS
End Property

Public Property Get Key() As String
    Key = pKey
End Property
Public Property Let Key(ByVal RHS As String)
    If RHS <> "" Then
        If pKey <> RHS And pKey <> "" Then If All.Exists(pKey) Then All.Remove pKey
        If All.Exists(RHS) Then
            All.Remove RHS
            All.Add Me, RHS
        End If
        pKey = RHS
    End If
End Property

Public Property Get FileName() As String
     FileName = pFileName
End Property
Public Property Let FileName(ByVal RHS As String)
    If pFileName <> "" Then
        StopWave Index
    End If
    If PathExists(AppPath & "Sounds\" & RHS, True) Then
        LoadWave Index, AppPath & "Sounds\" & RHS
        pFileName = RHS
    End If
End Property

Public Property Get Enabled() As Boolean
    Enabled = pEnabled
End Property
Public Property Let Enabled(ByVal RHS As Boolean)
    pEnabled = RHS
End Property

Public Property Get Repeat() As Boolean
    Repeat = pRepeat
End Property
Public Property Let Repeat(ByVal RHS As Boolean)
    pRepeat = RHS
End Property
Public Property Get Volume() As Single
    Volume = pVolume
End Property
Public Property Let Volume(ByVal RHS As Single)
    pVolume = RHS
    VolumeWave Index, RHS
End Property
Public Sub PlaySound()
    PlayWave Index, pRepeat
End Sub
Public Sub StopPlay()
    StopWave Index
End Sub

Friend Property Get Index() As Long
    Index = pIndex
End Property
Friend Property Let Index(ByVal RHS As Long)
    pIndex = RHS
End Property

Public Property Get Origin() As Point
    Set Origin = pOrigin
End Property
Public Property Set Origin(ByRef RHS As Point)
    Set pOrigin = RHS
End Property

Public Property Get Range() As Single
    Range = pRange
End Property
Public Property Let Range(ByVal RHS As Single)
    pRange = RHS
End Property

'Public Type MySound
'    Key As String
'    Enable As Boolean
'    Origin As D3DVECTOR
'    Range As Single
'
'    Repeat As Boolean
'    Index As Long
'End Type
Private Sub Class_Initialize()
    Stats_Sound_Count = Stats_Sound_Count + 1
    Index = Sounds.Count + 1
    Key = "K" & ObjPtr(Me)
    Set pOrigin = New Point
    
End Sub

Private Sub Class_Terminate()
    Set pOrigin = Nothing
    Stats_Sound_Count = Stats_Sound_Count - 1
End Sub
