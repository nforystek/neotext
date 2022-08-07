VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Timer"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'TOP DOWN

Option Compare Binary

Private lHwnd As Long
Private bEnabled As Boolean
Private mInterval As Long
Private mExecuted As Boolean

Public Event OnTicking()

Public Property Get Enabled() As Boolean
    Enabled = bEnabled Or mExecuted
End Property
Public Property Let Enabled(ByVal newVal As Boolean)
    If (mExecuted And newVal) Or _
        ((bEnabled Or mExecuted) _
        And (Not newVal)) Then
        StopTimer Me
    End If
    If (newVal And (Not (bEnabled = newVal))) Or mExecuted Then
        If newVal Then StartTimer Me
    End If
    mExecuted = False
    bEnabled = newVal
End Property

Public Property Let Interval(ByVal vData As Long)
    If Not (mInterval = vData) Then
        mInterval = vData
        If bEnabled Then
            StopTimer Me
            StartTimer Me
        End If
    End If
End Property
Public Property Get Interval() As Long
    Interval = mInterval
End Property

Friend Sub TickTimer()
    RaiseEvent OnTicking
End Sub

Friend Property Get Executed() As Boolean
    Executed = mExecuted
End Property

Friend Static Property Get hwnd() As Long
    hwnd = lHwnd
End Property

Public Sub Execute()
    If bEnabled Then
        StopTimer Me
        bEnabled = False
    End If
    mExecuted = True
    StartTimer Me
End Sub

Friend Sub Disable()
    mExecuted = False
    bEnabled = False
End Sub

Private Sub Class_Initialize()
    mInterval = 1000
    lHwnd = WindowInitialize()
End Sub

Private Sub Class_Terminate()
    Enabled = False
    WindowTerminate lHwnd
End Sub



