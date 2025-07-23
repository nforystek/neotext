VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Activities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Obj As MyObject
Private ActivityCount As Long
Private Activities() As MyActivity

Friend Property Get MyObj() As MyObject
    MyObj = Obj
End Property
Friend Property Let MyObj(ByRef RHS As MyObject)
    Obj = RHS
End Property


Public Property Get Count() As Long
    Count = ActivityCount
End Property

Public Property Get Item(ByVal Index As Long) As Activity
    Set Item = New Activity
    Item.MyObj = Obj
    Item.MyAct = Activities(Index)
    
End Property

Public Function AddActivity(ByRef Action As Long, ByVal aGUID As String, ByRef Data As D3DVECTOR, Optional ByRef Emphasis As Single = 0, Optional ByVal Friction As Single = 0, Optional ByVal Reactive As Single = -1, Optional ByVal Recount As Single = -1, Optional Script As String = "") As String
    ActivityCount = ActivityCount + 1
    Stats_Activity_Count = Stats_Activity_Count + 1
    ReDim Preserve Activities(1 To ActivityCount) As MyActivity
    With Activities(ActivityCount)
        .Identity = IIf(aGUID = "", Replace(modGuid.GUID, "-", "K"), aGUID)
        .Action = Action
        .Data = Data
        .Emphasis = Emphasis
        .Initials = Emphasis
        .Friction = Friction
        .Reactive = Reactive
        .latency = Timer
        .Recount = Recount
        .OnEvent = Script
        AddActivity = .Identity
    End With
End Function


Public Function DeleteActivity(ByVal MGUID As String) As Boolean
    Dim a As Long
    If ActivityCount > 0 Then
        a = 1
        Do While a <= ActivityCount
            If Activities(a).Identity = MGUID Or (Activities(a).Identity = "") Then
                SwapActivity Activities(a), Activities(ActivityCount)
                ActivityCount = ActivityCount - 1
                Stats_Activity_Count = Stats_Activity_Count - 1
                If ActivityCount > 0 Then
                    ReDim Preserve Activities(1 To ActivityCount) As MyActivity
                Else
                    Erase Activities
                End If
                DeleteActivity = True
            Else
                a = a + 1
            End If
        Loop
    End If
End Function

Private Sub Class_Terminate()
    Erase Activities
    ActivityCount = 0
End Sub
