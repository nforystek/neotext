#Const [True] = -1
#Const [False] = 0

Attribute VB_Name = "modMove"
#Const modMove = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module

Public Sub SwapActivity(ByRef val1 As MyActivity, ByRef val2 As MyActivity)
    Dim tmp As MyActivity
    tmp = val1
    val1 = val2
    val2 = tmp
End Sub

Public Function AddActivity(ByRef Obj As MyObject, ByRef dir As D3DVECTOR, ByRef burst As Single, ByVal Friction As Single) As String
    Obj.ActivityCount = Obj.ActivityCount + 1
    ReDim Preserve Obj.Activities(1 To Obj.ActivityCount) As MyActivity
    Obj.Activities(Obj.ActivityCount).Identity = Replace(GUID, "-","")
    Obj.Activities(Obj.ActivityCount).Direction = dir
    Obj.Activities(Obj.ActivityCount).BurstRate = burst
    Obj.Activities(Obj.ActivityCount).Friction = Friction
    AddActivity = Obj.Activities(Obj.ActivityCount).Identity
End Function

Public Function DeleteActivity(ByRef Obj As MyObject, ByVal MGUID As String) As Boolean
    Dim a As Long
    If Obj.ActivityCount > 0 Then
        If Obj.Activities(Obj.ActivityCount).Identity = MGUID Then
            Obj.ActivityCount = Obj.ActivityCount - 1
            If Obj.ActivityCount > 0 Then
                ReDim Preserve Obj.Activities(1 To Obj.ActivityCount) As MyActivity
            Else
                Erase Obj.Activities
            End If
            DeleteActivity = True
        Else
            For a = 1 To Obj.ActivityCount
                If Obj.Activities(a).Identity = MGUID Then
                    SwapActivity Obj.Activities(a), Obj.Activities(Obj.ActivityCount)
                    Obj.ActivityCount = Obj.ActivityCount - 1
                    If Obj.ActivityCount > 0 Then
                        ReDim Preserve Obj.Activities(1 To Obj.ActivityCount) As MyActivity
                    Else
                        Erase Obj.Activities
                    End If
                    DeleteActivity = True
                    Exit For
                End If
            Next
        End If
    End If
End Function

Public Function ValidActivity(ByRef Activity As MyActivity) As Boolean
    ValidActivity = (Activity.Identity <> "")
End Function

Public Sub CalculateActivity(ByRef Activity As MyActivity)

    If Activity.Friction <> 0 Then
        Activity.BurstRate = Activity.BurstRate - (Activity.BurstRate * Activity.Friction)
        If Activity.BurstRate < 0 Then
            Activity.BurstRate = 0
            Activity.Identity = ""
        End If
    End If
    If (Activity.BurstRate > 0.0001) Or (Activity.BurstRate < -0.0001) Then
        Activity.OffsetLoc.X = Activity.Direction.X * Activity.BurstRate
        Activity.OffsetLoc.Y = Activity.Direction.Y * Activity.BurstRate
        Activity.OffsetLoc.z = Activity.Direction.z * Activity.BurstRate
    Else
        Activity.BurstRate = 0
    End If
        
End Sub

Private Sub ApplyActivity(ByRef Obj As MyObject)
    If Obj.ActivityCount > 0 Then
        Dim m As Long
        For m = 1 To Obj.ActivityCount
            If ValidActivity(Obj.Activities(m)) Then
                'calc all activity states
                CalculateActivity Obj.Activities(m)
                'add then to object state
                D3DXVec3Add Obj.Motion.OffsetLoc, Obj.Motion.OffsetLoc, Obj.Activities(m).OffsetLoc
            End If
        Next
    End If

End Sub

Public Sub SetupActivity()
    Dim o As Long
    Dim a As Long
    Dim d As Boolean
    If Player.Object.ActivityCount > 0 Then
        a = 1
        Do While a <= Player.Object.ActivityCount
            If Player.Object.Activities(a).BurstRate = 0 Then
                DeleteActivity Player.Object, Player.Object.Activities(a).Identity
            Else
                a = a + 1
            End If
        Loop
    End If
    
    Player.Object.Motion.OffsetLoc = MakeVector(0, 0, 0)
    Do
        d = DeleteActivity(Player.Object, "")
    Loop Until (Not d)
    
    If Partner.Object.ActivityCount > 0 Then
        a = 1
        Do While a <= Partner.Object.ActivityCount
            If Partner.Object.Activities(a).BurstRate = 0 Then
                DeleteActivity Partner.Object, Partner.Object.Activities(a).Identity
            Else
                a = a + 1
            End If
        Loop
    End If
    
    
    Partner.Object.Motion.OffsetLoc = MakeVector(0, 0, 0)
    Do
        d = DeleteActivity(Partner.Object, "")
    Loop Until (Not d)
    If ObjectCount > 0 Then
        For o = 1 To ObjectCount
            If Objects(o).ActivityCount > 0 Then
                a = 1
                Do While a <= Objects(o).ActivityCount
                    If Objects(o).Activities(a).BurstRate = 0 Or Objects(o).Activities(a).Identity = "" Then
                        DeleteActivity Objects(o), Objects(o).Activities(a).Identity
                    Else
                        a = a + 1
                    End If
                Loop
            End If
            
            Objects(o).Motion.OffsetLoc = MakeVector(0, 0, 0)
            Do
                d = DeleteActivity(Objects(o), "")
            Loop Until (Not d)
        Next
    End If
        
End Sub

Public Sub ResetActivity(ByRef Obj As MyObject)
    Dim o As Long
    Dim d As Boolean
    
    Obj.Motion.OffsetLoc = MakeVector(0, 0, 0)
    Do Until Obj.ActivityCount = 0
        DeleteActivity Obj, Obj.Activities(1).Identity
    Loop
End Sub
Public Sub ResetAllActivities()
    Dim o As Long
    Dim d As Boolean
    
    Player.Object.Motion.OffsetLoc = MakeVector(0, 0, 0)
    Do Until Player.Object.ActivityCount = 0
        DeleteActivity Player.Object, Player.Object.Activities(1).Identity
    Loop
    
    Partner.Object.Motion.OffsetLoc = MakeVector(0, 0, 0)
    Do Until Partner.Object.ActivityCount = 0
        DeleteActivity Partner.Object, Partner.Object.Activities(1).Identity
    Loop
    If ObjectCount > 0 Then
        For o = 1 To ObjectCount
            Objects(o).Motion.OffsetLoc = MakeVector(0, 0, 0)
            Do Until Objects(o).ActivityCount = 0
                DeleteActivity Objects(o), Objects(o).Activities(1).Identity
            Loop
        Next
    End If
        
End Sub
Public Sub RenderActivity()
    
    Dim o As Long
    Dim d As Boolean
    ApplyActivity Player.Object
    
    Player.Object.Origin.X = Player.Object.Origin.X + Player.Object.Motion.OffsetLoc.X
    Player.Object.Origin.Y = Player.Object.Origin.Y + Player.Object.Motion.OffsetLoc.Y
    Player.Object.Origin.z = Player.Object.Origin.z + Player.Object.Motion.OffsetLoc.z
    
    ApplyActivity Partner.Object

    Partner.Object.Origin.X = Partner.Object.Origin.X + Partner.Object.Motion.OffsetLoc.X
    Partner.Object.Origin.Y = Partner.Object.Origin.Y + Partner.Object.Motion.OffsetLoc.Y
    Partner.Object.Origin.z = Partner.Object.Origin.z + Partner.Object.Motion.OffsetLoc.z
    
    If ObjectCount > 0 Then
        For o = 1 To ObjectCount
            ApplyActivity Objects(o)

            Objects(o).Origin.X = Objects(o).Origin.X + Objects(o).Motion.OffsetLoc.X
            Objects(o).Origin.Y = Objects(o).Origin.Y + Objects(o).Motion.OffsetLoc.Y
            Objects(o).Origin.z = Objects(o).Origin.z + Objects(o).Motion.OffsetLoc.z
        Next
    End If
    
End Sub