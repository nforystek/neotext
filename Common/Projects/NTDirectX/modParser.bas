Attribute VB_Name = "modParser"
Option Explicit

Public Enum Actions
    None = 0
    Directing = 1
    Rotating = 2
    Scaling = 4
    Script = 8
End Enum

Public Enum Moving
    None = 0
    Level = 1
    Flying = 2
    Falling = 4
    Stepping = 8
End Enum

Public Type MyActivity
    Identity As String
    Action As Actions
    OnEvent As String
    
    Reactive As Single
    Latency As Single
    Recount As Single

End Type

Public Type MyObject

    Identity As String
    Visible As Boolean

    States As Moving
    
    Activities() As MyActivity
    ActivityCount As Long

End Type

Public Type MyVariable
    Identity As String
    Value As Variant
    OnEdit As String
End Type

Public Type MyMethod
    Identity As String
    Script As String
End Type

Public Objects() As MyObject
Public ObjectCount As Long

Public Variables() As MyVariable
Public VariableCount As Long

Public Methods() As MyMethod
Public MethodCount As Long


Public Sub SwapActivity(ByRef val1 As MyActivity, ByRef val2 As MyActivity)
    Dim tmp As MyActivity
    tmp = val1
    val1 = val2
    val2 = tmp
End Sub

Public Function SetActivity(ByRef act As MyActivity, ByRef Action As Actions) As String
    act.Identity = Replace(modGuid.GUID, "-", "")
    act.Action = Action
End Function


Public Function ActivityExists(ByRef Obj As MyObject, ByVal MGUID As String) As Boolean
    Dim a As Long
    For a = 1 To Obj.ActivityCount
        If Obj.Activities(a).Identity = MGUID Then
            ActivityExists = True
            Exit Function
        End If
    Next
    ActivityExists = False
End Function

Public Function ValidActivity(ByRef Activity As MyActivity) As Boolean
    ValidActivity = (Activity.Identity <> "")
End Function

'Public Function CalculateActivity(ByRef Activity As MyActivity, ByRef Action As Actions) As D3DVECTOR
'
'    If (Activity.Action And Action) = Action Then
'        If Activity.Friction <> 0 Then
'            Activity.Emphasis = Activity.Emphasis - (Activity.Emphasis * Activity.Friction)
'            If Activity.Emphasis < 0 Then
'                Activity.Emphasis = 0
'                Activity.Identity = ""
'            End If
'        End If
'        If (Activity.Emphasis > 0.0001) Or (Activity.Emphasis < -0.0001) Then
'            CalculateActivity.X = Activity.Data.X * Activity.Emphasis
'            CalculateActivity.Y = Activity.Data.Y * Activity.Emphasis
'            CalculateActivity.z = Activity.Data.z * Activity.Emphasis
'        Else
'            Activity.Emphasis = 0
'        End If
'    End If
'
'End Function

'Private Sub ApplyActivity(ByRef Obj As MyObject)
'    Dim cnt As Long
'    Dim cnt2 As Long
'    Dim Offset As D3DVECTOR
'
'    If ((Not (Perspective = Spectator)) And (Obj.CollideObject = Player.Object.CollideObject)) Or (Not (Obj.CollideObject = Player.Object.CollideObject)) Then
'
'        If Obj.Gravitational Then
'            If Not Obj.States.OnLadder Then
'                If Obj.States.InLiquid Then
'                    D3DXVec3Add Obj.Direct, Obj.Direct, CalculateActivity(LiquidGravityDirect, Directing)
'                    D3DXVec3Add Obj.Twists, Obj.Twists, CalculateActivity(LiquidGravityRotate, Rotating)
'                    D3DXVec3Add Obj.Scalar, Obj.Scalar, CalculateActivity(LiquidGravityScaled, Scaling)
'                Else
'                    D3DXVec3Add Obj.Direct, Obj.Direct, CalculateActivity(GlobalGravityDirect, Directing)
'                    D3DXVec3Add Obj.Twists, Obj.Twists, CalculateActivity(GlobalGravityRotate, Rotating)
'                    D3DXVec3Add Obj.Scalar, Obj.Scalar, CalculateActivity(GlobalGravityScaled, Scaling)
'                End If
'            End If
'        End If
'    End If
'    If Obj.Effect = Collides.None Then
'        If Obj.ActivityCount > 0 Then
'            Dim a As Long
'            For a = 1 To Obj.ActivityCount
'                If ValidActivity(Obj.Activities(a)) Then
'                    D3DXVec3Add Obj.Direct, Obj.Direct, CalculateActivity(Obj.Activities(a), Directing)
''                    If (Not (Obj.Twists.X = 0)) And (Not (Obj.Twists.Y = 0)) And (Not (Obj.Twists.z = 0)) Then
''                        Debug.Print Obj.Twists.X & " " & Obj.Twists.Y & " " & Obj.Twists.z
''                    End If
'                    D3DXVec3Add Obj.Twists, Obj.Twists, CalculateActivity(Obj.Activities(a), Rotating)
'                    D3DXVec3Add Obj.Scalar, Obj.Scalar, CalculateActivity(Obj.Activities(a), Scaling)
'                End If
'            Next
'        End If
'    End If
'End Sub

'Public Sub ResetMotion()
'    Dim a As Long
'    Dim o As Long
'    Player.Object.Direct = MakeVector(0, 0, 0)
'    If ObjectCount > 0 Then
'        For o = 1 To ObjectCount
'            Objects(o).Direct = MakeVector(0, 0, 0)
'        Next
'    End If
'End Sub

Public Sub ClearActivities()
    Dim o As Long
    If ObjectCount > 0 Then
        For o = 1 To ObjectCount
            Do Until Objects(o).ActivityCount = 0
                DeleteActivity Objects(o), Objects(o).Activities(1).Identity
            Loop
        Next
    End If
End Sub

'Public Sub RenderActive()
'On Error GoTo ObjectError
'
'    Dim d As Boolean
'    Dim o As Long
'    Dim a As Long
'    Dim act As MyActivity
'    Dim trig As String
'    Dim line As String
'    Dim id As String
'
'    Do
'    Loop Until (Not DeleteActivity(Player.Object, ""))
'
'    If Player.Object.Visible Then
'        ApplyActivity Player.Object
'
'        If Player.Object.ActivityCount > 0 Then
'
'            a = 1
'            Do While a <= Player.Object.ActivityCount
'                If Player.Object.Activities(a).Reactive > -1 Then
'                    If (Timer - Player.Object.Activities(a).Latency) > Player.Object.Activities(a).Reactive Then
'                        Player.Object.Activities(a).Latency = Timer
'                        act = Player.Object.Activities(a)
'                        act.Emphasis = act.Initials
'                        DeleteActivity Player.Object, act.Identity
'                        If Not act.OnEvent = "" Then
'                            line = NextArg(act.OnEvent, ":")
'                            trig = RemoveArg(act.OnEvent, ":")
'                            If Left(Trim(trig), 1) = "<" Then
'                                id = RemoveQuotedArg(trig, "<", ">") & ","
'                                If ((InStr(id, Player.Object.Identity & ",") > 0) And (Player.Object.Identity <> "")) Or (id = ",") Then
'                                    ParseLand line, trig
'                                End If
'                            Else
'                                ParseLand line, trig
'                            End If
'                        End If
'                        If act.Recount > -1 Then
'                            If act.Recount > 0 Then
'                                act.Recount = act.Recount - 1
'                                AddActivity Player.Object, act.Action, act.Identity, act.Data, act.Emphasis, act.Friction, act.Reactive, act.Recount, act.OnEvent
'                            End If
'                        Else
'                            AddActivity Player.Object, act.Action, act.Identity, act.Data, act.Emphasis, act.Friction, act.Reactive, act.Recount, act.OnEvent
'                        End If
'                    End If
'                    a = a + 1
'                ElseIf Player.Object.Activities(a).Emphasis = 0 Then
'                    DeleteActivity Player.Object, Player.Object.Activities(a).Identity
'                Else
'                    a = a + 1
'                End If
'            Loop
'
'        End If
'    End If
'
'
'    If ObjectCount > 0 Then
'        For o = 1 To ObjectCount
'            Do
'            Loop Until (Not DeleteActivity(Objects(o), ""))
'
'            If Objects(o).Visible Then
'                ApplyActivity Objects(o)
'
'                If Objects(o).ActivityCount > 0 Then
'                    a = 1
'                    Do While a <= Objects(o).ActivityCount
'                        If Objects(o).Activities(a).Reactive > -1 Then
'                            If (Timer - Objects(o).Activities(a).Latency) > Objects(o).Activities(a).Reactive Then
'                                Objects(o).Activities(a).Latency = Timer
'                                act = Objects(o).Activities(a)
'                                act.Emphasis = act.Initials
'                                DeleteActivity Objects(o), act.Identity
'                                If Not act.OnEvent = "" Then
'                                    line = NextArg(act.OnEvent, ":")
'                                    trig = RemoveArg(act.OnEvent, ":")
'                                    If Left(Trim(trig), 1) = "<" Then
'                                        id = RemoveQuotedArg(trig, "<", ">") & ","
'                                        If ((InStr(id, Objects(o).Identity & ",") > 0) And (Objects(o).Identity <> "")) Or (id = ",") Then
'                                            ParseLand line, trig
'                                        End If
'                                    Else
'                                        ParseLand line, trig
'                                    End If
'                                End If
'                                If act.Recount > -1 Then
'                                    If act.Recount > 0 Then
'                                        act.Recount = act.Recount - 1
'                                        AddActivity Objects(o), act.Action, act.Identity, , act.Reactive, act.Recount, act.OnEvent
'                                    End If
'                                Else
'                                    AddActivity Objects(o), act.Action, act.Identity, , act.Reactive, act.Recount, act.OnEvent
'                                End If
'                            End If
'
'                            a = a + 1
'                        ElseIf Objects(o).Activities(a).Emphasis = 0 Then
'                            DeleteActivity Objects(o), Objects(o).Activities(a).Identity
'                        Else
'                            a = a + 1
'                        End If
'                    Loop
'                End If
'            End If
'
'        Next
'    End If
'
'
'    Exit Sub
'ObjectError:
'    If Err.Number = 6 Or Err.Number = 11 Then Resume
'    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
'    Resume
'End Sub

Public Function AddActivity(ByRef Obj As MyObject, ByRef Action As Actions, ByVal aGUID As String, Optional ByVal Reactive As Single = -1, Optional ByVal Recount As Single = -1, Optional Script As String = "") As String
    Obj.ActivityCount = Obj.ActivityCount + 1
    ReDim Preserve Obj.Activities(1 To Obj.ActivityCount) As MyActivity
    With Obj.Activities(Obj.ActivityCount)
        .Identity = IIf(aGUID = "", Replace(modGuid.GUID, "-", ""), aGUID)
        .Action = Action
        .Reactive = Reactive
        .Latency = Timer
        .Recount = Recount
        .OnEvent = Script
        AddActivity = .Identity
    End With
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

Public Function RemoveLineArg(ByRef TheParams As Variant, Optional ByVal EndOfLine As String = " ") As String
    If (InStr(TheParams, EndOfLine) > 0) And (InStr(TheParams, vbCrLf) > 1) Then
        If InStr(TheParams, EndOfLine) < InStr(TheParams, vbCrLf) Then
            RemoveLineArg = Left(TheParams, InStr(TheParams, EndOfLine) - 1)
            TheParams = Mid(TheParams, InStr(TheParams, EndOfLine))
        Else
            RemoveLineArg = Left(TheParams, InStr(TheParams, vbCrLf) - 1)
            TheParams = Mid(TheParams, InStr(TheParams, vbCrLf) + Len(vbCrLf))
        End If
        
    ElseIf (InStr(TheParams, vbCrLf) = 0) And (InStr(TheParams, EndOfLine) = 0) Then
        RemoveLineArg = TheParams
        TheParams = ""
    ElseIf (InStr(TheParams, vbCrLf) = 0) Then
        RemoveLineArg = Left(TheParams, InStr(TheParams, EndOfLine) - 1)
        TheParams = Mid(TheParams, InStr(TheParams, EndOfLine))
    Else
        RemoveLineArg = Left(TheParams, InStr(TheParams, vbCrLf) - 1)
        TheParams = Mid(TheParams, InStr(TheParams, vbCrLf) + Len(vbCrLf))
    End If
End Function

Public Sub ParseLine(ByRef inItem As String, ByRef inText As String, ByRef LineNumber As Long)
    If InStr(inText, vbCrLf) > 0 Then
        If InStr(inText, "[") < InStr(inText, vbCrLf) Then
            inItem = Left(inText, InStr(inText, "[") - 1)
            inText = Mid(inText, InStr(inText, "["))
        Else
        
        End If
    ElseIf InStr(inText, "[") > 0 Then
        inItem = Left(inText, InStr(inText, "[") - 1)
        inText = Mid(inText, InStr(inText, "["))
    Else
        inItem = inText
        inText = ""
    End If
End Sub
Public Function ParseLand(ByVal inLine As Long, ByVal inText As String) As String
On Error GoTo parseerror
    
    Dim r As Single
    Dim o As Long
    Dim i As Long
    Dim cnt As Long
    Dim cnt2 As Long

    Dim NumLines As Long
    NumLines = inLine + CountWord(inText, vbCrLf)
    
    Dim inArg() As String
    Dim inItem As String
    Dim inName As String
    Dim inData As String
    Dim inTrig As String
    
    Do Until inText = ""
    
        If (Left(Replace(Replace(inText, " ", ""), vbTab, ""), 1) = ";") Then
            RemoveNextArg inText, vbCrLf
        Else
            inItem = RemoveLineArg(inText, "{")
            If (Not (Trim(inItem) = "")) Then
                inLine = (NumLines - CountWord(inText, vbCrLf))
                Select Case inItem
                    Case "parse"
                        inData = RemoveQuotedArg(inText, "{", "}", True)
                        Do Until inData = ""
                            If (Left(Replace(Replace(inData, " ", ""), vbTab, ""), 1) = ";") Then
                                RemoveNextArg inData, vbCrLf
                            Else
                                inName = RemoveLineArg(inData, "[")
                                inLine = (NumLines - CountWord(inData & inText, vbCrLf))
                                If (Not (Trim(inName) = "")) Then
                                    inArg() = Split(Trim(inName), " ")
                                    ReDim Preserve inArg(0 To 1)
                                    Select Case inArg(0)
                                        Case "filename"
                                            If PathExists(AppPath & "Components\" & inArg(1), True) Then
                                                ParseLand 0, Replace(ReadFile(AppPath & "Components\" & inArg(1)), vbTab, "")
                                            Else
                                                'AddMessage "Invalid object file [" & AppPath & "Components\" & inArg(1) & "]"
                                            End If
                                        Case Else
                                            If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
                                                'AddMessage "Warning, Unknown Object at Line " & inLine
                                            End If
                                    End Select
                                End If
                            End If
                        Loop

                    Case "object"
                        inData = RemoveQuotedArg(inText, "{", "}", True)
                        Dim NewObj As MyObject
                        NewObj.Identity = ""
                        
                        Do Until NewObj.ActivityCount = 0
                            DeleteActivity NewObj, NewObj.Activities(1).Identity
                        Loop
                        
                        Do Until inData = ""
                            If (Left(Replace(Replace(inData, " ", ""), vbTab, ""), 1) = ";") Then
                                RemoveNextArg inData, vbCrLf
                            Else
                                inName = RemoveLineArg(inData, "[")
                                inLine = (NumLines - CountWord(inData & inText, vbCrLf))
                                If (Not (Trim(inName) = "")) Then
                                    inArg() = Split(Trim(inName), " ")
                                    ReDim Preserve inArg(0 To 9)
                                    Select Case inArg(0)
                                        Case "visible"
                                            If inArg(1) = "" Then
                                                NewObj.Visible = True
                                            Else
                                                NewObj.Visible = CBool(inArg(1))
                                            End If
                                            
                                        Case "filename", "boundsobj"
                                            
'                                            If MeshCount = 0 Then
'                                                MeshCount = MeshCount + 1
'                                                ReDim Meshes(1 To MeshCount) As MyMesh
'                                                NewObj.MeshIndex = MeshCount
'                                                Meshes(NewObj.MeshIndex).Filename = LCase(inArg(1))
'                                                If PathExists(AppPath & "Models\" & Meshes(NewObj.MeshIndex).Filename, True) Then
'                                                    CreateMesh AppPath & "Models\" & Meshes(NewObj.MeshIndex).Filename, Meshes(NewObj.MeshIndex).Mesh, Meshes(NewObj.MeshIndex).MaterialBuffer, _
'                                                            NewObj.Origin, NewObj.Scaled, Meshes(NewObj.MeshIndex).Materials, Meshes(NewObj.MeshIndex).Textures, _
'                                                            Meshes(NewObj.MeshIndex).Verticies, Meshes(NewObj.MeshIndex).Indicies, Meshes(NewObj.MeshIndex).MaterialCount
'                                                Else
'                                                    ReDim Meshes(NewObj.MeshIndex).Textures(0 To 0) As Direct3DTexture8
'                                                    ReDim Meshes(NewObj.MeshIndex).Materials(0 To 0) As D3DMATERIAL8
'                                                    NewObj.MeshIndex = 0
'                                                End If
'                                            Else
'                                                For i = LBound(Meshes) To UBound(Meshes)
'                                                    If Meshes(i).Filename = LCase(inArg(1)) Then
'                                                        NewObj.MeshIndex = i
'                                                        Exit For
'                                                    End If
'                                                Next
'
'                                                If NewObj.MeshIndex = 0 Then
'                                                    MeshCount = MeshCount + 1
'                                                    ReDim Preserve Meshes(1 To MeshCount) As MyMesh
'                                                    NewObj.MeshIndex = MeshCount
'                                                    Meshes(NewObj.MeshIndex).Filename = LCase(inArg(1))
'                                                    If PathExists(AppPath & "Models\" & Meshes(NewObj.MeshIndex).Filename, True) Then
'                                                        CreateMesh AppPath & "Models\" & Meshes(NewObj.MeshIndex).Filename, Meshes(NewObj.MeshIndex).Mesh, Meshes(NewObj.MeshIndex).MaterialBuffer, _
'                                                                NewObj.Origin, NewObj.Scaled, Meshes(NewObj.MeshIndex).Materials, Meshes(NewObj.MeshIndex).Textures, _
'                                                                Meshes(NewObj.MeshIndex).Verticies, Meshes(NewObj.MeshIndex).Indicies, Meshes(NewObj.MeshIndex).MaterialCount
'                                                    Else
'                                                        ReDim Meshes(NewObj.MeshIndex).Textures(0 To 0) As Direct3DTexture8
'                                                        ReDim Meshes(NewObj.MeshIndex).Materials(0 To 0) As D3DMATERIAL8
'                                                        NewObj.MeshIndex = 0
'                                                    End If
'
'                                                End If
'
'                                            End If

                                        Case "activity"

                                            Select Case LCase(CStr(inArg(1)))
                                                Case "direct"
                                                    AddActivity NewObj, Actions.Directing, inArg(2), IIf(IsNumeric(inArg(7)), inArg(7), 0), _
                                                                            IIf(IsNumeric(inArg(8)), inArg(8), -1), IIf(IsNumeric(inArg(9)), inArg(9), -1)
                                                Case "rotate"
                                                    
                                                    AddActivity NewObj, Actions.Rotating, inArg(2), IIf(IsNumeric(inArg(7)), inArg(7), 0), _
                                                                            IIf(IsNumeric(inArg(8)), inArg(8), -1), IIf(IsNumeric(inArg(9)), inArg(9), -1)
                                                Case "scale"
                                                    AddActivity NewObj, Actions.Scaling, inArg(2), IIf(IsNumeric(inArg(7)), inArg(7), 0), _
                                                                            IIf(IsNumeric(inArg(8)), inArg(8), -1), IIf(IsNumeric(inArg(9)), inArg(9), -1)
                                                Case "script"
                                                    If InStr(inData, "[") > 0 Then
                                                        inLine = inLine + CountWord(RemoveLineArg(inData, "["), vbCrLf)
                                                        inTrig = RemoveQuotedArg(inData, "[", "]", True)
                                                        inTrig = inLine & ":" & inTrig
                                                        
                                                        AddActivity NewObj, Actions.Script, inArg(2), IIf(IsNumeric(inArg(3)), inArg(3), -1), _
                                                                                IIf(IsNumeric(inArg(4)), inArg(4), -1), inTrig
                                                    Else
                                                        'AddMessage "Warning, Brackets Required at Line " & inLine
                                                    End If
        
                                            End Select
        
                                        Case Else
                                            If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
                                                'AddMessage "Warning, Unknown Object at Line " & inLine
                                            End If
                                    End Select
                                ElseIf Left(Trim(inData), 1) = "[" Then
                                    'AddMessage "Warning, Itemless Brackets at Line " & inLine
                                    GoTo throwerror
                                End If
                            End If
                        Loop

                        ObjectCount = ObjectCount + 1
                        ReDim Preserve Objects(1 To ObjectCount) As MyObject
                        Objects(ObjectCount) = NewObj
    
                    Case "method"
                        inData = RemoveQuotedArg(inText, "{", "}", True)
                        MethodCount = MethodCount + 1
                        ReDim Preserve Methods(1 To MethodCount) As MyMethod
                        With Methods(MethodCount)
                            Do Until inData = ""
                                If (Left(Replace(Replace(inData, " ", ""), vbTab, ""), 1) = ";") Then
                                    RemoveNextArg inData, vbCrLf
                                Else
                                    inName = RemoveLineArg(inData, "[")
                                    inLine = (NumLines - CountWord(inData & inText, vbCrLf))
                                    If (Not (Trim(inName) = "")) Then
                                        inArg() = Split(Trim(inName), " ")
                                        ReDim Preserve inArg(0 To 1)
                                        Select Case inArg(0)
                                            Case "identity"
                                                .Identity = inArg(1)
                                            Case "script"
                                                If InStr(inData, "[") > 0 Then
                                                    inLine = inLine + CountWord(RemoveLineArg(inData, "["), vbCrLf)
                                                    .Script = RemoveQuotedArg(inData, "[", "]", True)
                                                    .Script = inLine & ":" & .Script
                                                Else
                                                    'AddMessage "Warning, Brackets Required at Line " & inLine
                                                End If
                                            Case Else
                                                If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
                                                    'AddMessage "Warning, Unknown Object at Line " & inLine
                                                End If
                                        End Select
                                    ElseIf Left(Trim(inData), 1) = "[" Then
                                        'AddMessage "Warning, Itemless Brackets at Line " & inLine
                                        GoTo throwerror
                                    End If
                                End If
                            Loop
                        End With
                        
                        
                    Case "variable"
                        inData = RemoveQuotedArg(inText, "{", "}", True)
                        VariableCount = VariableCount + 1
                        ReDim Preserve Variables(1 To VariableCount) As MyVariable
                        With Variables(VariableCount)
                            Do Until inData = ""
                                If (Left(Replace(Replace(inData, " ", ""), vbTab, ""), 1) = ";") Then
                                    RemoveNextArg inData, vbCrLf
                                Else
                                    inName = RemoveLineArg(inData, "[")
                                    inLine = (NumLines - CountWord(inData & inText, vbCrLf))
                                    If (Not (Trim(inName) = "")) Then
                                        inArg() = Split(Trim(inName), " ")
                                        ReDim Preserve inArg(0 To 1)
                                        Select Case inArg(0)
                                            Case "identity"
                                                .Identity = inArg(1)
                                            Case "value"
                                                If InStr(inArg(1), """") > 0 Then
                                                    .Value = CVar(RemoveQuotedArg(inName, """", """"))
                                                Else
                                                     .Value = CVar(inArg(1))
                                                End If
                                            Case "onedit"
                                                If InStr(inData, "[") > 0 Then
                                                    inLine = inLine + CountWord(RemoveLineArg(inData, "["), vbCrLf)
                                                    .OnEdit = RemoveQuotedArg(inData, "[", "]", True)
                                                    .OnEdit = inLine & ":" & .OnEdit
                                                Else
                                                    'AddMessage "Warning, Brackets Required at Line " & inLine
                                                End If
                                            Case Else
                                                If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
                                                    'AddMessage "Warning, Unknown Object at Line " & inLine
                                                End If
                                        End Select
                                    ElseIf Left(Trim(inData), 1) = "[" Then
                                        'AddMessage "Warning, Itemless Brackets at Line " & inLine
                                        GoTo throwerror
                                    End If
                                End If
                            Loop
                        End With
                        
'                    Case "database"
'                        inData = RemoveQuotedArg(inText, "{", "}", True)
'                        Do Until inData = ""
'                            If (Left(Replace(Replace(inData, " ", ""), vbTab, ""), 1) = ";") Then
'                                RemoveNextArg inData, vbCrLf
'                            Else
'                                inName = RemoveLineArg(inData, "[")
'                                inLine = (NumLines - CountWord(inData & inText, vbCrLf))
'                                If (Not (Trim(inName) = "")) Then
'                                    inArg() = Split(Trim(inName), " ")
'                                    ReDim Preserve inArg(0 To 0)
'                                    Select Case inArg(0)
'                                        Case "deserialize"
'                                            If InStr(inData, "[") > 0 Then
'                                                inLine = inLine + CountWord(RemoveLineArg(inData, "["), vbCrLf)
'                                                Deserialize = RemoveQuotedArg(inData, "[", "]", True)
'                                                Deserialize = inLine & ":" & Deserialize
'                                            Else
'                                                'AddMessage "Warning, Brackets Required at Line " & inLine
'                                            End If
'                                        Case "serialize"
'                                            If InStr(inData, "[") > 0 Then
'                                                inLine = inLine + CountWord(RemoveLineArg(inData, "["), vbCrLf)
'                                                Serialize = RemoveQuotedArg(inData, "[", "]", True)
'                                                Serialize = inLine & ":" & Serialize
'                                            Else
'                                                'AddMessage "Warning, Brackets Required at Line " & inLine
'                                            End If
'                                        Case Else
'                                            If (Not (Left(Replace(Replace(inArg(0), " ", ""), vbTab, ""), 1) = ";")) Then
'                                                'AddMessage "Warning, Unknown Object at Line " & inLine
'                                            End If
'                                    End Select
'                                ElseIf Left(Trim(inData), 1) = "[" Then
'                                    'AddMessage "Warning, Itemless Brackets at Line " & inLine
'                                    GoTo throwerror
'                                End If
'                            End If
'                        Loop
                    Case Else
                        
                        Dim inVar As String
                        Dim inVal As Variant
                        inVal = RemoveArg(Trim(inItem), " ")
                        inVar = NextArg(Trim(inItem), " ")
                        If Left(inVar, 1) = ";" Then
                            inName = RemoveNextArg(inText, vbCrLf)
                            inLine = (NumLines - CountWord(inData & inText, vbCrLf))
                        ElseIf Left(inVar, 1) = "!" Then
                            Process Mid(inVar, 2) & " " & NextArg(RemoveArg(inItem, " "), vbCrLf)
                        ElseIf Left(inVar, 1) = "=" Then
                            If inVal = "" Then inVal = Mid(inVar, 2)
                            ParseLand = ParseLand & Mid(inVar, 2) & " " & ParseSetGet(inLine, inVal) & vbCrLf
                        ElseIf Left(inVar, 1) = "$" Then
                            ParseSetGet inLine, inVar, ParseExpr(inLine, inVal)
                        ElseIf Left(inVar, 1) = "&" Then
                            If (MethodCount > 0) Then
                                If InStr(Mid(inVar, 2), ".") > 0 Then
                                    inVar = ParseSetGet(inLine, inVar, inVal)
                                    'inVal = RemoveArg(inVar, ".")
                                    'inVar = NextArg(inVar, ".")
                                Else
                                    For cnt = 1 To MethodCount
                                        If LCase(Methods(cnt).Identity) = LCase(Mid(inVar, 2)) Then
                                            inVal = Methods(cnt).Script
                                            If Not (inVal = "") Then
                                                ParseLand NextArg(inVal, ":"), RemoveArg(inVal, ":")
                                            End If
                                            inVar = ""
                                        End If
                                    Next
                                End If
                            End If
                            
                            If Not inVar = "" Then
                                'AddMessage "Warning, Unknown Method at Line " & inLine
                            End If
                            
                        ElseIf inVar = "if" Then
                            inText = inVar & " " & inVal & inText
                            inLine = (NumLines - CountWord(inText, vbCrLf))
    
                            Dim ifCode As String
                            Dim inExp As String
                            Dim inIs As Variant
                            Dim elseCode As String
                            Dim useexp As Boolean
                            Dim lnum As Long
                            Dim calls As String
                            inName = RemoveLineArg(inText, "[")
                            inName = Trim(Mid(Trim(inName), 3))
                            
                            If Left(Replace(Replace(Replace(inName, vbTab, ""), vbCrLf, ""), " ", ""), 1) = "(" Then
                                inExp = RemoveQuotedArg(inName, "(", ")", True)
                                If LCase(Left(Trim(inName), 2)) = "is" Then
                                    inName = Trim(Mid(Trim(inName), 3))
                                    If Left(Replace(Replace(Replace(inName, vbTab, ""), vbCrLf, ""), " ", ""), 1) = "(" Then
                                        inIs = RemoveQuotedArg(inName, "(", ")", True)
                                    ElseIf Not (Left(Replace(Trim(inVar), vbTab, ""), 1) = ";") Then
                                        'AddMessage "Error, Is Expression at Line " & inLine
                                    End If
                                Else
                                    inIs = True
                                End If
                                inData = RemoveQuotedArg(inText, "[", "]", True)
                                
                                If ParseExpr(inLine, inExp) = ParseSetGet(inLine, inIs) Then
                                
                                    Do While Left(Trim(inText), 2) = vbCrLf
                                        inText = Mid(Trim(inText), 3)
                                    Loop
                                    inText = Trim(inText)
                                    inLine = (NumLines - CountWord(inText, vbCrLf))
                                    
                                    useexp = True
                                    calls = inLine & ":" & inData
    
                                Else
    
                                    Do While Left(Trim(inText), 2) = vbCrLf
                                        inText = Trim(Mid(Trim(inText), 3))
                                    Loop
                                    inText = Trim(inText)
                                    inLine = (NumLines - CountWord(inText, vbCrLf))
                                        
                                    Do While (NextArg(inText, vbCrLf) = "elseif" Or NextArg(inText, " ") = "elseif" Or NextArg(inText, "(") = "elseif") And (Not useexp)
                                        inName = RemoveLineArg(inText, "[")
                                        inName = Trim(Mid(Trim(inName), 7))
    
                                        If Left(Replace(Replace(Replace(inName, vbTab, ""), vbCrLf, ""), " ", ""), 1) = "(" Then
                                            inExp = RemoveQuotedArg(inName, "(", ")", True)
                                            If LCase(Left(Trim(inName), 2)) = "is" Then
                                                inName = Trim(Mid(Trim(inName), 3))
                                                If Left(Replace(Replace(Replace(inName, vbTab, ""), vbCrLf, ""), " ", ""), 1) = "(" Then
                                                    inIs = RemoveQuotedArg(inName, "(", ")", True)
                                                Else
                                                    'AddMessage "Error, Is Expression at Line " & inLine
                                                End If
                                            Else
                                                inIs = True
                                            End If
                                            inData = RemoveQuotedArg(inText, "[", "]", True)
                                        
                                            If ParseExpr(inLine, inExp) = ParseSetGet(inLine, inIs) Then
                                                useexp = True
                                                calls = inLine & ":" & inData
                                            End If
                                            
                                            Do While Left(Trim(inText), 2) = vbCrLf
                                                inText = Trim(Mid(Trim(inText), 3))
                                            Loop
                                            inText = Trim(inText)
                                            inLine = (NumLines - CountWord(inText, vbCrLf))
                                        Else
                                            'AddMessage "Error, If Expression at Line " & inLine
                                        End If
                                    Loop
                                    If Not useexp Then
    
                                        Do While Left(Trim(inText), 2) = vbCrLf
                                            inText = Trim(Mid(Trim(inText), 3))
                                        Loop
                                        inText = Trim(inText)
                                        inLine = (NumLines - CountWord(inText, vbCrLf))
                                            
                                        If NextArg(inText, vbCrLf) = "else" Or NextArg(inText, " ") = "else" Or NextArg(inText, "[") = "else" Then
                                            inName = RemoveLineArg(inText, "[")
    
                                            inName = Trim(Mid(Trim(inName), 5))
                                            Do While Left(Trim(inText), 2) = vbCrLf
                                                inText = Trim(Mid(Trim(inText), 3))
                                                inLine = inLine + 1
                                            Loop
                                            inText = Trim(inText)
                                            
                                            inData = RemoveQuotedArg(inText, "[", "]", True)
                                            calls = inLine & ":" & inData
    
                                        End If
                                    End If
                                    
                                End If
                                
                                If calls <> "" Then
                                    ParseLand NextArg(calls, ":"), RemoveArg(calls, ":")
                                End If
    
                                Do While Left(Trim(inText), 2) = vbCrLf
                                    inText = Trim(Mid(Trim(inText), 3))
                                Loop
                                inText = Trim(inText)
                                    
                                If useexp And ((NextArg(inText, vbCrLf) = "elseif" Or NextArg(inText, " ") = "elseif" Or NextArg(inText, "(") = "elseif") Or (NextArg(inText, vbCrLf) = "else" Or NextArg(inText, " ") = "else" Or NextArg(inText, "[") = "else")) Then
    
                                    Do While (NextArg(inText, vbCrLf) = "elseif" Or NextArg(inText, " ") = "elseif" Or NextArg(inText, "(") = "elseif")
                                        inName = RemoveLineArg(inText, "[")
                                        inName = Trim(Mid(Trim(inName), 7))
                                        If Left(Replace(Replace(Replace(inName, vbTab, ""), vbCrLf, ""), " ", ""), 1) = "(" Then
                                            inExp = RemoveQuotedArg(inName, "(", ")", True)
                                            If LCase(Left(Trim(inName), 2)) = "is" Then
                                                inName = Trim(Mid(Trim(inName), 3))
                                                If Left(Replace(Replace(Replace(inName, vbTab, ""), vbCrLf, ""), " ", ""), 1) = "(" Then
                                                    inIs = RemoveQuotedArg(inName, "(", ")", True)
                                                Else
                                                    'AddMessage "Error, Is Expression at Line " & inLine
                                                End If
                                            Else
                                                inIs = True
                                            End If
                                            inData = RemoveQuotedArg(inText, "[", "]", True)
                                            
                                            Do While Left(Trim(inText), 2) = vbCrLf
                                                inText = Trim(Mid(Trim(inText), 3))
                                            Loop
                                            inText = Trim(inText)
                                            
                                        Else
                                            'AddMessage "Error, If Expression at Line " & inLine
                                        End If
                                    Loop
                                    
                                    Do While Left(Trim(inText), 2) = vbCrLf
                                        inText = Trim(Mid(Trim(inText), 3))
                                    Loop
                                    inText = Trim(inText)
                                            
                                    If (NextArg(inText, vbCrLf) = "else" Or NextArg(inText, " ") = "else" Or NextArg(inText, "[") = "else") Then
                                        inName = RemoveLineArg(inText, "[")
    
                                        inName = Trim(Mid(Trim(inName), 5))
                                        inData = RemoveQuotedArg(inText, "[", "]", True)
    
                                    End If
        
                                End If
                            Else
                                'AddMessage "Error, If Expression at Line " & inLine
                            End If
                            
                        Else
                            'AddMessage "Warning, Unknown Object at Line " & inLine
                        End If
                End Select
            End If
        End If
    Loop
    
    Exit Function
parseerror:
    'AddMessage "Script Error at Line " & inLine
throwerror:
'    If Not ConsoleVisible Then ConsoleToggle
    Err.Clear
End Function

Public Function ParseExpr(ByVal inLine As Long, ByVal inExp As String) As Variant
    Dim exp As Variant
    Dim opr As String
    Dim val As Variant
    
    If InStr(inExp, "(") > 0 Then
        opr = Left(inExp, InStr(inExp, "(") - 1)
        opr = opr & " " & RemoveQuotedArg(inExp, "(", ")", True)
        inLine = inLine + CountWord(opr, vbCrLf)
        inExp = opr & " " & inExp
        opr = ""
    End If
    
    Do
        
        If opr = "" Then opr = RemoveNextArg(inExp, " ")
        If opr = "not" Then
            val = RemoveNextArg(inExp, " ")
            exp = Not ParseSetGet(inLine, val)
        Else
            Select Case LCase(opr)
                Case "or"
                    val = RemoveNextArg(inExp, " ")
                    exp = exp Or ParseSetGet(inLine, val)
                Case "and"
                    val = RemoveNextArg(inExp, " ")
                    exp = exp And ParseSetGet(inLine, val)
                Case Else
                    exp = ParseSetGet(inLine, opr)
            End Select
        End If
    
        opr = RemoveNextArg(inExp, " ")
        
    Loop Until inExp = ""
        
    ParseExpr = exp
End Function
Public Function ParseValues(ByVal inText As String) As String
    Dim outText As String
    
    Do Until inText = ""
        If InStr(inText, "$") > 0 Then
            outText = outText & Left(inText, InStr(inText, "$") - 1)
            inText = Mid(inText, InStr(inText, "$"))
            If InStr(inText, " ") > 0 Then
                outText = outText & ParseSetGet(0, Left(inText, InStr(inText, " ") - 1))
                inText = Mid(inText, InStr(inText, " ") + 1)
            Else
                outText = outText & ParseSetGet(0, inText)
                inText = ""
            End If
        Else
            outText = outText & inText
            inText = ""
        End If
    Loop
    ParseValues = outText
End Function
Public Function ParseSetGet(ByVal inLine As Long, ByVal inItem As Variant, Optional ByVal SetValue As Variant = Empty) As Variant
    
    If ((Left(Trim(inItem), 1) = "$") Or (Left(Trim(inItem), 1) = "&")) And InStr(inItem, ".") > 0 Then
        
        ParseSetGet = SetValue
        
        Dim inProp As String
        Dim cnt As Long
        Dim cnt2 As Long
    
        inProp = Trim(Mid(inItem, InStr(inItem, ".") + 1))
        inItem = Mid(Left(inItem, InStr(inItem, ".") - 1), 2)

        If (VariableCount > 0) And (Not inItem = "$") Then
            For cnt = 1 To VariableCount
                If LCase(Variables(cnt).Identity) = LCase(inItem) Then
                    Select Case inProp
                        Case "value"
                            If SetValue = Empty Then
                                ParseSetGet = CVar(Variables(cnt).Value)
                            Else
                                Variables(cnt).Value = CVar(SetValue)
                                If Not Variables(cnt).OnEdit = "" Then
                                    ParseLand NextArg(Variables(cnt).OnEdit, ":"), RemoveArg(Variables(cnt).OnEdit, ":")
                                End If
                            End If
                        Case Else
                            'AddMessage "Warning, Unknown Sub Entity " & inItem
                    End Select
                    inItem = "$"
                End If
            Next
        End If
      
        If (ObjectCount > 0) And (Not inItem = "$") Then
            For cnt = 1 To ObjectCount
                If LCase(Objects(cnt).Identity) = LCase(inItem) Then
                    Select Case inProp
                        Case "visible"
                            If SetValue = Empty Then
                                ParseSetGet = CBool(Objects(cnt).Visible)
                            Else
                                Objects(cnt).Visible = CBool(SetValue)
                            End If
                        Case Else
                            'AddMessage "Warning, Unknown Sub Entity " & inItem
                    End Select
                    inItem = "$"
                End If
                If Objects(cnt).ActivityCount And (Not inItem = "$") Then
                    For cnt2 = 1 To Objects(cnt).ActivityCount
                        If LCase(Objects(cnt).Activities(cnt2).Identity) = LCase(inItem) Then
                            Select Case inProp
                                Case Else
                                    'AddMessage "Warning, Unknown Sub Entity " & inItem
                            End Select
                            inItem = "$"
                        End If
                    Next
                End If
            Next
        End If
                
        If Not inItem = "$" Then
            'AddMessage "Warning, Unkown Identity " & inItem
        End If
    Else
        ParseSetGet = inItem
    End If
End Function

Public Sub Process(ByVal inArg As String)

    Dim o As Long
    Dim l As Long
    Dim cnt As Long
    Dim inNew As String
    Dim inTmp As String
    Dim inX As Single
    Dim inY As Single
    
    Dim inCmd As String
    
    inCmd = RemoveNextArg(inArg, " ")
    If Left(inCmd, 1) = "/" Then inCmd = Mid(inCmd, 2)
    
    Select Case LCase(inCmd)
        Case "debug"
    
    
    End Select
End Sub


