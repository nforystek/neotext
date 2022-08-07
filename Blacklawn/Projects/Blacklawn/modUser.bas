#Const [True] = -1
#Const [False] = 0

Attribute VB_Name = "modUser"
#Const modUser = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module

Public Score1 As Long
Public Score2 As Long
Public Score3 As Long
Public ScoreXth As String
Public Score1st As String
Public Score2nd As String
Public ScoreClr As Long
Public ScoreFor As String
Public ScoreFith As String
Public ScoreIdol As String
Public ScoreNth As Long
Public GainMode As Long
Public SakeTime As String
Public MuseTime As String
Public WonTime As Long
Public TooTime As Long
Public ClrTime As Double
Public XthTime As String
Public TriTime As Long
Public ForTime As String
Public FithTime As String
Public SixTime As Long
Public NthTime As Long
Public Bodykits As String
Public HandiCap As Boolean
'Public PortalHold As Long


Public Sub SetGainMode0()
                    
    Player.Gravity = 0
    
    If ((Not Player.Stalled) Or HandiCap) Then
        If ((Not (Player.Object.Origin.X = 0)) Or (Not (Player.Object.Origin.Y = 0)) Or (Not (Player.Object.Origin.Y = 0))) Then
            ResetActivity Player.Object
            Player.Object.Origin.X = 0
            Player.Object.Origin.Y = 0
            Player.Object.Origin.z = 0
            If Clocker.FollowingMode Then
                ResetActivity Partner.Object
                Partner.Object.Origin.X = 0
                Partner.Object.Origin.Y = 0
                Partner.Object.Origin.z = 0
            End If
        End If
    End If
    
    If (GainMode >= 1) Or HandiCap Then
        If ((Not HandiCap) And (Not Player.Stalled)) Or (Player.Stalled And (Not HandiCap)) Then
            GainMode = 0
            TriTime = 0
        End If
        FadeMessage "You have forfeit your score and no longer have a legit arrival time."
    End If
    
    If HandiCap Then
        HandiCap = False
        Player.Stalled = False
    End If
End Sub

Public Sub SetGainMode1()
    If (GainMode = 0) Or HandiCap Then

        If (Not Player.Stalled) Or HandiCap Then
            If (Not HandiCap) Then
                ResetActivity Player.Object
                Player.Object.Origin.X = IIf((Rnd < 0.5), -RandomPositive(FadeDistance, BlackBoundary), RandomPositive(FadeDistance, BlackBoundary))
                Player.Object.Origin.z = IIf((Rnd < 0.5), -RandomPositive(FadeDistance, BlackBoundary), RandomPositive(FadeDistance, BlackBoundary))
                If Clocker.FollowingMode Then
                    ResetActivity Partner.Object
                    Partner.Object.Origin.X = Player.Object.Origin.X
                    Partner.Object.Origin.Y = Player.Object.Origin.Y
                    Partner.Object.Origin.z = Player.Object.Origin.z
                End If

            End If
        End If
        
        If ((Not HandiCap) And (Not Player.Stalled)) Or (Player.Stalled And (Not HandiCap)) Then
            MuseTime = GetNow
            GainMode = 1
        End If
        
        If HandiCap Then
            HandiCap = False
            Player.Stalled = False
        End If
        FadeMessage "The pitch black clock has begun, you either return for your score or give up zero."
    End If
End Sub

Public Sub SetGainMode2()
    If (GainMode = 0) Or HandiCap Then
        
        If (Not Player.Stalled) Or HandiCap Then
            ResetActivity Player.Object
            Player.Object.Origin.X = 0
            Player.Object.Origin.z = 0
            If Clocker.FollowingMode Then
                ResetActivity Partner.Object
                Partner.Object.Origin.X = Player.Object.Origin.X
                Partner.Object.Origin.Y = Player.Object.Origin.Y
                Partner.Object.Origin.z = Player.Object.Origin.z
            End If

        End If

        If ((Not HandiCap) And (Not Player.Stalled)) Or (Player.Stalled And (Not HandiCap)) Then
            SakeTime = GetNow
            GainMode = 2
        End If
        
        If HandiCap Then
            HandiCap = False
            Player.Stalled = False
        End If
        FadeMessage "Penalty clock started, run from the streets far enough to start the pitch black clock."
    End If
End Sub

Public Sub ScoreUser()
    Dim cnt As Long
    Dim dis As Long
        
    Dim fiv As Long
    Dim vif As Long
    
    If (GainMode = 1) Then
        fiv = fiv + 1
    ElseIf (GainMode = 3) Then
        fiv = fiv + 2
    End If
    
    If GainMode = 2 Then
        dis = Distance(Player.Object.Origin.X, Player.Object.Origin.Y, Player.Object.Origin.z, 0, 0, 0)
        If (dis > FadeDistance) Then
            GainMode = 3
            
            WonTime = DateDiff("s", SakeTime, GetNow)
            
            MuseTime = GetNow
            
            FadeMessage "The pitch black clock has begun, you either return for your score or give up zero."
            
            vif = vif + 1
        End If
        
    ElseIf GainMode = 1 Or GainMode = 3 Then

        dis = Distance(Player.Object.Origin.X, Player.Object.Origin.Y, Player.Object.Origin.z, 0, 0, 0)
        If (dis < ZoneDistance) Then

            If GainMode = 1 Then
                TooTime = DateDiff("s", MuseTime, GetNow)
                Score1st = "1st " & TooTime & "s"
                
                If (Not frmMain.Recording) And (Not frmMain.Multiplayer) Then db.dbQuery "UPDATE Settings SET Score1st='" & Score1st & "' WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"
                
            ElseIf GainMode = 3 Then
                TooTime = DateDiff("s", MuseTime, GetNow)
                Score2nd = "2nd " & WonTime & "s " & TooTime & "s"

                If (Not frmMain.Recording) And (Not frmMain.Multiplayer) Then db.dbQuery "UPDATE Settings SET Score2nd='" & Score2nd & "' WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"
                
            End If
            GainMode = 0
            FadeMessage "Positive pitch black issue on the black lawn, well done."
            
            vif = vif + 1
        End If
        
    End If

    If (Not IsDate(XthTime)) Or (XthTime = "") Then
        If (Player.Object.Origin.Y = 0) Then
            If (Player.Object.Origin.X >= -3075 And Player.Object.Origin.X <= -2067) And (Player.Object.Origin.z >= -2565 And Player.Object.Origin.z <= -1029) Then
                XthTime = GetNow
                FadeMessage "Warning: No parking, no dumping, and no skateboarding in this lot..."
            End If
        End If
    End If
    
    If ScoreClr = 0 Then
        Player.Texture = 1
    Else
        Player.Texture = ScoreClr
    End If
    
    If (Player.Object.Origin.Y = 0) Then
        If (Player.Object.Origin.X >= -97 And Player.Object.Origin.X <= -78) And (Player.Object.Origin.z >= -2436 And Player.Object.Origin.z <= -2415) Then
        
            If (ClrTime = 0) Then ClrTime = GetTimer

            If CDbl(GetTimer - ClrTime) >= 2 Then
                ClrTime = GetTimer
                
                If ScoreClr = 0 Then
                    FadeMessage "You found the spray paint shop ability and gained the 3rd score award!"
                    ScoreClr = 1
                End If
                
                If ScoreClr = 1 Or ScoreClr = 0 Then
                    ScoreClr = 2
                ElseIf ScoreClr = 2 Then
                    ScoreClr = 3
                ElseIf ScoreClr = 3 Then
                    ScoreClr = 4
                ElseIf ScoreClr = 4 Then
                    ScoreClr = 5
                ElseIf ScoreClr = 5 Then
                    ScoreClr = 6
                ElseIf ScoreClr = 6 Then
                    ScoreClr = 7
                ElseIf ScoreClr = 7 Then
                    ScoreClr = 8
                ElseIf ScoreClr = 8 Then
                    ScoreClr = 9
                ElseIf ScoreClr = 9 Then
                    ScoreClr = 1
                End If
                
                If (Not frmMain.Recording) And (Not frmMain.Multiplayer) Then db.dbQuery "UPDATE Settings SET ScoreClr=" & ScoreClr & " WHERE Username='" & Replace(GetUserLoginName, "'", "''") & "';"
            End If
            
        Else
            ClrTime = 0
        End If
    Else
        ClrTime = 0
    End If
    
    Static DoMsg As Boolean
    If ((Player.Object.Origin.X >= 290 And Player.Object.Origin.X <= 443) And (Player.Object.Origin.z >= -1588 And Player.Object.Origin.z <= -1549)) And (Player.Object.Origin.Y = 0) Then
        If Not DoMsg Then
            DoMsg = True
            CenterText = "News: Explorers of this city speak of a scores above the 1st and 2nd!" & vbCrLf & _
                        "The news seems to think this city is for humans, and maybe some dogs." & vbCrLf & _
                        vbCrLf & _
                        "The Xth score is self expression for being AFK and may be exclusive." & vbCrLf & _
                        "The third score indicates various modifications to your player ship." & vbCrLf & _
                        "The fourth score is similar to the scoring of the 1st or 2nd score." & vbCrLf & _
                        "The fith score is a cumulative strand of simultanious timed scores." & vbCrLf & _
                        "The sixth score is letters and numbers for visited idol achievements." & vbCrLf & _
                        "The Nth score value is a accumulation of how many nickels you pick up." & vbCrLf & _
                        vbCrLf & _
                        "Everything is wrong, there's not much here to score with except time." & vbCrLf & _
                        "See ""How to Play"" in the Help Documentation for detailed information." & vbCrLf
        End If
    ElseIf DoMsg Then
        CenterText = ""
        DoMsg = False
    End If
    
    If (Not (TriTime = 3)) Then

        If (Player.Object.Origin.Y = 0) And ((TriTime = 0) Or (TriTime = 2) Or (TriTime > 3)) Then
        
            If ((Player.Object.Origin.X >= -1943 And Player.Object.Origin.X <= -1864) And (Player.Object.Origin.z >= 2352 And Player.Object.Origin.z <= 2560)) Then

                If (TriTime = 0) Or (TriTime = 2) Then
                    FadeMessage "Penalty clock started, run from the streets far enough to start the pitch black clock."
                    vif = vif + 1
                End If
                
                TriTime = 1
                ForTime = GetNow
                Player.Gravity = 1
                Player.Object.Origin.Y = -1
                
                If (Not frmMain.Recording) And (Not frmMain.Multiplayer) Then db.dbQuery "UPDATE Settings SET ScoreFor='" & ScoreFor & "' WHERE Username='" & Replace(GetUserLoginName, "'", "''") & "';"

            End If
            
        ElseIf ((Player.Object.Origin.Y < 0) And (Distance(Player.Object.Origin.X, Player.Object.Origin.Y, Player.Object.Origin.z, Player.Object.Origin.X, 0, Player.Object.Origin.z) > 200000)) And (TriTime = 1) Then
            TriTime = 3
            ForTime = GetNow
            ScoreFor = Replace(Replace(Replace(Replace(ScoreFor, "/", ""), "\", ""), "|", ""), "-", "")
            
            If (Not Player.Stalled) Then
                ResetActivity Player.Object
                Player.Object.Origin.Y = 50000
                Select Case SundialAim
                    Case 0
                        Player.Object.Origin.X = 0
                        Player.Object.Origin.z = 50000
                    Case 90
                        Player.Object.Origin.X = 50000
                        Player.Object.Origin.z = 0
                    Case 180
                        Player.Object.Origin.X = 0
                        Player.Object.Origin.z = -50000
                    Case 270
                        Player.Object.Origin.X = -50000
                        Player.Object.Origin.z = 0
                End Select
                If Clocker.FollowingMode Then
                    ResetActivity Partner.Object
                    Partner.Object.Origin.X = Player.Object.Origin.X
                    Partner.Object.Origin.Y = Player.Object.Origin.Y
                    Partner.Object.Origin.z = Player.Object.Origin.z
                End If
            End If
            
            If frmMain.Recording And Not frmMain.IsPlayback Then
                WarpData = WarpData & Player.Object.Origin.X & "," & Player.Object.Origin.Y & "," & Player.Object.Origin.z & ","
            ElseIf frmMain.Recording And frmMain.IsPlayback Then
                Player.Object.Origin.X = RemoveNextArg(WarpData, ",")
                Player.Object.Origin.Y = RemoveNextArg(WarpData, ",")
                Player.Object.Origin.z = RemoveNextArg(WarpData, ",")
                If Clocker.FollowingMode Then
                    Partner.Object.Origin.X = Player.Object.Origin.X
                    Partner.Object.Origin.Y = Player.Object.Origin.Y
                    Partner.Object.Origin.z = Player.Object.Origin.z
                End If
            End If

        ElseIf (Player.Object.Origin.Y >= 0) And ((TriTime = 1) Or (TriTime = 3)) Then
            ScoreFor = Replace(Replace(Replace(Replace(ScoreFor, "/", ""), "\", ""), "|", ""), "-", "")
            
            If (Not ((Player.Object.Origin.X >= -1943 And Player.Object.Origin.X <= -1864) And _
                (Player.Object.Origin.z >= 2352 And Player.Object.Origin.z <= 2560))) Then
                
                If (Distance(Player.Object.Origin.X, Player.Object.Origin.Y, Player.Object.Origin.z, 0, 0, 0) < 2000) Then
                    
                    If Not (TriTime = 2) Then
                        TriTime = 2
                        ForTime = Trim(DateDiff("s", ForTime, GetNow) & "s")
                    End If
                    FadeMessage "Positive pitch black issue on the black lawn, well done."
                
                    vif = vif + 1
                    If (Not frmMain.Recording) And (Not frmMain.Multiplayer) Then db.dbQuery "UPDATE Settings SET ScoreFor='" & ScoreFor & "' WHERE Username='" & Replace(GetUserLoginName, "'", "''") & "';"
                End If

            End If
            
        End If

    ElseIf (TriTime = 3) Then

        If (Distance(Player.Object.Origin.X, Player.Object.Origin.Y, Player.Object.Origin.z, 0, 0, 0) < 2000) Then

            TriTime = 5
            ScoreFor = Trim(ScoreFor) & " " & Trim(DateDiff("s", ForTime, GetNow)) & "s"
            ScoreFor = Replace(Replace(Replace(Replace(ScoreFor, "/", ""), "\", ""), "|", ""), "-", "")

            ForTime = ""
            FadeMessage "Positive pitch black issue on the black lawn, well done."
            vif = vif + 1
            If (Not frmMain.Recording) And (Not frmMain.Multiplayer) Then db.dbQuery "UPDATE Settings SET ScoreFor='" & ScoreFor & "' WHERE Username='" & Replace(GetUserLoginName, "'", "''") & "';"

        End If

    End If
    
    Static WarpStack1 As Boolean
    Static WarpStack2 As Boolean
    
    If ((Player.Object.Origin.z >= 320 And Player.Object.Origin.z <= 780) And (Player.Object.Origin.X >= -2850 And Player.Object.Origin.X <= -2329)) And (Player.Object.Origin.Y <= 0) Then
        If (Player.Object.Origin.Y = 0) Then
            Player.Stalled = True
        End If
        WarpStack2 = True
    ElseIf (Player.Object.Origin.Y > 0) Then
        WarpStack2 = False
    End If
    
    If ((Player.Object.Origin.z >= -910 And Player.Object.Origin.z <= -181) And (Player.Object.Origin.X >= -3099 And Player.Object.Origin.X <= -2318)) And (Player.Object.Origin.Y <= 0) Then
        If (Player.Object.Origin.Y = 0) Then
            Player.Gravity = -(GravityVelocity * 2)
            PlayWave SOUND_LAUNCH
        End If
        WarpStack1 = True
    ElseIf (Player.Object.Origin.Y = 0) And (Not WarpStack2) Then
        WarpStack1 = False
    End If
    
    If (WarpStack1 And WarpStack2) Then
        HandiCap = True
        WarpStack1 = False
        WarpStack2 = False
    End If
    
    If (Player.Object.Origin.Y > 0) Then Player.Stalled = False
    
    If (TriTime = 3) Then
        fiv = fiv + 2
    ElseIf (TriTime >= 1) And (TriTime < 5) Then
        fiv = fiv + 1
    End If
    
    If IsDate(XthTime) Then
        
        If DateDiff("s", XthTime, GetNow) > 120 Then
            ScoreXth = "Xth " & (DateDiff("s", XthTime, GetNow) - 120) & "v"
            fiv = fiv + 1
        End If
        
    ElseIf XthTime = "" Then
        
        If Not (ScoreXth = "") Then
            If (Not frmMain.Recording) And (Not frmMain.Multiplayer) Then db.dbQuery "UPDATE Settings SET ScoreXth='" & ScoreXth & "' WHERE Username='" & Replace(GetUserLoginName, "'", "''") & "';"
        End If
        
    End If

    If vif >= 2 Then fiv = fiv + vif
    
    If fiv > 1 Then
        If (Not (Right(FithTime, Len(CStr(fiv))) = CStr(fiv))) And (Not ((fiv = 2) And (FithTime = ""))) Then
            FithTime = Trim(CStr(FithTime)) & Trim(CStr(fiv))
            FithTime = Right(FithTime, 3)
        End If
    End If

    If Not (FithTime = "") Then
        ScoreFith = "5th " & FithTime & "!"
        If (Not frmMain.Recording) And (Not frmMain.Multiplayer) Then db.dbQuery "UPDATE Settings SET ScoreFith='" & ScoreFith & "' WHERE Username='" & Replace(GetUserLoginName, "'", "''") & "';"
    End If
    
    If ((Player.Object.Origin.z >= -2570 And Player.Object.Origin.z <= -2500) And (Player.Object.Origin.X >= 2200 And Player.Object.Origin.X <= 2600)) And (Player.Object.Origin.Y = 0) Then
        Player.Gravity = -(GravityVelocity * 2)
        PlayWave SOUND_LAUNCH
        Player.Trails = Not Player.Trails
    End If
    
    Static inGas As Boolean
    If ((Player.Object.Origin.X >= -297 And Player.Object.Origin.X <= -31) And _
        (Player.Object.Origin.z >= -2233 And Player.Object.Origin.z <= -1717)) And _
        ((Player.Object.Origin.Y >= 0) And (Player.Object.Origin.Y < 227)) Then
            If (Not inGas) Then
                inGas = True
                CenterText = "Collect enough nickels and you can change your body kits by driving through the model."
            End If
    ElseIf inGas Then
        CenterText = ""
        inGas = False
    End If
    
    If ObjectCount > 0 Then
        For cnt = 1 To ObjectCount
            If Objects(cnt).IsIdol Then
                If Distance(Objects(cnt).Origin.X, Objects(cnt).Origin.Y, Objects(cnt).Origin.z, Player.Object.Origin.X, Player.Object.Origin.Y, Player.Object.Origin.z) <= WithInCityLimits Then
                    If Objects(cnt).MeshIndex > 0 Then
                        If Not (InStr(ScoreIdol, Left(LCase(Meshes(Objects(cnt).MeshIndex).FileName), 1)) > 0) Then
                            ScoreIdol = ScoreIdol & Left(LCase(Meshes(Objects(cnt).MeshIndex).FileName), 1)
                        End If
                    End If
                End If
            End If
        Next
    End If
    
    db.rsClose rs
End Sub

Private Function AdjustDate(ByVal InDate As String, ByVal DownDate As String) As String
    If IsDate(InDate) Then
        AdjustDate = DateAdd("s", DateDiff("s", DownDate, GetNow), InDate)
    Else
        AdjustDate = InDate
    End If
End Function


Public Function DatabaseFilePath(Optional ByVal UserLoc As Boolean = False) As String
    Dim retVal As String
  '  If retVal = "" Then
        retVal = AppPath & Replace(App.EXEName, ".exe", "") & ".mdb"
        If PathExists(Replace(retVal, GetProgramFilesFolder, GetAllUsersAppDataFolder, , , vbTextCompare), True) And (Not PathExists(retVal, True)) Then
            retVal = Replace(retVal, GetProgramFilesFolder, GetAllUsersAppDataFolder, , , vbTextCompare)
        End If
   ' End If
    If UserLoc Then
        retVal = Replace(Replace(retVal, GetProgramFilesFolder, GetCurrentAppDataFolder, , , vbTextCompare), GetAllUsersAppDataFolder, GetCurrentAppDataFolder, , , vbTextCompare) & ".sql"
        MakeFolder GetFilePath(retVal)
    End If
    DatabaseFilePath = retVal
End Function

Public Sub BackupDB()
    Dim txt As String
    Dim tmp As String
    Set db = New clsDatabase
    Dim rs As ADODB.Recordset
    Set rs = CreateObject("ADODB.Recordset")
    db.rsQuery rs, "SELECT Username FROM Settings;"
    If Not db.rsEnd(rs) Then
        rs.MoveFirst
        Do
            tmp = GetUserData(rs("Username"))
            txt = txt & tmp & vbCrLf
            ClearUserData tmp
            rs.MoveNext
        Loop Until db.rsEnd(rs)
    End If
    db.rsClose rs
    WriteFile DatabaseFilePath(True), txt
    Set db = Nothing
    
End Sub

Public Function RestoreDB(Optional ByVal RemoveINI As Boolean = True)
    Set db = New clsDatabase
    db.dbQuery "DELETE * FROM Settings;"
    Dim txt As String
    Dim Data As String
    Dim Username As String
    Data = ReadFile(DatabaseFilePath(True))
    Do While Not (Data = "")
        txt = RemoveNextArg(Data, vbCrLf)
        If Not (txt = "") Then
           
            Username = Replace(RemoveArg(txt, "WHERE Username='"), "';", "")
            db.dbQuery "INSERT INTO Settings (Username) VALUES ('" & Replace(Username, "'", "''") & "');"
            db.dbQuery txt
        End If
        
    Loop
    If RemoveINI Then Kill DatabaseFilePath(True)
    Set db = Nothing
End Function

Public Sub ResetDB()
    Set db = New clsDatabase
    db.dbQuery "DELETE * FROM Settings;"
    Set db = Nothing
End Sub

Public Sub ClearUserData(Optional ByVal uData As String = "")
    Dim User As String
    
    ResetAllActivities
    
    If Not (uData = "") Then
        
        User = RemoveArg(uData, "UPDATE Settings SET Username='")
        User = NextArg(User, "', ")
        
        db.dbQuery "DELETE * FROM Settings WHERE Username='" & Replace(User, "'", "''") & "';"
        
        db.dbQuery uData
    
    Else
        User = GetUserLoginName
    End If
    
    Score1 = 0
    Score2 = 0
    Score3 = 0
    Player.Object.Origin = MakeVector(0, 0, 0)
    Player.Rotation = 0
    Player.CameraAngle = 0
    Player.CameraPitch = 0
    Player.CameraZoom = 500
    Player.MoveSpeed = 1
    Player.AutoMove = False
    Player.Gravity = 0
    Player.Trails = False
    Player.Model = 0
    Player.Stalled = False
    Partner.Object.Origin.X = 0
    Partner.Object.Origin.Y = 0
    Partner.Object.Origin.z = 0
    Partner.Rotation = 0
    Clocker.FollowingMode = False
    Bodykits = "0"
    HandiCap = False
    'PortalHold = 0
    ScoreXth = ""
    Score1st = "1st 0s"
    Score2nd = "2nd 0s 0s"
    ScoreClr = 0
    ScoreFor = ""
    ScoreFith = ""
    ScoreIdol = ""
    ScoreNth = 0
    GainMode = 0
    WonTime = 0
    TooTime = 0
    ClrTime = 0
    TriTime = 0
    ForTime = ""
    XthTime = ""
    MuseTime = ""
    SakeTime = ""
    FithTime = ""
    SixTime = 0
    
End Sub

Public Sub LoadUserData(Optional ByVal uData As String = "")
    Dim User As String
    
    If Not (uData = "") Then
        
        User = RemoveArg(uData, "UPDATE Settings SET Username='")
        User = NextArg(User, "', ")
        
        db.dbQuery "INSERT INTO Settings (Username) VALUES ('" & Replace(User, "'", "''") & "');"
        
        db.dbQuery uData
    
    Else
        User = GetUserLoginName
    End If
        
    db.rsQuery rs, "SELECT * FROM Settings WHERE Username = '" & Replace(User, "'", "''") & "';"
    
    If Not db.rsEnd(rs) Then
        Score1 = rs("Score1")
        Score2 = rs("Score2")
        Score3 = rs("Score3")
        Player.Object.Origin.X = rs("LocX")
        Player.Object.Origin.Y = rs("LocY")
        Player.Object.Origin.z = rs("LocZ")
        Player.Rotation = rs("Rot")
        Player.CameraAngle = rs("Angle")
        Player.CameraPitch = rs("Pitch")
        Player.CameraZoom = rs("CameraZoom")
        Player.MoveSpeed = rs("Speed")
        Player.AutoMove = rs("AutoMove")
        Player.Gravity = rs("Gravity")
        Player.Trails = rs("Trails")
        Player.Model = rs("Model")
        Player.Stalled = rs("Stalled")
        Partner.Object.Origin.X = rs("PartnerX")
        Partner.Object.Origin.Y = rs("PartnerY")
        Partner.Object.Origin.z = rs("PartnerZ")
        Partner.Rotation = rs("PartnerRot")
        Clocker.FollowingMode = rs("PartnerFollow")
        Bodykits = rs("Bodykits")
        HandiCap = rs("HandiCap")
        'PortalHold = rs("PortalHold")
        ScoreXth = rs("ScoreXth")
        Score1st = rs("Score1st")
        Score2nd = rs("Score2nd")
        ScoreClr = rs("ScoreClr")
        ScoreFor = rs("ScoreFor")
        ScoreFith = rs("ScoreFith")
        ScoreIdol = rs("ScoreIdol")
        ScoreNth = rs("ScoreNth")
        GainMode = rs("GainMode")
        WonTime = rs("WonTime")
        TooTime = rs("TooTime")
        ClrTime = rs("ClrTime")
        TriTime = rs("TriTime")
        DownTime = CStr(rs("DownTime"))
        ForTime = AdjustDate(CStr(rs("ForTime")), CStr(rs("DownTime")))
        MuseTime = AdjustDate(CStr(rs("MuseTime")), CStr(rs("DownTime")))
        SakeTime = AdjustDate(CStr(rs("SakeTime")), CStr(rs("DownTime")))
        XthTime = AdjustDate(CStr(rs("XthTime")), CStr(rs("DownTime")))
        FithTime = rs("FithTime")
        SixTime = rs("SixTime")
    Else
        ClearUserData
    End If
    
    db.rsClose rs
End Sub
Public Function GetUserData(Optional ByVal User As String = "") As String
    Dim txt As String
    If User = "" Then
        User = GetUserLoginName
        txt = "UPDATE Settings SET Username='" & Replace(GetUserLoginName, "'", "''") & "', "
        txt = txt & "Score1=" & Score1 & ", "
        txt = txt & "Score2=" & Score2 & ", "
        txt = txt & "Score3=" & Score3 & ", "
        txt = txt & "LocX=" & Player.Object.Origin.X & ", "
        txt = txt & "LocY=" & Player.Object.Origin.Y & ", "
        txt = txt & "LocZ=" & Player.Object.Origin.z & ", "
        txt = txt & "Rot=" & Player.Rotation & ", "
        txt = txt & "Angle=" & Player.CameraAngle & ", "
        txt = txt & "Pitch=" & Player.CameraPitch & ", "
        txt = txt & "CameraZoom=" & Player.CameraZoom & ", "
        txt = txt & "Speed=" & Player.MoveSpeed & ", "
        txt = txt & "AutoMove=" & Player.AutoMove & ", "
        txt = txt & "Gravity=" & Player.Gravity & ", "
        txt = txt & "Trails=" & Player.Trails & ", "
        txt = txt & "Model=" & Player.Model & ", "
        txt = txt & "Stalled=" & Player.Stalled & ", "
        txt = txt & "PartnerX=" & Partner.Object.Origin.X & ", "
        txt = txt & "PartnerY=" & Partner.Object.Origin.Y & ", "
        txt = txt & "PartnerZ=" & Partner.Object.Origin.z & ", "
        txt = txt & "PartnerRot=" & Partner.Rotation & ", "
        txt = txt & "PartnerFollow=" & Clocker.FollowingMode & ", "
        txt = txt & "Bodykits='" & Bodykits & "', "
        txt = txt & "HandiCap=" & HandiCap & ", "
        'txt = txt & "PortalHold=" & PortalHold & ", "
        txt = txt & "ScoreXth='" & ScoreXth & "', "
        txt = txt & "Score1st='" & Score1st & "', "
        txt = txt & "Score2nd='" & Score2nd & "', "
        txt = txt & "ScoreClr=" & ScoreClr & ", "
        txt = txt & "ScoreFor='" & ScoreFor & "', "
        txt = txt & "ScoreFith='" & ScoreFith & "', "
        txt = txt & "ScoreIdol='" & ScoreIdol & "', "
        txt = txt & "ScoreNth=" & ScoreNth & ", "
        txt = txt & "GainMode=" & GainMode & ", "
        txt = txt & "WonTime=" & WonTime & ", "
        txt = txt & "TooTime=" & TooTime & ", "
        txt = txt & "TriTime=" & TriTime & ", "
        txt = txt & "ForTime='" & ForTime & "', "
        txt = txt & "ClrTime=" & ClrTime & ", "
        txt = txt & "MuseTime='" & MuseTime & "', "
        txt = txt & "SakeTime='" & SakeTime & "', "
        txt = txt & "XthTime='" & XthTime & "', "
        txt = txt & "FithTime='" & FithTime & "', "
        txt = txt & "SixTime=" & SixTime & ", "
        txt = txt & "DownTime='" & GetNow & "' "
        txt = txt & "WHERE Username='" & Replace(GetUserLoginName, "'", "''") & "';"
        
    Else
        db.rsQuery rs, "SELECT * FROM Settings WHERE Username='" & User & "';"
        txt = "UPDATE Settings SET Username='" & Replace(User, "'", "''") & "', "
        txt = txt & "Score1=" & rs("Score1") & ", "
        txt = txt & "Score2=" & rs("Score2") & ", "
        txt = txt & "Score3=" & rs("Score3") & ", "
        txt = txt & "LocX=" & rs("LocX") & ", "
        txt = txt & "LocY=" & rs("LocY") & ", "
        txt = txt & "LocZ=" & rs("LocZ") & ", "
        txt = txt & "Rot=" & rs("Rot") & ", "
        txt = txt & "Angle=" & rs("Angle") & ", "
        txt = txt & "Pitch=" & rs("Pitch") & ", "
        txt = txt & "CameraZoom=" & rs("CameraZoom") & ", "
        txt = txt & "Speed=" & rs("Speed") & ", "
        txt = txt & "AutoMove=" & rs("AutoMove") & ", "
        txt = txt & "Gravity=" & rs("Gravity") & ", "
        txt = txt & "Trails=" & rs("Trails") & ", "
        txt = txt & "Model=" & rs("Model") & ", "
        txt = txt & "Stalled=" & rs("Stalled") & ", "
        txt = txt & "PartnerX=" & rs("PartnerX") & ", "
        txt = txt & "PartnerY=" & rs("PartnerY") & ", "
        txt = txt & "PartnerZ=" & rs("PartnerZ") & ", "
        txt = txt & "PartnerRot=" & rs("PartnerRot") & ", "
        txt = txt & "PartnerFollow=" & rs("PartnerFollow") & ", "
        txt = txt & "Bodykits='" & rs("Bodykits") & "', "
        txt = txt & "HandiCap=" & rs("HandiCap") & ", "
        'txt = txt & "PortalHold=" & rs("PortalHold") & ", "
        txt = txt & "ScoreXth='" & rs("ScoreXth") & "', "
        txt = txt & "Score1st='" & rs("Score1st") & "', "
        txt = txt & "Score2nd='" & rs("Score2nd") & "', "
        txt = txt & "ScoreClr=" & rs("ScoreClr") & ", "
        txt = txt & "ScoreFor='" & rs("ScoreFor") & "', "
        txt = txt & "ScoreFith='" & rs("ScoreFith") & "', "
        txt = txt & "ScoreIdol='" & rs("ScoreIdol") & "', "
        txt = txt & "ScoreNth=" & rs("ScoreNth") & ", "
        txt = txt & "GainMode=" & rs("GainMode") & ", "
        txt = txt & "WonTime=" & rs("WonTime") & ", "
        txt = txt & "TooTime=" & rs("TooTime") & ", "
        txt = txt & "TriTime=" & rs("TriTime") & ", "
        txt = txt & "ForTime='" & rs("ForTime") & "', "
        txt = txt & "ClrTime=" & rs("ClrTime") & ", "
        txt = txt & "MuseTime='" & rs("MuseTime") & "', "
        txt = txt & "SakeTime='" & rs("SakeTime") & "', "
        txt = txt & "XthTime='" & rs("XthTime") & "', "
        txt = txt & "FithTime='" & rs("FithTime") & "', "
        txt = txt & "SixTime=" & rs("SixTime") & ", "
        txt = txt & "DownTime='" & rs("DownTime") & "' "
        txt = txt & "WHERE Username='" & Replace(User, "'", "''") & "';"
        
    End If

    GetUserData = txt
End Function
Public Sub SaveUserData()
    db.dbQuery GetUserData
End Sub

Public Sub CompactDB()
On Error GoTo catch
    
    Dim JRO As New JRO.JetEngine
    Dim sPath As String
    Dim dPath As String

    sPath = DatabaseFilePath()
    dPath = DatabaseFilePath(True) & ".bak"
    
    If PathExists(dPath) Then Kill dPath
    JRO.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sPath & ";Jet OLEDB:Database Password=" & Replace(App.EXEName, ".exe", "", , , vbTextCompare) & ";", _
                        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dPath & ";Jet OLEDB:Database Password=" & Replace(App.EXEName, ".exe", "", , , vbTextCompare) & ";Jet OLEDB:Engine Type=5"

    Kill sPath
    FileCopy dPath, sPath
    Kill dPath

    Set JRO = Nothing
    
catch:
    If Err Then Err.Clear
    On Error GoTo 0
End Sub