Attribute VB_Name = "modCmds"
#Const modCmds = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module
Private DInput As DirectInput8
Private DIKeyBoardDevice As DirectInputDevice8
Private DIKEYBOARDSTATE As DIKEYBOARDSTATE

Private DIMouseDevice As DirectInputDevice8
Private DIMOUSESTATE As DIMOUSESTATE

Private ToggleSound1 As Boolean
Private ToggleSound2 As Boolean
Private ToggleSound3 As Boolean
Private TogglePress1 As Long
Private TogglePress2 As Long
Private TogglePress3 As Long

Private Const MaxConsoleMsgs = 20
Private Const MaxHistoryMsgs = 10

Private ConsoleMsgs As Collection
Private HistoryMsgs As Collection
Private HistoryPoint As Integer
Private CommandLine As String

Private Type KeyState
    VKState As Integer
    VKToggle As Boolean
    VKPressed As Boolean
    VKLatency As Single
End Type

Private CapsLock As Boolean
Private KeyState(255) As KeyState
Private KeyChars(255) As String

Private CursorPos As Long
Private CursorX As Long
Private CursorWidth As Long
Private ConsoleWidth As Single
Private ConsoleHeight As Single

Private Backdrop As Direct3DTexture8
Private Vertex(0 To 4) As TVERTEX1

Private State As Integer
Private Shift As Integer
Private Bottom As Long

Private Function Toggled(ByVal vkCode As Long) As Boolean
    Toggled = KeyState(vkCode).VKToggle
End Function
Private Function Pressed(ByVal vkCode As Long) As Boolean
    Pressed = KeyState(vkCode).VKPressed
End Function

Public Property Get TextHeight() As Single
    TextHeight = frmMain.TextHeight("A")
End Property

Public Property Get TextSpace() As Single
    TextSpace = 2
End Property

Public Sub InputScene()
        
    On Error GoTo pausing
    
    Dim Mov As D3DVECTOR
    Dim Rot As Single
        
    DIKeyBoardDevice.GetDeviceStateKeyboard DIKEYBOARDSTATE
    
    ConsoleInput DIKEYBOARDSTATE
    
    If DIKEYBOARDSTATE.Key(DIK_ESCAPE) Then
        If (Not TogglePress1 = DIK_ESCAPE) Then
            TogglePress1 = DIK_ESCAPE
            If ((Not FullScreen) And TrapMouse) Or ConsoleVisible Then
                TrapMouse = False
            ElseIf FullScreen Or (Not FullScreen And Not TrapMouse) Then
                StopGame = True
            End If
        End If
    ElseIf DIKEYBOARDSTATE.Key(DIK_GRAVE) Then
        If (Not TogglePress1 = DIK_GRAVE) Then
            TogglePress1 = DIK_GRAVE
            ConsoleToggle
        End If
    ElseIf DIKEYBOARDSTATE.Key(DIK_LALT) Then
        If (Not TogglePress1 = DIK_LALT) Then
            TogglePress1 = DIK_LALT
            If DIKEYBOARDSTATE.Key(DIK_TAB) Then
                frmMain.WindowState = 1
            End If
        End If
    ElseIf Not (TogglePress1 = 0) Then
        TogglePress1 = 0
    End If
       
    If (Not ConsoleVisible) And TrapMouse And (MenuMode = 0) Then

        If DIKEYBOARDSTATE.Key(DIK_Q) Then
            Player.MoveSpeed = Player.MoveSpeed + 0.1
            If Player.MoveSpeed > MaxDisplacement Then Player.MoveSpeed = MaxDisplacement
        End If
        
        If DIKEYBOARDSTATE.Key(DIK_A) Then
            Player.MoveSpeed = Player.MoveSpeed - 1
            If Player.MoveSpeed < 1 Then Player.MoveSpeed = 0.1
        End If
            
        If (DIKEYBOARDSTATE.Key(DIK_E)) Then
            Mov.X = Mov.X + (Sin(D720 - Player.CameraAngle) * Player.MoveSpeed)
            Mov.Z = Mov.Z + (Cos(D720 - Player.CameraAngle) * Player.MoveSpeed)
            AttemptMoves Mov
        End If
        If (DIKEYBOARDSTATE.Key(DIK_D)) Then
            Mov.X = Mov.X - (Sin(D720 - Player.CameraAngle) * Player.MoveSpeed)
            Mov.Z = Mov.Z - (Cos(D720 - Player.CameraAngle) * Player.MoveSpeed)
            AttemptMoves Mov
        End If
        If (DIKEYBOARDSTATE.Key(DIK_W)) Then
            Mov.X = Mov.X + (Sin((D720 - Player.CameraAngle) - D180) * Player.MoveSpeed)
            Mov.Z = Mov.Z + (Cos((D720 - Player.CameraAngle) - D180) * Player.MoveSpeed)
            AttemptMoves Mov
        End If
        If (DIKEYBOARDSTATE.Key(DIK_R)) Then
            Mov.X = Mov.X + (Sin((D720 - Player.CameraAngle) + D180) * Player.MoveSpeed)
            Mov.Z = Mov.Z + (Cos((D720 - Player.CameraAngle) + D180) * Player.MoveSpeed)
            AttemptMoves Mov
        End If
        
        If DIKEYBOARDSTATE.Key(DIK_F3) Then
            If (Not TogglePress3 = DIK_F3) Then
                TogglePress3 = DIK_F3
                
                Dim inStart As Boolean
                Dim o As Long
                
                For o = 1 To UBound(Level.Starts)
                    If (Player.Location.X > LeastX(Level.Starts(o).Point1, Level.Starts(o).Point2, Level.Starts(o).Point3, Level.Starts(o).Point4) And _
                        Player.Location.X < GreatestX(Level.Starts(o).Point1, Level.Starts(o).Point2, Level.Starts(o).Point3, Level.Starts(o).Point4) And _
                        Player.Location.Z > LeastZ(Level.Starts(o).Point1, Level.Starts(o).Point2, Level.Starts(o).Point3, Level.Starts(o).Point4) And _
                        Player.Location.Z < GreatestZ(Level.Starts(o).Point1, Level.Starts(o).Point2, Level.Starts(o).Point3, Level.Starts(o).Point4)) Then
                        inStart = True
                        
                    End If
                Next
                If inStart And (Level.Elapsed = 0) Then
                    CenterMessage ""
                    db.rsQuery rs, "SELECT * FROM Scores WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "' AND LevelNum=" & Replace(Replace(Replace(LCase(Level.Loaded), "level", ""), ".hog", ""), "'", "''") & ";"
                    If Not db.rsEnd(rs) Then
'                        AtLvl = AtLvl + 1
'                        If Not PathExists(AppPath & "Levels\level" & AtLvl & ".hog", True) Then AtLvl = 1
'                        FadeMessage "Level " & AtLvl
'                        LoadLevel AppPath & "Levels\level" & AtLvl & ".hog"
                        MenuMode = 1
                    Else
                        FadeMessage "You must pass this level to skip it."
                    End If
                End If
            End If
        ElseIf Not (TogglePress3 = 0) Then
            TogglePress3 = 0
        End If
        
    ElseIf (MenuMode <> 0) Then
        If ((MenuMode = -1) Or (MenuMode = 2)) And DIKEYBOARDSTATE.Key(DIK_F1) Then
            AtLvl = AtLvl - 1
            MenuMode = 1
        ElseIf (MenuMode = 2) And DIKEYBOARDSTATE.Key(DIK_F2) Then
            If Not PathExists(AppPath & "Levels\level" & AtLvl & ".hog", True) Then AtLvl = 0
            MenuMode = 1
        End If
    End If
        
    DIMouseDevice.GetDeviceStateMouse DIMOUSESTATE

    Dim mX As Integer
    Dim mY As Integer
    Dim mZ As Integer

    mX = DIMOUSESTATE.lX
    mY = DIMOUSESTATE.lY
    mZ = DIMOUSESTATE.lZ
    
    If (GetActiveWindow = frmMain.hwnd) And TrapMouse Then
        If Not (frmMain.MousePointer = 99) Then
            frmMain.MousePointer = 99
            frmMain.MouseIcon = LoadPicture(AppPath & "Base\mouse.cur")
        End If
        
        If (MenuMode = 0) Then MouseLook mX, mY, mZ
        
        Dim rec As RECT
        GetWindowRect frmMain.hwnd, rec
        SetCursorPos rec.Right + ((rec.Left - rec.Right) / 2), rec.Top + ((rec.Bottom - rec.Top) / 2)
    Else
        If Not (frmMain.MousePointer = 1) Then frmMain.MousePointer = 1
    End If
    
    Exit Sub
pausing:
    Err.Clear
    DoPauseGame
End Sub

Private Sub MouseLook(ByVal X As Integer, ByVal Y As Integer, ByVal Z As Integer)

    Dim cnt As Long
    If Z < 0 Then
        For cnt = (Z * MouseSensitivity) To 0
            Player.CameraZoom = Player.CameraZoom + 0.15
        Next
    ElseIf Z > 0 Then
        For cnt = 0 To (Z * MouseSensitivity)
            Player.CameraZoom = Player.CameraZoom - 0.15
        Next
    End If
    
    If Player.CameraZoom > 2000 Then Player.CameraZoom = 2000
    If Player.CameraZoom < 500 Then Player.CameraZoom = 500

    If X < 0 Then
        For cnt = (X * MouseSensitivity) To 0
            Player.CameraAngle = Player.CameraAngle - -0.0015
        Next
    ElseIf X > 0 Then
        For cnt = 0 To (X * MouseSensitivity)
            Player.CameraAngle = Player.CameraAngle - 0.0015
        Next
    End If

    If Player.CameraAngle > (PI * 2) Then Player.CameraAngle = Player.CameraAngle - (PI * 2)
    If Player.CameraAngle < -(PI * 2) Then Player.CameraAngle = Player.CameraAngle + (PI * 2)

    If Y < 0 Then
        For cnt = (Y * MouseSensitivity) To 0
            Player.CameraPitch = Player.CameraPitch - -0.0015
        Next
    ElseIf Y > 0 Then
        For cnt = 0 To (Y * MouseSensitivity)
            Player.CameraPitch = Player.CameraPitch - 0.0015
        Next
    End If
    
    If Player.CameraPitch < -1.5 Then Player.CameraPitch = -1.5
    If Player.CameraPitch > 1.5 Then Player.CameraPitch = 1.5
    
End Sub

Public Property Get ConsoleVisible() As Boolean
    ConsoleVisible = Not (State = 0)
End Property

Public Sub ConsoleToggle()
    If State = 1 Then
        Shift = -Shift
    ElseIf State = 0 Then
        Shift = 18
        State = 1
    ElseIf State = 2 Then
        Shift = -18
        State = 1
    End If
End Sub

Public Sub RenderCmds()

    Select Case State
        Case 0 'up
            Bottom = 0
            Shift = 0
        Case 2 'down
            Bottom = ConsoleHeight
            Shift = 0
        Case 1 'moving
            If (Bottom + Shift) > ConsoleHeight Then
                State = 2
            ElseIf (Bottom + Shift) < 0 Then
                State = 0
            Else
                Bottom = Bottom + Shift
            End If
    End Select

    Vertex(0).Y = Bottom - ConsoleHeight
    Vertex(1).Y = Bottom - ConsoleHeight
    Vertex(2).Y = Bottom
    Vertex(3).Y = Bottom
    
    If ConsoleVisible Then
    
        DDevice.SetVertexShader FVF_VTEXT1
        DDevice.SetRenderState D3DRS_ZENABLE, False
        DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    
        DDevice.SetRenderState D3DRS_LIGHTING, False
        DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    
        DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
        DDevice.SetRenderState D3DRS_ALPHATESTENABLE, False
    
        DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
        DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
    
        
        DDevice.SetTexture 0, Backdrop
        DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0), LenB(Vertex(0))
    
        If Len(CommandLine) > 0 Then
            DrawText CommandLine, TextSpace + 2, Bottom - (TextHeight / Screen.TwipsPerPixelY) - TextSpace + 2
        End If
        
        Static DrawCursor As Double
        Static DrawBlink As Boolean
        If DrawCursor = 0 Or CDbl(((GetTimer * 1000) - DrawCursor)) >= 2000 Then
            DrawCursor = GetTimer
            DrawBlink = Not DrawBlink
            
            If DrawBlink Then DrawText String(CursorPos, " ") & "_", TextSpace + 2, Bottom - (TextHeight / Screen.TwipsPerPixelY) - TextSpace + 2

        End If
        
        If ConsoleMsgs.Count > 0 Then
            Dim ConsoleMsgY As Long
            ConsoleMsgY = (Bottom - (((TextHeight / Screen.TwipsPerPixelY) + TextSpace) * 2))
            Dim cnt As Integer
            For cnt = ConsoleMsgs.Count To 1 Step -1
                DrawText ConsoleMsgs(cnt), TextSpace + 2, ConsoleMsgY - ((ConsoleMsgs.Count - cnt) * (((TextHeight / Screen.TwipsPerPixelY) + TextSpace)))
            Next
        End If
        
        DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
        DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
        
        DDevice.SetRenderState D3DRS_ZENABLE, 1
        DDevice.SetRenderState D3DRS_LIGHTING, 1
    End If
End Sub

Public Function AddMessage(ByVal Message As String)
    If ConsoleMsgs.Count > 0 Then
        If (ConsoleMsgs.Item(ConsoleMsgs.Count) = Message) Then
            Exit Function
        End If
    End If
        
    If ConsoleMsgs.Count > MaxConsoleMsgs Then
        ConsoleMsgs.Remove 1
    End If
    ConsoleMsgs.Add Message
    
End Function

Public Sub ConsoleInput(ByRef kState As DIKEYBOARDSTATE)

    If ConsoleVisible And (Not (Shift < 0)) Then
        Dim cnt As Integer
        Dim char As String
        
        For cnt = 0 To 255

            If (Not (KeyState(cnt).VKState = kState.Key(cnt))) Then
                KeyState(cnt).VKState = kState.Key(cnt)
                If (KeyState(cnt).VKLatency = 0) Then
                    KeyState(cnt).VKPressed = Not KeyState(cnt).VKPressed
                    KeyState(cnt).VKLatency = Timer
                Else
                    KeyState(cnt).VKPressed = False
                End If
                If KeyState(cnt).VKPressed = False Then
                    KeyState(cnt).VKToggle = Not KeyState(cnt).VKToggle
                ElseIf KeyState(cnt).VKToggle Then
                    KeyState(cnt).VKToggle = KeyState(cnt).VKToggle Xor (Not KeyState(cnt).VKPressed)
                End If
                If Not KeyState(cnt).VKPressed Then KeyState(cnt).VKLatency = 0
            Else
                If (KeyState(cnt).VKLatency <> 0) Then
                    If ((Timer - -KeyState(cnt).VKLatency) > 0.1) And (KeyState(cnt).VKLatency < 0) Then
                        KeyState(cnt).VKPressed = True
                        KeyState(cnt).VKLatency = KeyState(cnt).VKLatency + -0.1
                    ElseIf ((Timer - KeyState(cnt).VKLatency) > 0.6) And (KeyState(cnt).VKLatency > 0) Then
                        KeyState(cnt).VKPressed = True
                        KeyState(cnt).VKLatency = KeyState(cnt).VKLatency - 0.6
                    Else
                        KeyState(cnt).VKPressed = False
                    End If
                Else
                    KeyState(cnt).VKPressed = False
                    KeyState(cnt).VKLatency = 0
                End If
            End If

            If Pressed(cnt) Then
            
                If cnt = DIK_CAPITAL Then
                    CapsLock = Not CapsLock
                ElseIf cnt = DIK_LSHIFT Or cnt = DIK_RSHIFT Then
                
                ElseIf cnt = DIK_RETURN Then
                    HistoryMsgs.Add CommandLine
                    If HistoryMsgs.Count > MaxHistoryMsgs Then
                        HistoryMsgs.Remove 1
                    End If
                    HistoryPoint = HistoryMsgs.Count + 1
                    
                    Process CommandLine
                    
                    CommandLine = ""
                    CursorPos = 0
                ElseIf cnt = DIK_BACK Then
                    If CursorPos > 0 Then
                        CommandLine = Left(CommandLine, CursorPos - 1) & Mid(CommandLine, CursorPos + 1)
                        CursorPos = CursorPos - 1
                    End If
                
                ElseIf cnt = DIK_DELETE Then
                    CommandLine = Left(CommandLine, CursorPos) & Mid(CommandLine, CursorPos + 2)
                
                ElseIf cnt = DIK_LEFT Then
                    If CursorPos > 0 Then
                        CursorPos = CursorPos - 1
                    End If
                ElseIf cnt = DIK_HOME Then
                    CursorPos = 0
                ElseIf cnt = DIK_END Then
                    CursorPos = Len(CommandLine)
                ElseIf cnt = DIK_RIGHT Then
                    If CursorPos < Len(CommandLine) Then
                        CursorPos = CursorPos + 1
                    End If
                ElseIf cnt = DIK_UP Then
                    If HistoryMsgs.Count > 0 Then
                        If HistoryPoint > 1 Then
                            HistoryPoint = HistoryPoint - 1
                            CommandLine = HistoryMsgs(HistoryPoint)
                            CursorPos = Len(CommandLine)
                        End If
                    End If
                ElseIf cnt = DIK_DOWN Then
                    If HistoryPoint = HistoryMsgs.Count Then
                        HistoryPoint = HistoryPoint + 1
                        CommandLine = ""
                        CursorPos = 0
                    ElseIf HistoryPoint <= HistoryMsgs.Count Then
                    
                        HistoryPoint = HistoryPoint + 1
                        CommandLine = HistoryMsgs(HistoryPoint)
                        CursorPos = Len(CommandLine)
                    End If
                Else
                
                    char = KeyChars(cnt)
                    If Not char = "" Then
                        If CapsLock Or kState.Key(DIK_LSHIFT) Or kState.Key(DIK_RSHIFT) Then
                            Select Case char
                                Case "1"
                                    char = "!"
                                Case "2"
                                    char = "@"
                                Case "3"
                                    char = "#"
                                Case "4"
                                    char = "$"
                                Case "5"
                                    char = "%"
                                Case "6"
                                    char = "^"
                                Case "7"
                                    char = "&"
                                Case "8"
                                    char = "*"
                                Case "9"
                                    char = "("
                                Case "0"
                                    char = ")"
                                Case "-"
                                    char = "_"
                                Case "="
                                    char = "+"
                                Case "["
                                    char = "{"
                                Case "]"
                                    char = "}"
                                Case "\"
                                    char = "|"
                                Case ";"
                                    char = ":"
                                Case "'"
                                    char = """"
                                Case ","
                                    char = "<"
                                Case "."
                                    char = ">"
                                Case "/"
                                    char = "?"
                                Case Else
                                    char = UCase(char)
                            End Select
                        End If
                        
                        If Len(CommandLine) <= 40 Then
                        
                            CommandLine = CommandLine & char
                            CursorPos = CursorPos + 1
                            
                        End If
                    
                    End If
                End If
                
                CursorX = (frmMain.TextWidth(Left(CommandLine, CursorPos)) / Screen.TwipsPerPixelX)
                CursorWidth = (frmMain.TextWidth(Mid(CommandLine, CursorPos + 1, 1)) / Screen.TwipsPerPixelX)
                If CursorWidth = 0 Then CursorWidth = 8
                
            End If
        Next
    End If
End Sub

Private Function InitKeys()

    KeyChars(2) = "1"
    KeyChars(3) = "2"
    KeyChars(4) = "3"
    KeyChars(5) = "4"
    KeyChars(6) = "5"
    KeyChars(7) = "6"
    KeyChars(8) = "7"
    KeyChars(9) = "8"
    KeyChars(10) = "9"
    KeyChars(11) = "0"
    KeyChars(12) = "-"
    KeyChars(13) = "="
    KeyChars(16) = "q"
    KeyChars(17) = "w"
    KeyChars(18) = "e"
    KeyChars(19) = "r"
    KeyChars(20) = "t"
    KeyChars(21) = "y"
    KeyChars(22) = "u"
    KeyChars(23) = "i"
    KeyChars(24) = "o"
    KeyChars(25) = "p"
    KeyChars(26) = "["
    KeyChars(27) = "]"
    
    KeyChars(30) = "a"
    KeyChars(31) = "s"
    KeyChars(32) = "d"
    KeyChars(33) = "f"
    KeyChars(34) = "g"
    KeyChars(35) = "h"
    KeyChars(36) = "j"
    KeyChars(37) = "k"
    KeyChars(38) = "l"
    KeyChars(39) = ";"
    KeyChars(40) = "'"
    KeyChars(43) = "\"
    KeyChars(44) = "z"
    KeyChars(45) = "x"
    KeyChars(46) = "c"
    KeyChars(47) = "v"
    KeyChars(48) = "b"
    KeyChars(49) = "n"
    KeyChars(50) = "m"
    KeyChars(51) = ","
    KeyChars(52) = "."
    KeyChars(53) = "/"
    KeyChars(55) = "*"
    KeyChars(57) = " "

    KeyChars(71) = "7"
    KeyChars(72) = "8"
    KeyChars(73) = "9"
    KeyChars(75) = "4"
    KeyChars(76) = "5"
    KeyChars(77) = "6"
    KeyChars(79) = "1"
    KeyChars(80) = "2"
    KeyChars(81) = "3"
    KeyChars(82) = "0"
    KeyChars(74) = "-"
    KeyChars(78) = "+"
    KeyChars(83) = "."
    KeyChars(181) = "/"

End Function

Public Sub CreateCmds()
    
    Set ConsoleMsgs = New Collection
    Set HistoryMsgs = New Collection
    
    Bottom = 0
    
    Vertex(0).X = 0
    Vertex(0).Z = -1
    Vertex(0).RHW = 1
    Vertex(0).color = D3DColorARGB(255, 255, 255, 255)
    Vertex(0).tu = 0
    Vertex(0).tv = 0
    
    Vertex(1).X = (frmMain.width / Screen.TwipsPerPixelX)
    Vertex(1).Z = -1
    Vertex(1).RHW = 1
    Vertex(1).color = D3DColorARGB(255, 255, 255, 255)
    Vertex(1).tu = 1
    Vertex(1).tv = 0
    
    Vertex(2).X = 0
    Vertex(2).Z = -1
    Vertex(2).RHW = 1
    Vertex(2).color = D3DColorARGB(255, 255, 255, 255)
    Vertex(2).tu = 0
    Vertex(2).tv = 1
    
    Vertex(3).X = (frmMain.width / Screen.TwipsPerPixelX)
    Vertex(3).Z = -1
    Vertex(3).RHW = 1
    Vertex(3).color = D3DColorARGB(255, 255, 255, 255)
    Vertex(3).tu = 1
    Vertex(3).tv = 1
    
    ConsoleWidth = (frmMain.width / Screen.TwipsPerPixelX)
    ConsoleHeight = (MaxConsoleMsgs * (frmMain.TextHeight("A") / Screen.TwipsPerPixelY)) + (TextSpace * MaxConsoleMsgs) + TextSpace
    If ConsoleHeight > ((frmMain.height / Screen.TwipsPerPixelY) \ 2) Then ConsoleHeight = ((frmMain.height / Screen.TwipsPerPixelY) \ 2)
    
    InitKeys
    
    Set Backdrop = LoadTexture(AppPath & "Base\drop.bmp")
    
    Set DInput = dx.DirectInputCreate()
        
    Set DIKeyBoardDevice = DInput.CreateDevice("GUID_SysKeyboard")
    DIKeyBoardDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
    If FullScreen Then
        DIKeyBoardDevice.SetCooperativeLevel frmMain.hwnd, DISCL_EXCLUSIVE Or DISCL_FOREGROUND
    Else
        DIKeyBoardDevice.SetCooperativeLevel frmMain.hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    End If
    DIKeyBoardDevice.Acquire
    
    Set DIMouseDevice = DInput.CreateDevice("GUID_SysMouse")
    DIMouseDevice.SetCommonDataFormat DIFORMAT_MOUSE
    If FullScreen Then
        DIMouseDevice.SetCooperativeLevel frmMain.hwnd, DISCL_EXCLUSIVE Or DISCL_FOREGROUND
    Else
        DIMouseDevice.SetCooperativeLevel frmMain.hwnd, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND
    End If
    DIMouseDevice.Acquire
    
End Sub

Public Sub CleanupCmds()
    
    If Not (ConsoleMsgs Is Nothing) Then
        Do While ConsoleMsgs.Count > 0
            ConsoleMsgs.Remove 1
        Loop
    End If
    
    If Not (HistoryMsgs Is Nothing) Then
        Do While HistoryMsgs.Count > 0
            HistoryMsgs.Remove 1
        Loop
    End If
    
    Set ConsoleMsgs = Nothing
    Set HistoryMsgs = Nothing
    
    Set Backdrop = Nothing
    
    DIKeyBoardDevice.Unacquire
    Set DIKeyBoardDevice = Nothing
    
    DIMouseDevice.Unacquire
    Set DIMouseDevice = Nothing
    
    Set DInput = Nothing
End Sub

Public Sub InitialCommands()

    Dim CommandsINI As String

    If PathExists(AppPath & "commands.ini") Then CommandsINI = ReadFile(AppPath & "commands.ini")
    
    Dim inLine As String
    Do Until CommandsINI = ""
        inLine = RemoveNextArg(CommandsINI, vbCrLf)
        Process inLine
    Loop
    
End Sub

Private Sub Process(ByVal inArg As String)
    Dim inCmd As String
    
    inCmd = RemoveNextArg(inArg, " ")
    If Left(inCmd, 1) = "/" Then inCmd = Mid(inCmd, 2)
    
    Select Case LCase(inCmd)

        Case "quit"
            StopGame = True
            
        Case "wire"
            Select Case LCase(Trim(inArg))
                Case "-1", "1", "true", "on"
                    WireFrame = True
                Case "0", "false", "off"
                    WireFrame = False
            End Select
            DDevice.SetRenderState D3DRS_FILLMODE, IIf(WireFrame, D3DFILL_WIREFRAME, D3DFILL_SOLID)
            db.dbQuery "UPDATE Settings SET WireFrame = " & IIf(WireFrame, "Yes", "No") & " WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"
            AddMessage "WireFrame = " & WireFrame
            
        Case "music"
            If Not DisableSound Then
                Select Case LCase(Trim(inArg))
                    Case "-1", "1", "true", "on"
                        PlayMusic = True
                        Track1.PlaySound
                        Track2.PlaySound
                    Case "0", "false", "off"
                        PlayMusic = False
                        Track1.StopSound
                        Track2.StopSound
                End Select
                db.dbQuery "UPDATE Settings SET MusicEnabled = " & IIf(PlayMusic, "Yes", "No") & " WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"
                AddMessage "MusicEnabled = " & PlayMusic
            Else
                AddMessage "Sound system disabled."
            End If
        Case "echo"
            AddMessage inArg
            
        Case "help", "cmdlist", "?", "--?"
            AddMessage ""
            AddMessage "Console Commands:"
            AddMessage "   QUIT (completely exit out of the game)"
            AddMessage "   MUSIC [<ON/OFF>] (get or set whether music is enabled)"
            AddMessage "   WIRE [<ON/OFF>] (get or set whether the lawn is rendered in wireframe)"
            AddMessage ""
        Case "load"
            LoadLevel AppPath & "Levels\" & inArg & ".hog"
            
        Case Else
            AddMessage "Unknown command."
    End Select
End Sub


