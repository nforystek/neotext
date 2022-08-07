#Const [True] = -1
#Const [False] = 0

Attribute VB_Name = "modInfo"
#Const modInfo = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module

Private FadeTime As Long
Private FadeText As String

Public HelpToggle As Boolean
Public CenterText As String

Private cTalkMsgs As Collection
Private tTalkTime As Double

Public Sub CreateInfo()
    Set cTalkMsgs = New Collection
End Sub

Public Sub CleanupInfo()
    Do While cTalkMsgs.Count > 0
        cTalkMsgs.Remove 1
    Loop
    Set cTalkMsgs = Nothing
End Sub

Public Function TalkMessage(ByVal Msg As String)
    If cTalkMsgs.Count >= MaxTalkMsgs Then
        cTalkMsgs.Remove 1
    End If
    tTalkTime = Timer
    cTalkMsgs.Add Msg
End Function

Public Sub FadeMessage(ByVal txt As String)
    FadeTime = Timer
    FadeText = txt
    AddMessage txt
End Sub

Private Function Row(ByVal num As Long) As Long
    Row = ((TextHeight \ Screen.TwipsPerPixelY) * num) + (2 * num)
End Function

Public Sub RenderInfo()
    DDevice.SetVertexShader FVF_SCREEN
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

    If Not ConsoleVisible Then
        
        Dim txt As String
        If HelpToggle Then
        
            DrawText "E/D/W/R=Move, S/F=Rotate, ARROWS=Camera, PAGES=Zoom, T=Auto, Q/A=Speed, SPACE=Climb, Y/H/G=Actions", 2, 2
            
    #If VBIDE = -1 Then
                DrawText "D: " & CSng(Distance(Player.Object.Origin.X, Player.Object.Origin.Y, Player.Object.Origin.z, 0, 0, 0)), 2, Row(2)
                DrawText "X: " & CSng(Player.Object.Origin.X), 2, Row(3)
                DrawText "Y: " & CSng(Player.Object.Origin.Y), 2, Row(4)
                DrawText "Z: " & CSng(Player.Object.Origin.z), 2, Row(5)
                DrawText "F: " & FPSRate, 2, Row(6)
    #Else
            If GodMode Then
                DrawText "D: " & CSng(Distance(Player.Object.Origin.X, Player.Object.Origin.Y, Player.Object.Origin.z, 0, 0, 0)), 2, Row(2)
                DrawText "X: " & CSng(Player.Object.Origin.X), 2, Row(3)
                DrawText "Y: " & CSng(Player.Object.Origin.Y), 2, Row(4)
                DrawText "Z: " & CSng(Player.Object.Origin.z), 2, Row(5)
                DrawText "F: " & FPSRate, 2, Row(6)
            End If
    #End If
        
            txt = "TILDA=Console, ESC=" & IIf(TrapMouse, "Exit", "Close")
            DrawText txt, (frmMain.ScaleWidth / Screen.TwipsPerPixelX) - (frmMain.TextWidth(txt) / Screen.TwipsPerPixelX), 2
                
            txt = "Begin score play by using number keys:"
            DrawText txt, (frmMain.ScaleWidth / Screen.TwipsPerPixelX) - (frmMain.TextWidth(txt) / Screen.TwipsPerPixelX), Row(2)
        
            txt = "1 = Warp random to a pitch black sight."
            DrawText txt, (frmMain.ScaleWidth / Screen.TwipsPerPixelX) - (frmMain.TextWidth(txt) / Screen.TwipsPerPixelX), Row(3)
        
            txt = "2 = Ride out to pitch black your self."
            DrawText txt, (frmMain.ScaleWidth / Screen.TwipsPerPixelX) - (frmMain.TextWidth(txt) / Screen.TwipsPerPixelX), Row(4)
        
            txt = "0 = Forfeit timed score, and warp back."
            DrawText txt, (frmMain.ScaleWidth / Screen.TwipsPerPixelX) - (frmMain.TextWidth(txt) / Screen.TwipsPerPixelX), Row(5)
            
            If Not frmMain.Multiplayer Then
                txt = "Lawn stars indicate points of interest."
                DrawText txt, (frmMain.ScaleWidth / Screen.TwipsPerPixelX) - (frmMain.TextWidth(txt) / Screen.TwipsPerPixelX), Row(6)
            End If
        Else
        
            DrawText "F1=Help", 2, 2

            txt = "ESC=" & IIf(TrapMouse, "Exit", "Close")
            DrawText txt, (frmMain.ScaleWidth / Screen.TwipsPerPixelX) - (frmMain.TextWidth(txt) / Screen.TwipsPerPixelX), 2

        End If
    End If
    
    If (frmMain.Recording Or frmMain.IsPlayback) Then
        txt = "Film: " & frmMain.RecordSize & " Bytes"
        DrawText txt, (frmMain.ScaleWidth / Screen.TwipsPerPixelX) - (frmMain.TextWidth(txt) / Screen.TwipsPerPixelX), Row(RowCount)
    End If
    
    If frmMain.Multiplayer Then
        DrawText Player.name, ((frmMain.ScaleWidth / Screen.TwipsPerPixelX) / 2) - ((frmMain.TextWidth(Player.name) / Screen.TwipsPerPixelX) / 2), Row(RowCount - 1)
    ElseIf frmMain.Recording Then
        DrawText IIf(frmMain.IsPlayback, "-=[Playback]=-", "+=[Recording]=+"), 10, Row(RowCount)
    End If
       
    If (cTalkMsgs.Count > 0) Then
        If ((Timer - tTalkTime) >= 8) Then
            tTalkTime = Timer
            cTalkMsgs.Remove 1
        End If

        Dim idx As Long
        For idx = 1 To cTalkMsgs.Count
            txt = cTalkMsgs.Item(idx)
            DrawText txt, 10, Row(1 + idx)
        Next
        
    End If
    
    If Not (FadeText = "") Then
        If (Timer - FadeTime) >= 6 Then
            FadeText = ""
        Else
            DrawText FadeText, ((frmMain.ScaleWidth / Screen.TwipsPerPixelX) / 2) - ((frmMain.TextWidth(FadeText) / Screen.TwipsPerPixelX) / 2), ((frmMain.ScaleHeight / Screen.TwipsPerPixelY) / 2) - (TextHeight / 2)
        End If
    End If
    
    If Not (CenterText = "") Then

        Dim cHeight As Single
        Dim cWidth As Single
        cWidth = GreatestWidth(CenterText)
        cHeight = TextHeight * CountWord(CenterText, vbCrLf)

        DrawText CenterText, ((frmMain.ScaleWidth / Screen.TwipsPerPixelX) / 2) - ((cWidth / Screen.TwipsPerPixelX) / 2), (((frmMain.ScaleHeight / Screen.TwipsPerPixelY) / 4) * 3) - ((cHeight / Screen.TwipsPerPixelY) / 2)

    End If
    
    If ((frmMain.IsPlayback And ((Not PauseGame) And TrapMouse And (Not ConsoleVisible))) Or (Not frmMain.IsPlayback)) Then
        txt = "Scores: " & TheScore
        DrawText txt, ((frmMain.ScaleWidth / Screen.TwipsPerPixelX) / 2) - ((frmMain.TextWidth(txt) / Screen.TwipsPerPixelX) / 2), Row(RowCount)
        If frmMain.Multiplayer Then frmMain.MyScoreStrand = TheScore
    End If

    DDevice.SetRenderState D3DRS_ZENABLE, 1
    
    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, RENDER_MAGFILTER
    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, RENDER_MINFILTER
End Sub

Public Property Get TheScore() As String
    
    If (GainMode = 1) Then
        TheScore = "1st " & DateDiff("s", CDate(MuseTime), GetNow) & "s - " & Score2nd
    ElseIf (GainMode = 2) Then
        TheScore = Score1st & " - 2nd " & DateDiff("s", CDate(SakeTime), GetNow) & "s 0s"
    ElseIf (GainMode = 3) Then
        TheScore = Score1st & " - 2nd " & WonTime & "s " & DateDiff("s", CDate(MuseTime), GetNow) & "s"
    ElseIf GainMode = 0 Then
        TheScore = Score1st & " - " & Score2nd
    End If
    
    If Player.Stalled Then
        SixTime = 1
        TheScore = TheScore & " - 3rd STALLED" & IIf(HandiCap, "&", "")
    Else
        If ((Not (ScoreClr = 0)) Or Player.Trails Or (Not (Player.Model = 0))) Or (SixTime = 1) Then
            TheScore = TheScore & " - 3rd "
            TheScore = TheScore & ScoreClr & IIf(Player.Model > 0, "$", "%") & " "
            TheScore = TheScore & Trim(Round(Player.MoveSpeed * 100, 0)) & IIf(Player.Trails, "#", "^") & IIf(HandiCap, "&", "")
            SixTime = 1
        End If
    End If
    Static ForDisp As String
    Select Case TriTime
        Case 1

            If ForDisp = "" Then ForDisp = "|"

            Static StartTim As Double
            
            If CDbl(GetTimer - StartTim) >= 0.05 Then
                StartTim = GetTimer
            
                Select Case ForDisp
                    Case "|"
                        ForDisp = "/"
                    Case "\"
                        ForDisp = "|"
                    Case "-"
                        ForDisp = "\"
                    Case "/"
                        ForDisp = "-"
                    Case "|"
                        ForDisp = "/"
                End Select
            End If
        Case Else
            ForDisp = ""
    End Select
    
    If (TriTime = 2) Or (ScoreFor = "") Then
        ScoreFor = "4th " & Trim(ForTime)
        If Not (ScoreFor = "") Then TheScore = Trim(TheScore) & " - " & Trim(ScoreFor)
    ElseIf (TriTime = 1) Then
        ScoreFor = "4th " & Trim(DateDiff("s", ForTime, GetNow)) & "s " & Trim(ForDisp)
        If Not (ScoreFor = "") Then TheScore = TheScore & " - " & Trim(ScoreFor)
    ElseIf (TriTime = 3) Then
        TheScore = TheScore & " - " & Trim(ScoreFor) & " " & Trim(DateDiff("s", ForTime, GetNow)) & "s"
    ElseIf (TriTime = 5) Then
        If Not (ScoreFor = "") Then TheScore = Trim(TheScore) & " - " & Trim(ScoreFor)
    End If
    
    If Not (ScoreXth = "") Then
        TheScore = ScoreXth & " - " & TheScore
    End If
    
    If Not (ScoreFith = "") Then
        TheScore = TheScore & " - " & ScoreFith
    End If
    
    If Not (ScoreIdol = "") Then
        TheScore = TheScore & " - 6th " & ScoreIdol & "@"
    End If
    
    If Not (ScoreNth = 0) Then
        TheScore = TheScore & " - Nth " & IIf(ScoreNth = -1, 0, ScoreNth) & "+"
    End If
    
End Property
