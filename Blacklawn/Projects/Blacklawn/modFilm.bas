#Const [True] = -1
#Const [False] = 0

Attribute VB_Name = "modFilm"
#Const modFilm = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module

Public FilmFileNum As Integer

Public Function GetFilmVersion() As String
    Dim inSet As String
    
    inSet = String(2, Chr(0))
    
    Get #FilmFileNum, 1, inSet

    GetFilmVersion = inSet
End Function

Public Function GetFilmUserData() As String
    Dim inSet As String
    Dim inLoc As Long
    
    inLoc = 1
    
    Do Until inSet = vbCr
        inLoc = inLoc + 1
        inSet = String(1, Chr(0))
        Get #FilmFileNum, inLoc, inSet
        
    Loop
    
    inSet = String(inLoc - 2, Chr(0))
    
    Get #FilmFileNum, 3, inSet

    GetFilmUserData = inSet
End Function

Public Function GetFilmNextControlData() As String
    Dim inSet As String
    Dim inLoc As Long
    Dim inData As String
    Dim inLof As Long
    
    inLof = LOF(FilmFileNum)
    inLoc = Seek(FilmFileNum)
    
    Do Until (inSet = "|") Or (inLoc > inLof)
        inData = inData & inSet
        inSet = String(1, Chr(0))
        Get #FilmFileNum, inLoc, inSet
        inLoc = inLoc + 1
    Loop
    
    If (inLoc > inLof) Then
        GetFilmNextControlData = ""
    Else
        GetFilmNextControlData = inData
    End If
End Function

Public Function GetFilmWarpData() As String

    Dim inSet As String
    Dim inLoc As Long
    
    inLoc = LOF(FilmFileNum)
    
    Do Until (inSet = vbLf) Or (inLoc < 2)
        inLoc = inLoc - 1
        inSet = String(1, Chr(0))
        Get #FilmFileNum, inLoc, inSet
    Loop
    
    inLoc = inLoc + 1
    
    inSet = String(LOF(FilmFileNum) - inLoc, Chr(0))
    
    Get #FilmFileNum, inLoc, inSet
    
    GetFilmWarpData = inSet

End Function

Public Sub SetFilmControlPointer()
    Dim inSet As String
    Dim inLoc As Long
    Dim inLof As Long
    
    inLof = LOF(FilmFileNum)
    inLoc = 1
    
    Do Until (inSet = vbCr) Or (inLoc > inLof)
        inSet = String(1, Chr(0))
        Get #FilmFileNum, inLoc, inSet
        inLoc = inLoc + 1
    Loop
    
    Seek #FilmFileNum, inLoc + 1

End Sub

Public Sub RecordFilm(ByVal inArg As String)
    Dim FileName As String
    
    If (Not frmMain.Multiplayer) Then
        If (Not frmMain.Recording) Then

            AddMessage "Recording..."
            If ConsoleVisible Then ConsoleToggle
            
            SaveUserData
            
            If Not PathExists(AppPath & "Films") Then MkDir AppPath & "Films"
            FileName = AppPath & "Films\" & IIf(NextArg(inArg, vbCrLf) = "", GetUserLoginName, NextArg(inArg, vbCrLf)) & ".blklwn"
            
            On Error GoTo recorderr
            
            FilmFileNum = FreeFile
            Open FileName For Output As #FilmFileNum
            Close #FilmFileNum
            Open FileName For Binary As #FilmFileNum
            
            UserData = Replace(GetUserData, GetUserLoginName, Replace(modGUID.GUID, "-",""))
            
            WarpData = ""
            ViewData = ""

            Put #FilmFileNum, 1, BlacklawnVer & UserData & vbCrLf
            
            StartNow = Now
            frmMain.IsPlayback = False
            frmMain.Recording = True
            
        Else
            AddMessage "Recording or playback is in progress."
        End If
    Else
        AddMessage "Unable to use recording or playback commands in multiplayer."
    End If
Exit Sub
recorderr:
    AddMessage "Error recording [" & FileName & "]"
    Err.Clear
End Sub

Public Sub StopFilm()
    If (Not frmMain.Multiplayer) Then
        If frmMain.Recording Then
            frmMain.Recording = False
            
            If Not frmMain.IsPlayback Then
                Put #FilmFileNum, , vbCrLf & WarpData
            End If
            
            Close #FilmFileNum
            
            AddMessage "Stopped..."
            
            If frmMain.IsPlayback Then
                frmMain.IsPlayback = False
                ClearUserData UserData
                LoadUserData
            End If
        Else
            AddMessage "Recording or playback is not in progress."
        End If
    Else
        AddMessage "Unable to use recording or playback commands in multiplayer."
    End If
End Sub

Public Sub PlayFilm(ByVal inArg As String)
    Dim FileName As String
    Dim fVersion As String
    
    If (Not frmMain.Multiplayer) Then
        If Not frmMain.Recording Then
            FileName = AppPath & "Films\" & IIf(NextArg(inArg, vbCrLf) = "", GetUserLoginName, NextArg(inArg, vbCrLf)) & ".blklwn"
            If PathExists(FileName, True) Then
                SaveUserData
                AddMessage "Loading... [" & FileName & "]"
                    
                On Error GoTo playerr
                
                FilmFileNum = FreeFile
                
                Open FileName For Binary As #FilmFileNum
        
                fVersion = GetFilmVersion
                
                If fVersion = BlacklawnVer Then
                    If ConsoleVisible Then ConsoleToggle
                    
                    UserData = GetFilmUserData
                    WarpData = GetFilmWarpData
                    ViewData = ""
                    
                    SetFilmControlPointer
                    
                    CleanupLawn
                    CreateLawn
                    
                    ClearUserData
                    LoadUserData UserData
                    
                    StartNow = Now
                    frmMain.IsPlayback = True
                    frmMain.Recording = True
                    
                Else
                    AddMessage "Invalid film verison... [" & FileName & "]"
                End If
            Else
                AddMessage "File not found... [" & FileName & "]"
            End If
        Else
            AddMessage "Recording or playback is in progress."
        End If
    Else
        AddMessage "Unable to use recording or playback commands in multiplayer."
    End If
Exit Sub
playerr:
    AddMessage "Error playing [" & FileName & "]"
    Err.Clear
End Sub