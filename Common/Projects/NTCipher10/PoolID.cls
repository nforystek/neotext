VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PoolID"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary

Public Enum TimedSums
    NoTiming = 0
    CheckSums = 1
    BoundSums = 2
End Enum

Private Dividen As Single
Private LastGen As Single

Private NEncode As NCode

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private SumCheck As TimedSums

Public Static Property Get Pacerate() As TimedSums
    Pacerate = SumCheck
End Property
Public Static Property Let Pacerate(ByRef nval As TimedSums)
    SumCheck = nval
End Property
    
Public Function Generate(Optional ByVal ParaPhrase As String = "*") As String
    Dim tick As Long
    Dim start As Single
    Dim ohmz As String
    Dim dash As Byte
    Static sum As Boolean
    
    If (SumCheck = NoTiming) Then
        ohmz = NEncode.EncryptString(CStr(Timer \ Dividen), CStr(GetTickCount \ Dividen), False)
        ohmz = NEncode.EncryptString(ParaPhrase, NEncode.EncryptString(CStr(Timer), ohmz, False), False)
            
        ohmz = NEncode.EncryptString( _
            Int((9 * Rnd)) & Int((9 * Rnd)) & Int((9 * Rnd)) & _
            Int((9 * Rnd)) & Int((9 * Rnd)) & Int((9 * Rnd)) & _
            Int((9 * Rnd)) & Int((9 * Rnd)) & Int((9 * Rnd)) & _
            Int((9 * Rnd)) & Int((9 * Rnd)) & Int((9 * Rnd)) & _
            Int((9 * Rnd)) & Int((9 * Rnd)) & Int((9 * Rnd)) & _
            Int((9 * Rnd)), ohmz, True)
        
        dash = 1 + Int(2 + ((8 - 2) * Rnd))
        ohmz = Left(ohmz, dash - 1) & "-" & Mid(ohmz, dash)
        dash = dash + Int(3 + ((9 - 3) * Rnd))
        ohmz = Left(ohmz, dash - 1) & "-" & Mid(ohmz, dash)
        dash = dash + Int(3 + ((9 - 3) * Rnd))
        ohmz = Left(ohmz, dash - 1) & "-" & Mid(ohmz, dash)
        dash = dash + Int(3 + ((7 - 3) * Rnd))
        ohmz = Left(ohmz, dash - 1) & "-" & Mid(ohmz, dash)
        If Not (dash Mod 2 = 0) Then ohmz = StrReverse(ohmz)
        
        Generate = ohmz
    ElseIf (SumCheck = CheckSums) Then
        
        start = Timer
        tick = GetTickCount

        ohmz = NEncode.EncryptString(CStr(Timer \ Dividen), CStr(GetTickCount \ Dividen), False)
        ohmz = NEncode.EncryptString(ParaPhrase, NEncode.EncryptString(CStr(Timer), ohmz, False), False)
            
        ohmz = NEncode.EncryptString( _
            Int((9 * Rnd)) & Int((9 * Rnd)) & Int((9 * Rnd)) & _
            Int((9 * Rnd)) & Int((9 * Rnd)) & Int((9 * Rnd)) & _
            Int((9 * Rnd)) & Int((9 * Rnd)) & Int((9 * Rnd)) & _
            Int((9 * Rnd)) & Int((9 * Rnd)) & Int((9 * Rnd)) & _
            Int((9 * Rnd)) & Int((9 * Rnd)) & Int((9 * Rnd)) & _
            Int((9 * Rnd)), ohmz, True)
        
        dash = 1 + Int(2 + ((8 - 2) * Rnd))
        ohmz = Left(ohmz, dash - 1) & "-" & Mid(ohmz, dash)
        dash = dash + Int(3 + ((9 - 3) * Rnd))
        ohmz = Left(ohmz, dash - 1) & "-" & Mid(ohmz, dash)
        dash = dash + Int(3 + ((9 - 3) * Rnd))
        ohmz = Left(ohmz, dash - 1) & "-" & Mid(ohmz, dash)
        dash = dash + Int(3 + ((7 - 3) * Rnd))
        ohmz = Left(ohmz, dash - 1) & "-" & Mid(ohmz, dash)
        If Not (dash Mod 2 = 0) Then ohmz = StrReverse(ohmz)
        
        Generate = ohmz
        
        tick = CSng(GetTickCount - tick)
        start = CSng(Timer - start)
            
        If Not sum Then
            sum = True
                
            Static tsu As Single
            Static eta As Single
            If tsu > 0 Then Sleep tsu
            eta = eta - 1
            If (eta < 0) Then
                tsu = eta
                eta = CheckSum(ParaPhrase)
            End If
            sum = False
        End If

        LastGen = LastGen + Sqr(tick * start)
    ElseIf (SumCheck = BoundSums) Then

        If Not sum Then
            sum = True
            Sleep CheckSum(ParaPhrase)
            sum = False
        End If

        start = Timer
        tick = GetTickCount
        
        ohmz = NEncode.EncryptString(CStr(Timer \ Dividen), CStr(GetTickCount \ Dividen), False)
        ohmz = NEncode.EncryptString(ParaPhrase, NEncode.EncryptString(CStr(Timer), ohmz, False), False)
            
        ohmz = NEncode.EncryptString( _
            Int((9 * Rnd)) & Int((9 * Rnd)) & Int((9 * Rnd)) & _
            Int((9 * Rnd)) & Int((9 * Rnd)) & Int((9 * Rnd)) & _
            Int((9 * Rnd)) & Int((9 * Rnd)) & Int((9 * Rnd)) & _
            Int((9 * Rnd)) & Int((9 * Rnd)) & Int((9 * Rnd)) & _
            Int((9 * Rnd)) & Int((9 * Rnd)) & Int((9 * Rnd)) & _
            Int((9 * Rnd)), ohmz, True)
        
        dash = 1 + Int(2 + ((8 - 2) * Rnd))
        ohmz = Left(ohmz, dash - 1) & "-" & Mid(ohmz, dash)
        dash = dash + Int(3 + ((9 - 3) * Rnd))
        ohmz = Left(ohmz, dash - 1) & "-" & Mid(ohmz, dash)
        dash = dash + Int(3 + ((9 - 3) * Rnd))
        ohmz = Left(ohmz, dash - 1) & "-" & Mid(ohmz, dash)
        dash = dash + Int(3 + ((7 - 3) * Rnd))
        ohmz = Left(ohmz, dash - 1) & "-" & Mid(ohmz, dash)
        If Not (dash Mod 2 = 0) Then ohmz = StrReverse(ohmz)
        
        Generate = ohmz

        tick = CSng(GetTickCount - tick)
        start = CSng(Timer - start)

        LastGen = LastGen + Sqr(tick * start)

    End If

End Function

Private Function CheckSum(Optional ByVal ParaPhrase As String = "*") As Single
    Dim ohmz As String
    Dim hert As String
    Dim temp As String
        
    Dim tick As Long
    Dim start As Single
    If (SumCheck = NoTiming) Then

        start = Timer
        tick = GetTickCount

        hert = NEncode.EncryptString(CStr(start \ Dividen), CStr(GetTickCount \ Dividen), False)
        ohmz = NEncode.EncryptString(ParaPhrase, NEncode.EncryptString(CStr(Timer), hert, False), False)
        hert = NEncode.DecryptString(ohmz, hert, False)
    
        tick = CSng(GetTickCount - tick)
        start = CSng(Timer - start)

        CheckSum = Sqr(tick * start)
        
    ElseIf (SumCheck = CheckSums) Then
    
        start = Timer
        tick = GetTickCount
        
        Do
    
            hert = NEncode.EncryptString(CStr(start \ Dividen), CStr(GetTickCount \ Dividen), False)
            ohmz = NEncode.EncryptString(ParaPhrase, NEncode.EncryptString(CStr(Timer), hert, False), False)
    
            Static toggle As Boolean
            toggle = Not toggle
            If toggle Then hert = StrReverse(hert)
            hert = NEncode.DecryptString(ohmz, hert, False)
    
        Loop Until hert = ParaPhrase
    
        tick = CSng(GetTickCount - tick)
        start = CSng(Timer - start)
        If toggle Then
            LastGen = 0
            Generate Generate
            CheckSum = LastGen
        Else
            CheckSum = Sqr(tick * start)
        End If
    ElseIf (SumCheck = BoundSums) Then
    
        start = Timer
        tick = GetTickCount
        
        Static tolerance As Single
        If tolerance <= 0 Then tolerance = 0.3543
        Do Until hert = ParaPhrase Or (Timer - start) >= tolerance
    
            hert = NEncode.EncryptString(CStr(Timer \ Dividen), CStr(GetTickCount \ Dividen), False)
            ohmz = NEncode.EncryptString(ParaPhrase, NEncode.EncryptString(CStr(Timer), hert, False), False)
    
            hert = NEncode.DecryptString(ohmz, hert, False)
    
        Loop
    
        If Not (hert = ParaPhrase) And (Timer - start) >= tolerance Then
            tolerance = 0.0103
        ElseIf tolerance < 0.3543 Then
            tolerance = tolerance + 0.0199
        End If
    
        tick = CSng(GetTickCount - tick)
        start = CSng(Timer - start)
        
        If Sqr(tick * start) - LastGen < 0 Then
            LastGen = 0
            Generate Generate
            CheckSum = Sqr(tick * start)
        ElseIf LastGen > 0 Then
            CheckSum = Sqr(tick * start) - LastGen
            LastGen = (Sqr(tick * start) - LastGen) + Sqr(tick * start)
        End If
    
    End If

End Function

Private Sub Class_Initialize()
    Set NEncode = New NCode
    Dividen = GetTickCount / Timer
End Sub

Private Sub Class_Terminate()
    Set NEncode = Nothing
    Randomize
End Sub







