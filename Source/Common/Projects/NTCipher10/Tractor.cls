VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tractor"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private tsServer As String

Private someTiming As String
Private someGlobal As String
Private somePhrase As String
Private someEquals As String
Private someHandle As String

Private Seeking As String
Private Elapsed As Single
Private StartAt As Date

Public Property Get TimeSync() As String
    TimeSync = tsServer
End Property
Public Property Let TimeSync(ByVal RHS As String)
    tsServer = RHS
End Property

Private Function MyNow(Optional ByVal Deserial As Variant) As String
    Static gotDate As String
    Static nowDate As String
    If IsMissing(Deserial) Then
        If gotDate = "" Then
            nowDate = Now
            Dim frmSync As frmTimeSync
            Set frmSync = New frmTimeSync
            Load frmSync
            frmSync.SetServer tsServer
            gotDate = frmSync.SynchronizeClock
            Unload frmSync
            Set frmSync = Nothing
        End If
        MyNow = DateAdd("s", -DateDiff("s", gotDate, nowDate), gotDate)
        nowDate = Now
    ElseIf Deserial = "" Then
        Deserial = gotDate & "|" & nowDate
    Else
        gotDate = RemoveNextArg(Deserial, "|")
        nowDate = RemoveNextArg(Deserial, "|")
    End If
End Function

Public Function Serialize(Optional ByVal Key As String) As String
    If IsMissing(Key) Then Key = "58344A80-C129-ADA2-AD29-C1804A345858"
    Dim n As New NTCipher10.NCode
    Dim myDates As String
    MyNow myDates
    Serialize = n.EncryptString(someTiming & "|" & someGlobal & "|" & somePhrase & "|" & someEquals & "|" & someHandle & "|" & myDates, Key)
    Set n = Nothing
End Function
Public Function Deserialize(ByVal Object As String, Optional ByVal Key As String)
    If IsMissing(Key) Then Key = "58344A80-C129-ADA2-AD29-C1804A345858"
    Dim n As New NTCipher10.NCode
    Object = n.DecryptString(Object, Key)
    someTiming = RemoveNextArg(Object, "|")
    someGlobal = RemoveNextArg(Object, "|")
    somePhrase = RemoveNextArg(Object, "|")
    someEquals = RemoveNextArg(Object, "|")
    someHandle = RemoveNextArg(Object, "|")
    MyNow Object
    Set n = Nothing
End Function

Public Static Property Get Seeding() As String
    Seeding = someHandle
End Property

Public Static Property Get Initial() As String
    Initial = someGlobal
End Property

Public Static Function CheckSum(Optional ByRef InitialSeeding As String = vbNullString) As Boolean
    If IsDate(InitialSeeding) Then
        someTiming = DateDiff("s", Initiate, InitialSeeding)
    ElseIf IsNumeric(InitialSeeding) Then
        someTiming = InitialSeeding
    Else
        someTiming = someHandle
    End If
    If someGlobal = vbNullString Then someGlobal = MyNow
    If someTiming = vbNullString Then someTiming = MyNow
    If somePhrase = vbNullString Then somePhrase = Rnd
    someTiming = DateDiff("s", someGlobal, FindTiming(True, someTiming, somePhrase))
    CheckSum = (someEquals = someTiming)
End Function

Public Static Function Initiate(Optional ByRef InitialSeeding As String = vbNullString) As String
    If IsDate(InitialSeeding) Then
        someTiming = InitialSeeding
    ElseIf IsNumeric(InitialSeeding) Then
        someTiming = DateAdd("s", InitialSeeding, MyNow)
    Else
        someTiming = MyNow
    End If
    If somePhrase = vbNullString Then somePhrase = Rnd
    someGlobal = FindTiming(False, someTiming, somePhrase)
    someEquals = someTiming
    Initiate = someGlobal
    someHandle = someEquals
End Function

Private Sub Class_Initialize()
    tsServer = "128.138.140.44:37" ' "time.windows.com:37" '"time.nist.gov:37"
    Randomize
    Initiate
End Sub

Private Function FindTiming(ByVal PartBit As Boolean, ByRef Informed As String, ByRef Passing As String) As String
    If Seeking = Passing Then
        Elapsed = Timer
    End If
    If PartBit Then
        FindTiming = Part1Timing(Informed, Passing)
    Else
        FindTiming = Part2Timing(Informed, Passing)
    End If
End Function

Private Function Part1Timing(ByRef Informed As String, ByRef Passing As String) As String
    If Seeking <> Passing Then
        StartAt = MyNow
    End If
    Informed = (((DateDiff("s", StartAt, MyNow) / 1000) + (Timer - Elapsed)) - Informed)
    Passing = Seeking
    Part1Timing = StartAt
End Function

Private Function Part2Timing(ByRef Informed As String, ByRef Passing As String) As String
    If Seeking <> Passing Then
        StartAt = Informed
    End If
    Seeking = Passing
    Informed = Round((DateDiff("s", StartAt, MyNow) / 1000) + (Timer - Elapsed), 0)
    Part2Timing = StartAt
End Function
