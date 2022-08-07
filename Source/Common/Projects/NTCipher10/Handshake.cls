VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Shaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private tsServer As String

Private someGlobal As String
Private someTiming As String
Private someOffset As String

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
    Serialize = n.EncryptString(someGlobal & "|" & someTiming & "|" & someOffset & "|" & myDates, "58344A80-C129-ADA2-AD29-C1804A345858")
    Set n = Nothing
End Function
Public Function Deserialize(ByVal Object As String, Optional ByVal Key As String)
    If IsMissing(Key) Then Key = "58344A80-C129-ADA2-AD29-C1804A345858"
    Dim n As New NTCipher10.NCode
    Object = n.DecryptString(Object, Key)
    someGlobal = RemoveNextArg(Object, "|")
    someTiming = RemoveNextArg(Object, "|")
    someOffset = RemoveNextArg(Object, "|")
    MyNow Object
    Set n = Nothing
End Function

Public Property Get Seed() As String ' _
Gets the Seeding valule for this object that is offset in calls to Checksum and time defined for use against another object's (or its self) to train valid sequence checks.
Attribute Seed.VB_Description = "Gets the Seeding valule for this object that is offset in calls to Checksum and time defined for use against another object's (or its self) to train valid sequence checks."
    Seed = IIf(someOffset = "", someTiming, someOffset)
End Property
Public Property Get Para() As String ' _
Gets the Parallel valule for this object that is globally defined and used with Seed in calls to Checksum for use against another object (or its self) to train valid sequence checks.
Attribute Para.VB_Description = "Gets the Parallel valule for this object that is globally defined and used with Seed in calls to Checksum for use against another object (or its self) to train valid sequence checks."
    Para = someGlobal
End Property
Private Property Get Offset() As String
    Offset = DateDiff("s", MyNow, someTiming)
End Property
Private Property Get Timing() As String
    Timing = DateAdd("s", someOffset, MyNow)
End Property

Public Function Checksum(ByVal Para As String, ByVal Seed As String) As Boolean ' _
Accepts arguments for the parallel and seeding value of another's object (or its self) and returns whether or not if the arguments passed validate prior Checksum combination calls in sequences.
Attribute Checksum.VB_Description = "Accepts arguments for the parallel and seeding value of another's object (or its self) and returns whether or not if the arguments passed validate prior Checksum combination calls in sequences."
    If (Seed = vbNullString) And someGlobal = "" Then someGlobal = modGuid.GUID
    If IsDate(Seed) And (someTiming = "") Then
        someOffset = DateDiff("s", MyNow, Seed)
    ElseIf IsNumeric(Seed) And (someOffset = "") Then
        someTiming = (Timer * 1000) - Seed
    ElseIf someOffset = "" And someTiming = "" Then
        someOffset = (Timer * 1000)
        someTiming = ""
        someGlobal = modGuid.GUID
        Checksum = True
    ElseIf Not (someOffset = "" Or someTiming = "") Then
        someOffset = DateDiff("s", MyNow, someTiming)
        someTiming = DateAdd("s", ((Timer * 1000) - someOffset), MyNow)
        someGlobal = modGuid.GUID
        Checksum = True
    End If
    
    If (someGlobal = Para) And (Para <> vbNullString) Then
        If someOffset = "" Then
            Checksum = True
            If IsNumeric(someTiming) Then
                someOffset = DateAdd("s", -someTiming, MyNow)
            Else
                someOffset = DateDiff("s", MyNow, someTiming)
            End If
            someTiming = ""
        ElseIf someTiming = "" Then
            Checksum = True
            someTiming = DateAdd("s", ((Timer * 1000) - someOffset), MyNow)
            someOffset = ""
        End If
    End If
End Function

Private Sub Class_Initialize()
    tsServer = "128.138.140.44:37"  '"time.windows.com:37" ' "time.nist.gov:37"
    Randomize
    someGlobal = modGuid.GUID
End Sub




