VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Handshake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const SyncWithInternetTime = False

Private someGlobal As String
Private someTiming As String
Private someOffset As String
   
Private Function MyNow() As String
    Static gotDate As String
    Static nowDate As String
    If gotDate = "" Then
        nowDate = Now
        #If SyncWithInternetTime Then
            Dim frmSync As frmTimeSync
            Set frmSync = New frmTimeSync
            Load frmSync
            gotDate = frmSync.SynchronizeClock
            Unload frmSync
            Set frmSync = Nothing
        #Else
            gotDate = Month(Now) & "/" & Day(Now) & "/" & Year(Now) & " " & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
        #End If
    End If
    MyNow = DateAdd("s", -DateDiff("s", gotDate, nowDate), gotDate)
    nowDate = Now
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

Private Sub Backup(Optional ByVal sGlobal As String = "(none)", Optional ByVal sTiming As String = "(none)", Optional ByVal sOffset As String = "(none)")
    Static buGlobal As String
    Static buTiming As String
    Static buOffset As String
    
    If sGlobal <> "(none)" Then
        buGlobal = sGlobal
    Else
        someGlobal = buGlobal
    End If
    If sTiming <> "(none)" Then
        buTiming = sTiming
    Else
        someTiming = buTiming
    End If
    If sOffset <> "(none)" Then
        buOffset = sOffset
    Else
        someOffset = buOffset
    End If
End Sub

Public Function Checksum(ByVal Para As String, ByVal Seed As String) As Boolean ' _
Accepts arguments for the parallel and seeding value of another's object (or its self) and returns whether or not if the arguments passed validate prior Checksum combination calls in sequences.
Attribute Checksum.VB_Description = "Accepts arguments for the parallel and seeding value of another's object (or its self) and returns whether or not if the arguments passed validate prior Checksum combination calls in sequences."
    Backup someGlobal, someTiming, someOffset
    
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

Public Sub Rollback() ' _
Rolls back the entire object's state and values to just before the last most recent call to the CheckSum method.  System clocks, or sync of the system clock, is not affected by this method.
Attribute Rollback.VB_Description = "Rolls back the entire object's state and values to just before the last most recent call to the CheckSum method.  System clocks, or sync of the system clock, is not affected by this method."
    Backup
End Sub

Private Sub Class_Initialize()
    Randomize
    someGlobal = modGuid.GUID
End Sub


