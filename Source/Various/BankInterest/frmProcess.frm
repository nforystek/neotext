VERSION 5.00
Begin VB.Form frmProcess 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Securities"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2925
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProcess.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   92
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   195
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
End
Attribute VB_Name = "frmProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Interest As NTSchedule20.Schedule
Attribute Interest.VB_VarHelpID = -1
Public WithEvents Dividens As NTSchedule20.Schedule
Attribute Dividens.VB_VarHelpID = -1

Public MyPooling As Continent.Address

Private Sub Interest_ScheduledEvent()
    'every seven days calculate interest
    ElapsedCounter
    
    Dim Bank As Bank
    Dim Acct As Account
    Dim Funds As Currency
    Dim Inter As Currency
    
    For Each Bank In Shore.Banks
        For Each Acct In Bank.Accounts
        
            frmProcess.MyPooling.UnFreeze Acct.Number
            frmProcess.MyPooling.RtlMoveMemory Funds, Acct.Number, LenB(Funds)
            If Funds > 0 Then
                Inter = (Funds * Bank.DownTime)
                Funds = Funds + Inter
                frmProcess.MyPooling.RtlMoveMemory Acct.Number, Funds, LenB(Funds)
                Exchange.LogEvent "Interest", Success, Inter, Bank.Routing, Acct.Number, 0, 0
            End If
            frmProcess.MyPooling.Freeze Acct.Number

        Next
        Bank.DownTime = 0
    Next
    
    Shore.PassTime = Shore.PassTime + ElapsedCounter

End Sub

Private Sub Dividens_ScheduledEvent()
    'every 226 days calculate dividens
    ElapsedCounter
    
    Dim Bank As Bank
    Dim Acct As Account
    Dim Funds As Currency
    Dim Divden As Currency
    
    For Each Bank In Shore.Banks
        For Each Acct In Bank.Accounts
        
            frmProcess.MyPooling.UnFreeze Acct.Number
            frmProcess.MyPooling.RtlMoveMemory Funds, Acct.Number, LenB(Funds)
            If Funds > 0 Then
                Divden = (Funds * Shore.PassTime)
                Funds = Funds + Divden
                frmProcess.MyPooling.RtlMoveMemory Acct.Number, Funds, LenB(Funds)
                Exchange.LogEvent "Dividens", Success, Divden, Bank.Routing, Acct.Number, 0, 0
            End If
            frmProcess.MyPooling.Freeze Acct.Number

        Next
    Next
    
    Shore.PassTime = ElapsedCounter
    
End Sub

Private Sub Form_Load()

    Set Interest = New NTSchedule20.Schedule
    Interest.IncrementType = IncrementTypes.Day
    Interest.IncrementInterval = 7
    Interest.ScheduleType = Increment
    Interest.ExecuteDate = Date
    Interest.ExecuteDate = Now
    
    Set Dividens = New NTSchedule20.Schedule
    Dividens.IncrementType = IncrementTypes.Day
    Dividens.IncrementInterval = 226
    Dividens.ScheduleType = Increment
    Dividens.ExecuteDate = Date
    Dividens.ExecuteDate = Now
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set Interest = Nothing
    Set Dividens = Nothing
    
    Do Until MyPooling.Count = 0

        GlobalFree CLng(MyPooling.AddrCol(1))
        MyPooling.AddrCol.Remove 1

    Loop
    
    Set MyPooling = Nothing
End Sub
