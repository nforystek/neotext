VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Shore"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pAddress As Long

Private pPassTime As Currency

'make our cache object controller and available events
Public WithEvents Controller As NTService20.Controller
Attribute Controller.VB_VarHelpID = -1

Private pBanks As NTNodes10.Collection

Friend Property Get Ordinal() As Long
    Ordinal = pAddress
End Property

Friend Property Get PassTime() As Currency
    PassTime = pPassTime
End Property
Friend Property Let PassTime(ByVal RHS As Currency)
    pPassTime = RHS
End Property

Friend Property Get Banks() As NTNodes10.Collection
    Set Banks = pBanks
End Property

Friend Function FindBank(ByVal Routing As Long) As Bank
    Dim Bank As Bank
    For Each Bank In Shore.Banks
        If Routing = Bank.Routing Then
            Set FindBank = Bank
            Exit Function
        End If
    Next
End Function

Public Function OpenBank() As Double
    Dim Bank As New Bank
    Banks.Add Bank, "R" & Bank.Routing
    OpenBank = Bank.Routing
    Set Bank = Nothing
End Function

Public Sub KillBank(ByVal Routing As Double)
    Dim cnt As Long
    cnt = 1
    Do While cnt <= Banks.Count
        If Routing = Banks(cnt).Routing Then
            Banks.Remove cnt
            Exit Do
        End If
        cnt = cnt + 1
    Loop
End Sub

Private Static Sub Class_Initialize()
    Set pBanks = New NTNodes10.Collection
    Set Controller = New NTService20.Controller 'lbind
    pAddress = HeapCreate(HEAP_NO_SERIALIZE, LenB(pPassTime), 0)
End Sub

Private Static Sub Class_Terminate()
    HeapDestroy pAddress
    Set Controller = Nothing 'properlly set to nothing
    Set pBanks = Nothing
End Sub

Private Sub Controller_ContinueService(Success As Boolean)
    Success = True 'always set success to your worries
End Sub

Private Sub Controller_PauseService(Success As Boolean)
    Success = True 'passed it back acknowledge success
End Sub

Private Sub Controller_StartService(Success As Boolean)
    Success = True 'this can occur solo and cause main
End Sub

Private Sub Controller_StopService()
    StopService 'call our premaid of sub main reversed
End Sub

