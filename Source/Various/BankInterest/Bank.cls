VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bank"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pDownTime As Currency

Private pAddress As Long

Private pAccounts As NTNodes10.Collection

Friend Property Get DownTime() As Currency
    DownTime = pDownTime
End Property
Friend Property Let DownTime(ByVal RHS As Currency)
    pDownTime = RHS
End Property

Friend Property Get Accounts() As NTNodes10.Collection
    Set Accounts = pAccounts
End Property

Friend Function FindAccount(ByRef Bank As Bank, ByVal Number As Double) As Account
    Dim Acct As Account
    For Each Acct In Bank.Accounts
        If Number = Acct.Number Then
            Set FindAccount = Acct
            Exit Function
        End If
    Next
End Function

Public Property Get Routing()
    Routing = pAddress
End Property

Public Function OpenAccount() As Double
    GlobalLock pAddress
    Dim Account As New Account
    Accounts.Add Account, "N" & Account.Number
    OpenAccount = Account.Number
    Set Account = Nothing
    GlobalUnlock pAddress
End Function

Public Sub KillAccount(ByVal Number As Double)
    GlobalLock pAddress
    Dim cnt As Long
    cnt = 1
    Do While cnt <= Accounts.Count
        If Number = Accounts(cnt).Number Then
            Accounts.Remove cnt
            Exit Do
        End If
        cnt = cnt + 1
    Loop
    GlobalUnlock pAddress
End Sub

Private Sub Class_Initialize()
    pAddress = GlobalAlloc(GMEM_SHARE, LenB(pDownTime))
    Set pAccounts = New NTNodes10.Collection
End Sub

Private Sub Class_Terminate()
    GlobalFree pAddress
    Set pAccounts = Nothing
End Sub
