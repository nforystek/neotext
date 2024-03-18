VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Exchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'the functions in this module cause bank downtime and accounts to lose their "instantly worth" quality between exchange when really is not
'therefore we add up the downtime for however long each function takes to the banks that are invovled, and repay interest to make up for it

Public Enum BANK_RETURNS
    Success = 0
    InvalidRouting = 1
    InvalidAccount = 2
    InvalidFunding = 3
End Enum

Public Function Withdraw(ByRef Routing As Double, ByRef Account As Double, ByVal Amount As Currency) As BANK_RETURNS
    'This function is used to deduct money from a bank account that is being exchanged into paper funds, i.e. ATM or bank teller withdraw
    ElapsedCounter
   
    Dim FromBank As Bank
    Set FromBank = Shore.FindBank(Routing)
    If Not FromBank Is Nothing Then
        Dim FromAcct As Account
        Set FromAcct = FromBank.FindAccount(FromBank, Account)
        If Not FromAcct Is Nothing Then
            Dim Funds As Currency
            GlobalLock FromBank.Routing
            LocalUnlock FromAcct.Number
            RtlMoveMemory Funds, FromAcct.Number, LenB(Funds)
            If Funds >= Amount Then
                Funds = Funds - Amount
                RtlMoveMemory FromAcct.Number, Funds, LenB(Funds)
            Else
                Withdraw = InvalidFunding
            End If
            LocalLock FromAcct.Number
            GlobalUnlock FromBank.Routing
        Else
            Withdraw = InvalidAccount
        End If
        
        FromBank.DownTime = FromBank.DownTime + ElapsedCounter
    Else
        Withdraw = InvalidRouting
        Shore.PassTime = Shore.PassTime + ElapsedCounter
    End If
    LogEvent "Withdraw", Withdraw, Amount, 0, 0, Routing, Account
End Function

Public Function Deposit(ByRef Routing As Double, ByRef Account As Double, ByVal Amount As Currency) As BANK_RETURNS
    'This function is used to deposit paper money into digital form adding to the value exhcnaged, i.e. depositing at the bank teller
    ElapsedCounter

    Dim ToBank As Bank
    Set ToBank = Shore.FindBank(Routing)
    If Not ToBank Is Nothing Then
        Dim ToAcct As Account
        Set ToAcct = ToBank.FindAccount(ToBank, Account)
        If Not ToAcct Is Nothing Then
            Dim Funds As Currency
            GlobalLock ToBank.Routing
            LocalUnlock ToAcct.Number
            RtlMoveMemory Funds, ToAcct.Number, LenB(Funds)
            Funds = Funds + Amount
            RtlMoveMemory ToAcct.Number, Funds, LenB(Funds)
            LocalLock ToAcct.Number
            GlobalUnlock ToBank.Routing
        Else
            Deposit = InvalidAccount
        End If
        ToBank.DownTime = ToBank.DownTime + ElapsedCounter
    Else
        Deposit = InvalidRouting
        Shore.PassTime = Shore.PassTime + ElapsedCounter
    End If
    LogEvent "Deposit", Deposit, Amount, Routing, Account, 0, 0
End Function


Public Function Transfer(ByRef RoutingFrom As Double, ByRef AccountFrom As Double, ByRef RoutingTo As Double, ByRef AccountTo As Double, ByVal Amount As Currency) As BANK_RETURNS
    'this function is used to transfer money from one account to another
    ElapsedCounter
    
    Dim FromBank As Bank
    Set FromBank = Shore.FindBank(RoutingFrom)
    Dim ToBank As Bank
    Set ToBank = Shore.FindBank(RoutingTo)
    If Not ToBank Is Nothing And Not FromBank Is Nothing Then
        Dim FromAcct As Account
        Set FromAcct = FromBank.FindAccount(FromBank, AccountFrom)
        If Not FromAcct Is Nothing Then
            Dim ToAcct As Account
            Set ToAcct = ToBank.FindAccount(ToBank, AccountTo)
            If Not ToAcct Is Nothing Then
                Dim Funds As Currency
                GlobalLock FromBank.Routing
                LocalUnlock FromAcct.Number
                RtlMoveMemory Funds, FromAcct.Number, LenB(Funds)
                If Funds >= Amount Then
                    Funds = Funds - Amount
                    RtlMoveMemory FromAcct.Number, Funds, LenB(Funds)
                    GlobalLock ToBank.Routing
                    LocalUnlock ToAcct.Number
                    RtlMoveMemory Funds, ToAcct.Number, LenB(Funds)
                    Funds = Funds + Amount
                    RtlMoveMemory ToAcct.Number, Funds, LenB(Funds)
                    LocalLock ToAcct.Number
                    GlobalUnlock ToBank.Routing
                Else
                    Transfer = InvalidFunding
                End If
                LocalLock FromAcct.Number
                GlobalUnlock FromBank.Routing
            Else
                Transfer = InvalidAccount
            End If
        Else
            Transfer = InvalidAccount
        End If
        
        Dim elap As Currency
        elap = ElapsedCounter
        FromBank.DownTime = FromBank.DownTime + elap
        ToBank.DownTime = ToBank.DownTime + elap
    Else
        Transfer = InvalidRouting
        Shore.PassTime = Shore.PassTime + ElapsedCounter
    End If
    LogEvent "Transfer", Transfer, Amount, RoutingTo, AccountTo, RoutingFrom, AccountFrom
End Function

Friend Sub LogEvent(ByVal method As String, ByVal result As BANK_RETURNS, ByVal Amount As Currency, ByVal BankTo As Double, ByVal AccountTo As Double, ByVal BankFrom As Double, ByVal AccountFrom As Double)

    Dim fNum As Long
    fNum = FreeFile
    If Not PathExists(GetLogFilename, True) Then
        Open GetLogFilename For Output As #fNum
        Print #fNum, "date,time,method,result,bank-to,account-to,bank-from,account-from,amount"
        Close #fNum
    End If
    
    Dim strEvents As String
    
    strEvents = Year(Now) & "-" & Month(Now) & "-" & Day(Now)
    strEvents = strEvents & "," & Format(Now, "hh:mm:ss")
    strEvents = strEvents & "," & method & "," & result

    If BankTo <> 0 Then
        strEvents = strEvents & "," & BankTo & "," & AccountTo
    Else
        strEvents = strEvents & ",,"
    End If

    If BankFrom <> 0 Then
        strEvents = strEvents & "," & BankFrom & "," & AccountFrom
    Else
        strEvents = strEvents & ",,"
    End If
    
    strEvents = strEvents & Amount

    fNum = FreeFile
    Open GetLogFilename For Append Shared As #fNum
        Print #fNum, strEvents
    Close #fNum
End Sub

Private Function GetLogFilename() As String
    GetLogFilename = SysPath & "LogFiles\NTACH\" & "ex" & Padding(2, Right(Year(Now), 2), "0") & Padding(2, Month(Now), "0") & Padding(2, Day(Now), "0") & ".log"
End Function

Private Function Padding(ByVal Length As Long, ByVal Value As String, Optional ByVal PadWith As String = " ") As String
    Padding = String(Abs((Length * Len(PadWith)) - (Len(Value) \ Len(PadWith))), PadWith) & Value
End Function

Private Sub Class_Initialize()

    If Not PathExists(SysPath & "LogFiles", False) Then MkDir SysPath & "LogFiles"
    If Not PathExists(SysPath & "LogFiles\NTACH", False) Then MkDir SysPath & "LogFiles\NTACH"

End Sub
