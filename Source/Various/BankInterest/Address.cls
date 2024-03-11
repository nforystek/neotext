VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Address"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private ErrorMsg As String

Private pAddresses As Collection
Private pShore As Object

Private Function Exists(ByVal Key As String) As Boolean
    On Error Resume Next
    Dim ks As String
    ks = pAddresses(Key)
    If Err.Number = 0 Then
        Exists = True
    Else
        Err.Clear
    End If
    On Error GoTo 0
End Function
Friend Property Get AddrCol() As Collection
    Set AddrCol = pAddresses
End Property

Public Property Get Shore()
    Set Shore = pShore
End Property
Public Property Set Shore(ByRef RHS As Object)
    Set pShore = RHS
End Property

Private Sub Class_Initialize()
    Set pAddresses = New Collection
End Sub

Private Sub Class_Terminate()
    Do Until pAddresses.Count = 0
        GlobalFree CLng(pAddresses(1))
        pAddresses.Remove 1
    Loop
    Set pAddresses = Nothing
End Sub

Private Sub SetLastError(ByVal Msg As String)
    ErrorMsg = Msg
End Sub
Private Sub ResetLastError()
    ErrorMsg = ""
End Sub

Public Function GetLastError() As String
    Static ShowPrior As String
    If ErrorMsg <> ShowPrior Then
        GetLastError = ErrorMsg
        ShowPrior = ErrorMsg
        ErrorMsg = ""
    ElseIf ShowPrior <> "" Then
        GetLastError = ShowPrior
    End If
End Function

Public Function UnFreeze(ByRef Addr As Long) As Long
    ResetLastError
    If Exists("A" & Trim(CStr(CLng(Addr)))) Then
                
        Dim hMem As Long
        hMem = GlobalUnlock(Addr)
        If hMem = 0 Then
            SetLastError "Error attempting to unfreeze the memory."
        Else
            UnFreeze = hMem
        End If
    Else
        SetLastError "Memory address specified is not pooled."
    End If
End Function
Public Function Freeze(ByRef Addr As Long) As Long
    ResetLastError
    If Exists("A" & Trim(CStr(CLng(Addr)))) Then
        Dim hMem As Long
        hMem = GlobalLock(Addr)
        If hMem <> Addr Then
            SetLastError "Error attempting to freese the memory."
        Else
            Freeze = hMem
        End If
    Else
        SetLastError "Memory address specified is not pooled."
    End If
End Function

Public Function ReAlloc(ByRef Addr As Long, ByVal Size As Long, ByVal flags As Long) As Long
    ResetLastError
    If Exists("A" & Trim(CStr(CLng(Addr)))) Then
        Dim hMem As Long
        hMem = GlobalReAlloc(Addr, Size, flags)
        If hMem = 0 Then
            SetLastError "Error attempting to reallocate the memory."
        Else
            pAddresses.Remove "A" & Trim(CStr(CLng(Addr)))
            If Not Exists("A" & Trim(CStr(CLng(hMem)))) Then
                pAddresses.Add hMem, "A" & Trim(CStr(CLng(hMem)))
            End If
            ReAlloc = hMem
        End If
    Else
        SetLastError "Memory address specified is not pooled."
    End If
End Function

Public Function Size(ByVal Addr As Long) As Long
    ResetLastError
    If Exists("A" & Trim(CStr(CLng(Addr)))) Then

        Size = GlobalSize(Addr)
    Else
        SetLastError "Memory address specified is not pooled."
    End If
End Function

Public Function Alloc(ByVal flags As Long, ByVal Size As Long) As Long
    ResetLastError
    Dim hMem As Long
    hMem = GlobalAlloc(flags, Size)
    If hMem = 0 Then
        SetLastError "Error attempting to allocate the memory."
    Else
        pAddresses.Add hMem, "A" & Trim(CStr(CLng(hMem)))
        Alloc = hMem
    End If
End Function

Public Function Free(ByRef Addr As Long) As Long
    ResetLastError
    If Exists("A" & Trim(CStr(CLng(Addr)))) Then
        If GlobalFree(Addr) <> 0 Then
            SetLastError "Error attempting to free the memory."
        Else
            pAddresses.Remove CStr("A" & Trim(CStr(CLng(Addr))))
        End If
    Else
        SetLastError "Memory address specified is not pooled."
    End If
End Function



Public Sub RtlMoveMemory(Dest, Source, ByVal Length As Long)
    ResetLastError
    On Error GoTo errcatch
    
    RtlMoveMemory_ Dest, Source, Length
    
    Exit Sub
errcatch:
    SetLastError "RtlMoveMemory Error: " & Err.Description
End Sub


Public Property Get Count() As Long
    Count = pAddresses.Count
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = pAddresses.[_NewEnum]
End Function


