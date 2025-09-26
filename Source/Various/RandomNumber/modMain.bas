Attribute VB_Name = "Module1"
Option Explicit
'(c) Copyright by CC
'   Email: cyber_chris235@gmx.net
'
'Please mail me, when you want to use my code!


Private Const NCBASTAT                       As Long = &H33
Private Const NCBNAMSZ                       As Integer = 16
Private Const HEAP_ZERO_MEMORY               As Long = &H8
Private Const HEAP_GENERATE_EXCEPTIONS       As Long = &H4
Private Const NCBRESET                       As Long = &H32
Private Type NCB
    ncb_command                                As Byte
    ncb_retcode                                As Byte
    ncb_lsn                                    As Byte
    ncb_num                                    As Byte
    ncb_buffer                                 As Long
    ncb_length                                 As Integer
    ncb_callname                               As String * NCBNAMSZ
    ncb_name                                   As String * NCBNAMSZ
    ncb_rto                                    As Byte
    ncb_sto                                    As Byte
    ncb_post                                   As Long
    ncb_lana_num                               As Byte
    ncb_cmd_cplt                               As Byte
    ncb_reserve(9)                             As Byte
    ncb_event                                  As Long
End Type
Private Type ADAPTER_STATUS
    adapter_address(5)                         As Byte
    rev_major                                  As Byte
    reserved0                                  As Byte
    adapter_type                               As Byte
    rev_minor                                  As Byte
    duration                                   As Integer
    frmr_recv                                  As Integer
    frmr_xmit                                  As Integer
    iframe_recv_err                            As Integer
    xmit_aborts                                As Integer
    xmit_success                               As Long
    recv_success                               As Long
    iframe_xmit_err                            As Integer
    recv_buff_unavail                          As Integer
    t1_timeouts                                As Integer
    ti_timeouts                                As Integer
    Reserved1                                  As Long
    free_ncbs                                  As Integer
    max_cfg_ncbs                               As Integer
    max_ncbs                                   As Integer
    xmit_buf_unavail                           As Integer
    max_dgram_size                             As Integer
    pending_sess                               As Integer
    max_cfg_sess                               As Integer
    max_sess                                   As Integer
    max_sess_pkt_size                          As Integer
    name_count                                 As Integer
End Type
Private Type NAME_BUFFER
    name                                       As String * NCBNAMSZ
    name_num                                   As Integer
    name_flags                                 As Integer
End Type
Private Type ASTAT
    adapt                                      As ADAPTER_STATUS
    NameBuff(30)                               As NAME_BUFFER
End Type
Private Declare Function Netbios Lib "netapi32.dll" (pncb As NCB) As Byte
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, _
                                                                     ByVal hpvSource As Long, _
                                                                     ByVal cbCopy As Long)
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, _
                                                   ByVal dwFlags As Long, _
                                                   ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, _
                                                  ByVal dwFlags As Long, _
                                                  lpMem As Any) As Long

Private Declare Function GetVolumeInformation& Lib "kernel32" _
    Alias "GetVolumeInformationA" (ByVal lpRootPathName _
    As String, ByVal pVolumeNameBuffer As String, ByVal _
    nVolumeNameSize As Long, lpVolumeSerialNumber As Long, _
    lpMaximumComponentLength As Long, lpFileSystemFlags As _
    Long, ByVal lpFileSystemNameBuffer As String, ByVal _
    nFileSystemNameSize As Long)
    Const MAX_FILENAME_LEN = 256


Private Function SerialNumber(Drive$) As Long
    Dim No&, s As String * MAX_FILENAME_LEN
    Call GetVolumeInformation(Drive + ":\", s, MAX_FILENAME_LEN, _
    No, 0&, 0&, s, MAX_FILENAME_LEN)
    SerialNumber = No
End Function

Private Function HexEx(ByVal B As Long) As String
 
  Dim aa As String

    aa = Hex$(B)
    If Len(aa) < 2 Then
        aa = "0" & aa
    End If
    HexEx = aa

End Function

Private Function MacAddress() As String
    Dim bRet    As Byte
    Dim myNcb   As NCB
    Dim myASTAT As ASTAT
    Dim pASTAT  As Long
    myNcb.ncb_command = NCBRESET
    bRet = Netbios(myNcb)
    With myNcb
        .ncb_command = NCBASTAT
        .ncb_lana_num = 0
        .ncb_callname = "* "
        .ncb_length = Len(myASTAT)
        pASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS Or HEAP_ZERO_MEMORY, .ncb_length)
    End With
    If pASTAT = 0 Then
        MacAddress = "00-00-00-00-00-00"
        Exit Function
    End If
    myNcb.ncb_buffer = pASTAT
    bRet = Netbios(myNcb)
    CopyMemory myASTAT, myNcb.ncb_buffer, Len(myASTAT)
    MacAddress = HexEx(myASTAT.adapt.adapter_address(0)) & "-" & HexEx(myASTAT.adapt.adapter_address(1)) & "-" & HexEx(myASTAT.adapt.adapter_address(2)) & "-" & HexEx(myASTAT.adapt.adapter_address(3)) & "-" & HexEx(myASTAT.adapt.adapter_address(4)) & "-" & HexEx(myASTAT.adapt.adapter_address(5))
    Call HeapFree(GetProcessHeap(), 0, pASTAT)
End Function


Private Function Modify(ByRef Data As String, ByVal CharNumber As Byte, ByVal Modby As String, ByRef Direction As Boolean)
    Dim item As String
    Dim l As String
    Dim r As String
    Dim change As Byte
    If CharNumber > 1 Then
        l = left(Data, CharNumber - 1)
    End If
    If CharNumber < Len(Data) Then
        r = Mid(Data, CharNumber + 1)
    End If
    item = CByte(Mid(Data, CharNumber, 1))
    change = Asc(Modby)

    Do While change > 0
        If item = 0 And (Not Direction) Then
            Direction = Not Direction
        ElseIf item = 9 And Direction Then
            Direction = Not Direction
        End If
        If Direction And item < 9 Then
            item = item + 1
            change = change - 1
        ElseIf Not Direction And item > 0 Then
            item = item - 1
            change = change - 1
        Else
            Direction = Not Direction
        End If
        If change = 0 Then Direction = Not Direction
    Loop
    Data = l & Trim(CStr(item)) & r
End Function

Public Static Function Random() As Double
    Static toggleSet As Integer
    toggleSet = toggleSet + 1
    Dim ID As String
    Dim UN As String
    Dim CN As String
    Dim HD As String
    Dim SR As String
    Dim MA As String

    ID = (Environ("PROCESSOR_IDENTIFIER"))
    UN = (Environ("USERNAME"))
    CN = (Environ("COMPUTERNAME"))
    HD = (Environ("HOMEDRIVE"))
    SR = SerialNumber(left(Environ("HOMEDRIVE"), 1))
    MA = MacAddress
    
    Dim ran  As String
    Select Case Abs(toggleSet)
        Case 1
            ran = ID & UN & CN & HD & SR & MA
        Case 2
            ran = UN & CN & HD & SR & MA & ID
        Case 3
            ran = CN & HD & SR & MA & ID & UN
        Case 4
            ran = HD & SR & MA & ID & UN & CN
        Case 5
            ran = SR & MA & ID & UN & CN & HD
        Case 6
            ran = MA & ID & UN & CN & HD & SR
    End Select
    ran = Replace(ran, " ", "")
    
    If Abs(toggleSet) = 6 Then toggleSet = -toggleSet
    
    Dim out As String
    out = Replace(CStr(Timer), ".", "")
    
    Dim size As Byte
    size = Len(out)

    Dim Direction As Boolean
    Dim loc As Byte
    loc = 1
    Do While ran <> ""
    
        Modify out, loc, left(ran, 1), Direction
        ran = Mid(ran, 2)
        loc = loc + 1
        If loc > Len(out) Then loc = 1
    Loop
    
    If Len(Trim(out)) < size Then out = "0" & out
    Random = CDbl("0." & Trim(CStr(out)))

End Function

Public Sub Main()
    Dim cnt As Long
    For cnt = 1 To 10
        Debug.Print Random
    Next
    Debug.Print
End Sub




