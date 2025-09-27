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
    'I didn't write this portion of the code
    Dim No&, s As String * MAX_FILENAME_LEN
    Call GetVolumeInformation(Drive + ":\", s, MAX_FILENAME_LEN, _
    No, 0&, 0&, s, MAX_FILENAME_LEN)
    SerialNumber = No
End Function

Private Function HexEx(ByVal B As Long) As String
    'I didn't write this portion of the code
    Dim aa As String
    aa = Hex$(B)
    If Len(aa) < 2 Then
        aa = "0" & aa
    End If
    HexEx = aa
End Function

Private Function MacAddress() As String
    'I didn't write this portion of the code
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
        'this may exist a drawback, the network is required for this code
        'to return a mac address, even if you have a network card and, a
        'network mac address is used for a unique location distriubted
        Exit Function
    End If
    myNcb.ncb_buffer = pASTAT
    bRet = Netbios(myNcb)
    CopyMemory myASTAT, myNcb.ncb_buffer, Len(myASTAT)
    MacAddress = HexEx(myASTAT.adapt.adapter_address(0)) & "-" & HexEx(myASTAT.adapt.adapter_address(1)) & "-" & HexEx(myASTAT.adapt.adapter_address(2)) & "-" & HexEx(myASTAT.adapt.adapter_address(3)) & "-" & HexEx(myASTAT.adapt.adapter_address(4)) & "-" & HexEx(myASTAT.adapt.adapter_address(5))
    Call HeapFree(GetProcessHeap(), 0, pASTAT)
End Function


Private Sub Modify(ByRef Data As String, ByVal CharNumber As Byte, ByVal Modby As String, ByRef Direction As Boolean)
    'This sub accepts Data, and wil modify the character at CharNumber in Data, by a Modby, going in the Direction
    'A Modby is another character whose Asc() value is used as a count by which CharNumber in Data is changed, and
    'the Direction is whehter or not the change is Adding or Subtracting by Modby, when a char digit falls above 9
    'or below 0 then direction is changed, and Direction should presist through subsequent calls from the Callee.
    
    Dim item As String 'the character digit at CharNumber's position in Data
    item = CByte(Mid(Data, CharNumber, 1))
    
    Dim l As String 'left portion of Data up to and excluding the character at CharNumber
    If CharNumber > 1 Then l = Left(Data, CharNumber - 1)
    
    Dim r As String 'right portion of Data up to and excluding the character at CharNumber
    If CharNumber < Len(Data) Then r = Mid(Data, CharNumber + 1)
    
    Dim change As Byte 'the amount of Asc(Modby) to 0, that must change the digit
    change = Asc(Modby)

    Do While change > 0 'while we have change left
    
        'if item is at 0 or 9 and direction is such
        'that it will be -1 or 10 next, then redirect
        If item = 0 And (Not Direction) Then
            Direction = Not Direction
        ElseIf item = 9 And Direction Then
            Direction = Not Direction
        End If
        
        If Direction And item < 9 Then 'going up (true) and room to do so?
            item = item + 1
            change = change - 1
        ElseIf Not Direction And item > 0 Then 'going down (false) and room to do so?
            item = item - 1
            change = change - 1
        End If
    Loop
    
    'per every call, direction swap
    Direction = Not Direction
    
    Data = l & Trim(CStr(item)) & r 'reform the Data as teh return result
End Sub

Public Static Function Random() As Double
    Static toggleSet As Integer
    toggleSet = toggleSet + 1
    
    Dim ID As String 'Processor ID
    Dim UN As String 'User Name
    Dim CN As String 'Compuer Name
    Dim HD As String 'Home Drive
    Dim SR As String 'Drive Serial
    Dim MA As String 'Mac Address

    'gather elements that make a unique
    'band of information per any user
    'and computer location combination
    ID = (Environ("PROCESSOR_IDENTIFIER"))
    UN = (Environ("USERNAME"))
    CN = (Environ("COMPUTERNAME"))
    HD = (Environ("HOMEDRIVE"))
    SR = SerialNumber(Left(Environ("HOMEDRIVE"), 1))
    MA = MacAddress
    
    'in subsequent calls to Random, we shift
    'the order at which this band is formed
    Dim UniqueBand  As String
    Select Case toggleSet
        Case 1, -2
            UniqueBand = ID & UN & CN & HD & SR & MA
        Case 2, -3
            UniqueBand = UN & CN & HD & SR & MA & ID
        Case 3, -4
            UniqueBand = CN & HD & SR & MA & ID & UN
        Case 4, -5
            UniqueBand = HD & SR & MA & ID & UN & CN
        Case 5, -6
            UniqueBand = SR & MA & ID & UN & CN & HD
        Case 6, -1
            UniqueBand = MA & ID & UN & CN & HD & SR
    End Select
    UniqueBand = Replace(UniqueBand, " ", "") 'probably not nessiary
    
    'the toggleSet is staitc to ensure subsequent calls
    'this next line resets it to come again backwards
    If Abs(toggleSet) = 6 Or toggleSet = -1 Then toggleSet = -toggleSet
    
    Dim ReturnNum As String 'the secret data our band modifies
    ReturnNum = Replace(CStr(Timer), ".", "") 'shhhhhhhhhhhhhh
    
    'now blend the UniqueBand into
    'the value returned by Timer
    Dim Location As Byte
    Dim Direction As Boolean
    Do While UniqueBand <> ""
        Location = Location + 1
        If Location > Len(ReturnNum) Then Location = 1
        Modify ReturnNum, Location, Left(UniqueBand, 1), Direction
        UniqueBand = Mid(UniqueBand, 2)
    Loop
    
    'based on float precision of the timer
    If Len(Trim(CStr(CDbl("0." & ReturnNum)))) < 9 Then
        ReturnNum = String(9 - Len(Trim(CStr(CDbl("0." & ReturnNum)))), "0") & ReturnNum
    End If
    
    Random = CDbl("0." & Trim(ReturnNum)) 'build the final number
End Function

Public Sub Main()
    Dim cnt As Long
    For cnt = 1 To 10
        Debug.Print Random
    Next
    Debug.Print
End Sub




