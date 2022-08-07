#Const [True] = -1
#Const [False] = 0
Attribute VB_Name = "modUserInfo"
#Const modUserInfo = -1
Option Explicit
'TOP DOWN

Option Compare Binary

Private Const TokenUser = 1

Private Const TOKEN_QUERY = &H8

Private Enum SID_NAME_USE
    SidTypeUser = 1
  SidTypeGroup = 2
  SidTypeDomain = 3
  SidTypeAlias = 4
  SidTypeWellKnownGroup = 5
  SidTypeDeletedAccount = 6
  SidTypeInvalid = 7
  SidTypeUnknown = 8
  SidTypeComputer = 9
  SidTypeLabe = 10
End Enum

Private Type SID_AND_ATTRIBUTES
    Sid As Long
    Attributes As Long
End Type
 
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Private Declare Function ConvertSidToStringSid Lib "advapi32.dll" Alias "ConvertSidToStringSidA" (ByVal lpSid As Long, lpString As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long

Private Declare Function OpenProcessToken Lib "advapi32" ( _
    ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, _
    TokenHandle As Long) As Long
 
Private Declare Function GetTokenInformation Lib "advapi32" ( _
    ByVal TokenHandle As Long, TokenInformationClass As Integer, _
    TokenInformation As Any, ByVal TokenInformationLength As Long, _
    ReturnLength As Long) As Long
 
Private Declare Function IsValidSid Lib "advapi32" (ByVal pSid As Long) As Long
 
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const PROCESS_QUERY_INFORMATION = &H400

Private Declare Function OpenProcess Lib "kernel32" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long

Private Const TOKEN_READ                    As Long = &H20008

Private Declare Function LookupAccountName Lib "advapi32.dll" Alias "LookupAccountNameA" (ByVal lpSystemName As String, ByVal lpAccountName As String, Sid As Long, cbSID As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Integer) As Long

Private Declare Function LookupAccountSid Lib "advapi32" Alias "LookupAccountSidA" (ByVal lpSystemName As String, ByVal Sid As Long, ByVal name As String, cbName As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Long) As Long

Private Declare Function GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetCompName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function GetUserLoginName() As String
    Dim sBuffer As String
    Dim lSize As Long
    sBuffer = String(255, Chr(0))
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    If lSize > 0 Then
        sBuffer = Left(sBuffer, lSize)
    End If
    sBuffer = Replace(sBuffer, Chr(0), "")
    If sBuffer = "" Then sBuffer = "LocalSystem"
    GetUserLoginName = sBuffer
End Function

Public Function GetNetworkName() As String
    Dim sBuffer As String
    Dim objNameSpace As Object
    Dim objDomain As Object
    Set objNameSpace = GetObject("WinNT:")
    For Each objDomain In objNameSpace
        sBuffer = objDomain.name
        Exit For
    Next
    If sBuffer = "" Then sBuffer = "Unknown"
    GetNetworkName = sBuffer
End Function

Public Function GetMachineName() As String
    Dim sBuffer As String
    Dim lSize As Long
    sBuffer = String(255, Chr(0))
    lSize = Len(sBuffer)
    Call GetCompName(sBuffer, lSize)
    If lSize > 0 Then
        sBuffer = Left(sBuffer, lSize)
    End If
    sBuffer = Replace(sBuffer, Chr(0), "")
    If sBuffer = "" Then sBuffer = "Unknown"
    GetMachineName = sBuffer
End Function

Private Function UserNameBySID(ByVal Sid As Long) As String
    Dim lpSystemName As String
    Dim name As String
    Dim cbName As Long
    Dim ReferencedDomainName As String
    Dim cbReferencedDomainName As Long
    Dim peUse As Long
    
    LookupAccountSid lpSystemName, Sid, name, cbName, ReferencedDomainName, cbReferencedDomainName, peUse
    name = String(cbName, Chr(0))
    ReferencedDomainName = String(cbReferencedDomainName, Chr(0))
    LookupAccountSid lpSystemName, Sid, name, cbName, ReferencedDomainName, cbReferencedDomainName, peUse
    name = Replace(name, Chr(0), "")
    ReferencedDomainName = Replace(ReferencedDomainName, Chr(0), "")
    If ReferencedDomainName = "" Then ReferencedDomainName = GetMachineName
    If name = "" Then name = GetUserLoginName
    UserNameBySID = Replace(ReferencedDomainName, Chr(0), "") & "\" & Replace(name, Chr(0), "")
End Function

Public Function GetSidString(ByVal Sid As Long) As String

    Dim sBuffer         As String
    Dim lpSid           As Long
    Dim lpString        As Long

    Call CopyMemory(lpSid, Sid, 4)
    If ConvertSidToStringSid(lpSid, lpString) Then
        sBuffer = Space(lstrlen(lpString))
        Call CopyMemory(ByVal sBuffer, ByVal lpString, Len(sBuffer))
        Call LocalFree(lpString)
        GetSidString = sBuffer
    End If

End Function

Public Function GetUserByProcessID(ByVal ProcessID As Long, Optional ByVal FormOfSID As Boolean = False) As String
    Dim hProc As Long
    Dim hToken As Long
    
    Dim BufferSize As Long
    Dim lResult As Long

    Dim tpSid As SID_AND_ATTRIBUTES
    Dim InfoBuffer() As Long
    
    hProc = OpenProcess(PROCESS_QUERY_INFORMATION, 0&, ProcessID)
    If hProc Then
        If OpenProcessToken(hProc, TOKEN_QUERY, hToken) Then
            Call GetTokenInformation(hToken, ByVal TokenUser, 0, 0, BufferSize)     ' Determine required buffer size
            If BufferSize Then
                ReDim InfoBuffer((BufferSize \ 4) - 1) As Long
                lResult = GetTokenInformation(hToken, ByVal TokenUser, InfoBuffer(0), BufferSize, BufferSize)
                If lResult <> 1 Then Exit Function
                Call RtlMoveMemory(tpSid, InfoBuffer(0), Len(tpSid))
                If IsValidSid(tpSid.Sid) Then
                    If FormOfSID Then
                        GetUserByProcessID = GetSidString(tpSid.Sid)
                    Else
                        GetUserByProcessID = UserNameBySID(tpSid.Sid)
                    End If
                End If
            End If

            lResult = CloseHandle(hToken)
        End If
        lResult = CloseHandle(hProc)
    End If

End Function




