#Const [True] = -1
#Const [False] = 0
Attribute VB_Name = "modWin32Ver"

#Const modWin32Ver = -1
Option Explicit
'TOP DOWN

Option Compare Binary

'#####################################################################################
'#  Determine the Win32 Operating System Version Via API (modWin32Ver.bas)
'#      By: Nick Campbeln
'#
'#      Revision History:
'#          1.0.2 (Aug 11, 2002):
'#              Switched GetVersionEx() form Public to Private
'#          1.0.1 (Aug 6, 2002):
'#              Fixed a (very) stupid coding error in isWin2k() - Renamed function from isWin2000() to isWin2k() and forgot to change the return values in the function to the same name - D'oh!
'#          1.0 (Aug 4, 2002):
'#              Initial Release
'#
'#      Copyright © 2002 Nick Campbeln (opensource@nick.campbeln.com)
'#          This source code is provided 'as-is', without any express or implied warranty. In no event will the author(s) be held liable for any damages arising from the use of this source code. Permission is granted to anyone to use this source code for any purpose, including commercial applications, and to alter it and redistribute it freely, subject to the following restrictions:
'#          1. The origin of this source code must not be misrepresented; you must not claim that you wrote the original source code. If you use this source code in a product, an acknowledgment in the product documentation would be appreciated but is not required.
'#          2. Altered source versions must be plainly marked as such, and must not be misrepresented as being the original source code.
'#          3. This notice may not be removed or altered from any source distribution.
'#              (NOTE: This license is borrowed from zLib.)
'#
'#  Please remember to vote on PSC.com if you like this code!
'#  Code URL: http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=37628&lngWId=1
'#####################################################################################
'# modifications by nick forystek as well contributes to this piticular source viewing

    '#### Functions/Consts/Types used for Win32Ver()
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Type OSVERSIONINFO
    OSVSize As Long
    dwVerMajor As Long
    dwVerMinor As Long
    dwBuildNumber As Long           '#### NT: Build Number, 9x: High-Order has Major/Minor ver, Low-Order has build
    PlatformID As Long
    szCSDVersion As String * 128    '#### NT: ie- "Service Pack 3", 9x: 'arbitrary additional information'
End Type

Public Enum cnWin32Ver
    UnknownOS = 0
    Win95 = 1
    Win98 = 2
    WinME = 3
    WinNT4 = 4
    Win2k = 5
    WinXP = 6
End Enum

'Private Enum OSSERIALFLAGS
'    AllFields = 0
'    CPUDescription = 1
'    CPUManufacturer = 2
'    CPUProcessorID = 4
'    MainBoardSerialNumber = 8
'    MainBoardDescription = 16
'    MainBoardManufacturer = 32
'    BIOSManufacturer = 64
'    OSSerialNumber = 128
'    OSDescription = 256
'    OSStartupDisk = 512
'End Enum
'
'Public Function WinSerialInfo(Optional ByVal Flags As OSSERIALFLAGS) As String()
'    Dim arr As Long
'
'    If (Flags And CPUDescription) = CPUDescription Then arr = arr + 1
'    If (Flags And CPUManufacturer) = CPUManufacturer Then arr = arr + 1
'    If (Flags And CPUProcessorID) = CPUProcessorID Then arr = arr + 1
'    If (Flags And MainBoardSerialNumber) = MainBoardSerialNumber Then arr = arr + 1
'    If (Flags And MainBoardDescription) = MainBoardDescription Then arr = arr + 1
'    If (Flags And MainBoardManufacturer) = MainBoardManufacturer Then arr = arr + 1
'    If (Flags And BIOSManufacturer) = BIOSManufacturer Then arr = arr + 1
'    If (Flags And OSSerialNumber) = OSSerialNumber Then arr = arr + 1
'    If (Flags And OSDescription) = OSDescription Then arr = arr + 1
'    If (Flags And OSStartupDisk) = OSStartupDisk Then arr = arr + 1
'
'
'    Dim SWbemSet(arr) As SWbemObjectSet
'    Dim SWbemObj As SWbemObject
'    Dim varObjectToId(arr) As String
'    Dim varSerial(arr) As String
'    Dim i, j As Integer
'    On Error Resume Next
'    i = 1
'
'    If (Flags = 0) Or ((Flags And CPUDescription) = CPUDescription) Then
'        varObjectToId(i) = "Win32_Processor,Name"
'        i = i + 1
'    End If
'    If (Flags = 0) Or ((Flags And CPUManufacturer) = CPUManufacturer) Then
'        varObjectToId(i) = "Win32_Processor,Manufacturer"
'        i = i + 1
'    End If
'    If (Flags = 0) Or ((Flags And CPUProcessorID) = CPUProcessorID) Then
'        varObjectToId(i) = "Win32_Processor,ProcessorId"
'        i = i + 1
'    End If
'    If (Flags = 0) Or ((Flags And MainBoardSerialNumber) = MainBoardSerialNumber) Then
'        varObjectToId(i) = "Win32_BaseBoard,SerialNumber"
'        i = i + 1
'    End If
'    If (Flags = 0) Or ((Flags And MainBoardDescription) = MainBoardDescription) Then
'        varObjectToId(i) = "Win32_Baseboard,product"
'        i = i + 1
'    End If
'    If (Flags = 0) Or ((Flags And MainBoardManufacturer) = MainBoardManufacturer) Then
'        varObjectToId(i) = "Win32_BaseBoard,manufacturer"
'        i = i + 1
'    End If
'    If (Flags = 0) Or ((Flags And BIOSManufacturer) = BIOSManufacturer) Then
'        varObjectToId(i) = "Win32_BIOS,Manufacturer"
'        i = i + 1
'    End If
'    If (Flags = 0) Or ((Flags And OSSerialNumber) = OSSerialNumber) Then
'        varObjectToId(i) = "Win32_OperatingSystem,SerialNumber"
'        i = i + 1
'    End If
'    If (Flags = 0) Or ((Flags And OSDescription) = OSDescription) Then
'        varObjectToId(i) = "Win32_OperatingSystem,Caption"
'        i = i + 1
'    End If
'    If (Flags = 0) Or ((Flags And OSStartupDisk) = OSStartupDisk) Then
'        varObjectToId(i) = "Win32_DiskDrive,Model"
'        i = i + 1
'    End If
'
'    For i = 1 To arr
'        Set SWbemSet(i) = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf(Split(varObjectToId(i), ",")(0))
'        varSerial(i) = ""
'        For Each SWbemObj In SWbemSet(i)
'            varSerial(i) = SWbemObj.Properties_(Split(varObjectToId(i), ",")(1)) 'Property value
'            varSerial(i) = Trim(varSerial(i))
'            If Len(varSerial(i)) < 1 Then varSerial(i) = "Unknown value"
'        Next
'        Text1(i) = varSerial(i)
'    Next
'
'    WinSerialInfo = varSerial
'
'End Function


Public Function WinVerInfo() As String
    Dim oOSV As OSVERSIONINFO
    oOSV.OSVSize = Len(oOSV)
    If GetVersionEx(oOSV) = 1 Then
        
        Select Case Trim(CStr(oOSV.dwVerMajor)) & "." & Trim(CStr(oOSV.dwVerMinor)) & "." & Trim(CStr(oOSV.dwBuildNumber))
        
            Case "4.0.950"
                WinVerInfo = "Windows 95 OEM Service Release 1 (95A) 4.00.950"
            Case "4.0.1111"
                WinVerInfo = "Windows 95 OEM Service Release 2 (95B) 4.00.1111"
            Case "4.3.1212"
                WinVerInfo = "Windows 95 OEM Service Release 2.1 4.03.1212"
            Case "4.3.1213"
                WinVerInfo = "Windows 95 OEM Service Release 2.1 4.03.1213"
            Case "4.3.1214"
                WinVerInfo = "Windows 95 OEM Service Release 2.1 4.03.1214"
            Case "4.3.1214"
                WinVerInfo = "Windows 95 OEM Service Release 2.5 C   4.03.1214"
            Case "4.10.1998"
                WinVerInfo = "Windows 98 4.10.1998"
            Case "4.10.2222" '"4.10.2222 A"
                WinVerInfo = "Windows 98 Second Edition (SE) 4.10.2222 A"
            Case "4.90.2476"
                WinVerInfo = "Windows Millenium Beta 4.90.2476"
            Case "4.90.3000"
                WinVerInfo = "Windows Millenium  4.90.3000"
            Case "3.10.528"
                WinVerInfo = "Windows NT 3.1 3.10.528"
            Case "3.50.807"
                WinVerInfo = "Windows NT 3.5 3.50.807"
            Case "3.51.1057"
                WinVerInfo = "Windows NT 3.51    3.51.1057"
            Case "4.0.1381"
                WinVerInfo = "Windows NT 4.00    4.00.1381"
            Case "5.0.1515"
                WinVerInfo = "Windows NT 5.00 (Beta 2)   5.00.1515"
            Case "5.0.2031"
                WinVerInfo = "Windows 2000 (Beta 3)  5.00.2031"
            Case "5.0.2128"
                WinVerInfo = "Windows 2000 (Beta 3 RC2)  5.00.2128"
            Case "5.0.2183"
                WinVerInfo = "Windows 2000 (Beta 3)  5.00.2183"
            Case "5.0.2195"
                WinVerInfo = "Windows 2000   5.00.2195"
            Case "5.0.2250"
                WinVerInfo = "Windows (Whistler Server Preview)  2250"
            Case "5.0.2257"
                WinVerInfo = "Windows (Whistler Server alpha)    2257"
            Case "5.0.2267"
                WinVerInfo = "Windows (Whistler Server interim release)  2267"
            Case "5.0.2410"
                WinVerInfo = "Windows (Whistler Server interim release)  2410"
            Case "5.1.2505"
                WinVerInfo = "Windows XP (RC 1)  5.1.2505"
            Case "5.1.2600"
                WinVerInfo = "Windows XP 5.1.2600"
'            Case "5.1.2600" '"5.1.2600.1105"
                'WinVerInfo = "Windows XP, Service Pack 1 5.1.2600.1105
'            Case "5.1.2600" '"5.1.2600.1106"
                'WinVerInfo = "Windows XP, Service Pack 1 5.1.2600.1106
'            Case "5.1.2600" '"5.1.2600.2180"
                'WinVerInfo = "Windows XP, Service Pack 2 5.1.2600.2180
'            Case "5.1.2600"
                'WinVerInfo = "Windows XP, Service Pack 3 5.1.2600
            Case "5.2.3541"
                WinVerInfo = "Windows .NET Server interim    5.2.3541"
            Case "5.2.3590"
                WinVerInfo = "Windows .NET Server Beta 3 5.2.3590"
            Case "5.2.3660"
                WinVerInfo = "Windows .NET Server Release Candidate 1    5.2.3660"
            Case "5.2.3718"
                WinVerInfo = "Windows .NET Server 2003 RC2   5.2.3718"
            Case "5.2.3763"
                WinVerInfo = "Windows Server 2003 (Beta?)    5.2.3763"
            Case "5.2.3790"
                WinVerInfo = "Windows Server 2003    5.2.3790"
'            Case "5.2.3790" '"5.2.3790.1180"
                'WinVerInfo = "Windows Server 2003, Service Pack 1    5.2.3790.1180
'            Case "5.2.3790" '"5.2.3790.1218"
                'WinVerInfo = "Windows Server 2003    5.2.3790.1218
            Case "5.2.3790"
                WinVerInfo = "Windows Home Server    5.2.3790"
            Case "6.0.5048"
                WinVerInfo = "Windows Longhorn   6.0.5048"
            Case "6.0.5112"
                WinVerInfo = "Windows Vista, Beta 1  6.0.5112"
            Case "6.0.5219"
                WinVerInfo = "Windows Vista, Community Technology Preview    6.0.5219"
            Case "6.0.5259"
                WinVerInfo = "Windows Vista, TAP Preview 6.0.5259"
            Case "6.0.5270"
                WinVerInfo = "Windows Vista, CTP 6.0.5270"
            Case "6.0.5308"
                WinVerInfo = "Windows Vista, CTP 6.0.5308"
            Case "6.0.5342"
                WinVerInfo = "Windows Vista, CTP (Refresh)   6.0.5342"
            Case "6.0.5365"
                WinVerInfo = "Windows Vista, April EWD   6.0.5365"
            Case "6.0.5381"
                WinVerInfo = "Windows Vista, Beta 2 Preview  6.0.5381"
            Case "6.0.5384"
                WinVerInfo = "Windows Vista, Beta 2  6.0.5384"
            Case "6.0.5456"
                WinVerInfo = "Windows Vista, Pre-RC1 6.0.5456"
            Case "6.0.5472"
                WinVerInfo = "Windows Vista, Pre-RC1, Build 5472 6.0.5472"
            Case "6.0.5536"
                WinVerInfo = "Windows Vista, Pre-RC1, Build 5536 6.0.5536"
            Case "6.0.5600" '"6.0.5600.16384"
                WinVerInfo = "Windows Vista, RC1 6.0.5600.16384"
            Case "6.0.5700"
                WinVerInfo = "Windows Vista, Pre-RC2 6.0.5700"
            Case "6.0.5728"
                WinVerInfo = "Windows Vista, Pre-RC2, Build 5728 6.0.5728"
            Case "6.0.5744" '"6.0.5744.16384"
                WinVerInfo = "Windows Vista, RC2 6.0.5744.16384"
            Case "6.0.5808"
                WinVerInfo = "Windows Vista, Pre-RTM, Build 5808 6.0.5808"
            Case "6.0.5824"
                WinVerInfo = "Windows Vista, Pre-RTM, Build 5824 6.0.5824"
            Case "6.0.5840"
                WinVerInfo = "Windows Vista, Pre-RTM, Build 5840 6.0.5840"
            Case "6.0.6000" '"6.0.6000.16386"
                WinVerInfo = "Windows Vista, RTM 6.0.6000.16386"
            Case "6.0.6000"
                WinVerInfo = "Windows Vista  6.0.6000"
            Case "6.0.6002"
                WinVerInfo = "Windows Vista, Service Pack 2  6.0.6002"
            Case "6.0.6001"
                WinVerInfo = "Windows Server 2008    6.0.6001"
            Case "6.1.7600" '"6.1.7600.16385"
                WinVerInfo = "Windows 7, RTM 6.1.7600.16385"
            Case "6.1.7601"
                WinVerInfo = "Windows 7  6.1.7601"
            Case "6.1.7600" '"6.1.7600.16385"
                WinVerInfo = "Windows Server 2008 R2, RTM    6.1.7600.16385"
'            Case "6.1.7601"
                'WinVerInfo = "Windows Server 2008 R2, SP1    6.1.7601
            Case "6.1.8400"
                WinVerInfo = "Windows Home Server 2011   6.1.8400"
            Case "6.2.9200"
                WinVerInfo = "Windows Server 2012    6.2.9200"
            Case "6.2.9200"
                WinVerInfo = "Windows 8  6.2.9200"
            Case "'6.2.10211"
                WinVerInfo = "Windows Phone 8    6.2.10211"
            Case "6.3.9200"
                WinVerInfo = "Windows Server 2012 R2 6.3.9200"
            Case "6.3.9200"
                WinVerInfo = "Windows 8.1    6.3.9200"
            Case "'6.3.9600"
                WinVerInfo = "Windows 8.1, Update 1  6.3.9600"
            Case "10.0.10240"
                WinVerInfo = "Windows 10 10.0.10240"
            Case "6.3.9600"
                WinVerInfo = "Windows Server 2012 R2 6.3.9600"
        End Select
    End If
End Function
'#####################################################################################
'# Public subs/functions
'#####################################################################################
'#########################################################
'# Returns the asso. cnWin32Ver eNum value of the current Win32 OS
'#########################################################
Public Function Win32Ver() As cnWin32Ver
    Dim oOSV As OSVERSIONINFO
    oOSV.OSVSize = Len(oOSV)
   
        '#### If the API returned a valid value
    If GetVersionEx(oOSV) = 1 Then
            '#### If we're running WinXP
            '####    If VER_PLATFORM_WIN32_NT, dwVerMajor is 5 and dwVerMinor is 1, it's WinXP
        If (oOSV.PlatformID = VER_PLATFORM_WIN32_NT And oOSV.dwVerMajor = 5 And oOSV.dwVerMinor = 1) Then
           Win32Ver = WinXP

            '#### If we're running WinNT2000 (NT5)
            '####    If VER_PLATFORM_WIN32_NT, dwVerMajor is 5 and dwVerMinor is 0, it's Win2k
        ElseIf (oOSV.PlatformID = VER_PLATFORM_WIN32_NT And oOSV.dwVerMajor = 5 And oOSV.dwVerMinor = 0) Then
           Win32Ver = Win2k

            '#### If we're running WinNT4
            '####    If VER_PLATFORM_WIN32_NT and dwVerMajor is 4
        ElseIf (oOSV.PlatformID = VER_PLATFORM_WIN32_NT And oOSV.dwVerMajor = 4) Then
           Win32Ver = WinNT4

            '#### If we're running Windows ME
            '####    If VER_PLATFORM_WIN32_WINDOWS and
            '####    dwVerMajor = 4,  and dwVerMinor > 0, return true
        ElseIf (oOSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS And oOSV.dwVerMajor = 4 And oOSV.dwVerMinor = 90) Then
           Win32Ver = WinME

            '#### If we're running Win98
            '####    If VER_PLATFORM_WIN32_WINDOWS and
            '####    dwVerMajor => 4, or dwVerMajor = 4 and
            '####    dwVerMinor > 0, return true
        ElseIf (oOSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And (oOSV.dwVerMajor > 4) Or (oOSV.dwVerMajor = 4 And oOSV.dwVerMinor > 0) Then
           Win32Ver = Win98

            '#### If we're running Win95
            '####    If VER_PLATFORM_WIN32_WINDOWS and
            '####    dwVerMajor = 4, and dwVerMinor = 0,
        ElseIf (oOSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS And oOSV.dwVerMajor = 4 And oOSV.dwVerMinor = 0) Then
           Win32Ver = Win95

            '#### Else the OS is not reconized by this function
        Else
            Win32Ver = UnknownOS
        End If
    
        '#### Else the OS is not reconized by this function
    Else
        Win32Ver = UnknownOS
    End If
End Function


'#########################################################
'# Returns true if the OS is WinNT4, Win2k or WinXP
'#########################################################
Public Function isNT() As Boolean
        '#### Determine the return value of Win32Ver() and set the return value accordingly
        '#### This is intristic upon how the OS handles services as in a "WinNT" model
    Select Case Win32Ver()
        Case Win95, Win98, WinME
            isNT = False
        Case Else
            isNT = True
    End Select
End Function


'#########################################################
'# Returns true if the OS is Win95, Win98 or WinME
'#########################################################
Public Function is9x() As Boolean
        '#### Determine the return value of Win32Ver() and set the return value accordingly
        '#### This is intristic upon how the OS handles services as in a "MSDOS" model
    Select Case Win32Ver()
        Case Win95, Win98, WinME
            is9x = True
        Case Else
            is9x = False
    End Select
End Function


'#########################################################
'# Returns true if the OS is WinXP
'#########################################################
Public Function isWinXP() As Boolean
    Dim oOSV As OSVERSIONINFO
    oOSV.OSVSize = Len(oOSV)

        '#### If the API returned a valid value
    If (GetVersionEx(oOSV) = 1) Then
        isWinXP = (oOSV.PlatformID = VER_PLATFORM_WIN32_NT And oOSV.dwVerMajor = 5 And oOSV.dwVerMinor = 1)
    End If
End Function


'#########################################################
'# Returns true if the OS is Win2k
'#########################################################
Public Function isWin2k() As Boolean
    Dim oOSV As OSVERSIONINFO
    oOSV.OSVSize = Len(oOSV)

        '#### If the API returned a valid value
    If (GetVersionEx(oOSV) = 1) Then
        isWin2k = (oOSV.PlatformID = VER_PLATFORM_WIN32_NT And oOSV.dwVerMajor = 5 And oOSV.dwVerMinor = 0)
    End If
End Function


'#########################################################
'# Returns true if the OS is WinNT4
'#########################################################
Public Function isWinNT4() As Boolean
    Dim oOSV As OSVERSIONINFO
    oOSV.OSVSize = Len(oOSV)

        '#### If the API returned a valid value
    If (GetVersionEx(oOSV) = 1) Then
        isWinNT4 = (oOSV.PlatformID = VER_PLATFORM_WIN32_NT And oOSV.dwVerMajor = 4)
    End If
End Function


'#########################################################
'# Returns true if the OS is WinME
'#########################################################
Public Function isWinME() As Boolean
    Dim oOSV As OSVERSIONINFO
    oOSV.OSVSize = Len(oOSV)

        '#### If the API returned a valid value
    If (GetVersionEx(oOSV) = 1) Then
        isWinME = (oOSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS And oOSV.dwVerMajor = 4 And oOSV.dwVerMinor = 90)
    End If
End Function


'#########################################################
'# Returns true if the OS is Win98
'#########################################################
Public Function isWin98() As Boolean
    Dim oOSV As OSVERSIONINFO
    oOSV.OSVSize = Len(oOSV)

        '#### If the API returned a valid value
    If (GetVersionEx(oOSV) = 1) Then
         isWin98 = (oOSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And (oOSV.dwVerMajor > 4) Or (oOSV.dwVerMajor = 4 And oOSV.dwVerMinor > 0)
    End If
End Function


'#########################################################
'# Returns true if the OS is Win95
'#########################################################
Public Function isWin95() As Boolean
    Dim oOSV As OSVERSIONINFO
    oOSV.OSVSize = Len(oOSV)

        '#### If the API returned a valid value
    If (GetVersionEx(oOSV) = 1) Then
         isWin95 = (oOSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS And oOSV.dwVerMajor = 4 And oOSV.dwVerMinor = 0)
    End If
End Function
