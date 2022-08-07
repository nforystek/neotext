Attribute VB_Name = "modFileInfo"
#Const modFileInfo = -1
Option Explicit
'TOP DOWN

Option Private Module
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias _
   "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal _
   dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias _
   "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, _
   lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias _
   "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, _
   lplpBuffer As Any, puLen As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias _
   "GetSystemDirectoryA" (ByVal Path As String, ByVal cbBytes As _
   Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Dest As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" ( _
    ByVal lpString1 As String, ByVal lpString2 As Long) As Long
    
Public Type FILEPROPERTIE
    CompanyName As String
    FileDescription As String
    FileVersion As String
    InternalName As String
    LegalCopyright As String
    OrigionalFileName As String
    ProductName As String
    ProductVersion As String
End Type

Public Declare Sub CoFreeUnusedLibraries Lib "ole32" ()

Public Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" ( _
           ByVal lpExistingFileName As String, _
           ByVal lpNewFileName As String, _
           ByVal dwFlags As Long _
) As Long

Public Const MOVEFILE_DELAY_UNTIL_REBOOT = &H4
Private Const MOVEFILE_REPLACE_EXISTING = 1
Private Const MOVEFILE_COPY_ALLOWED = 2
Private Const MOVEFILE_WRITE_THROUGH = 8

Public Function MoveLibrary(ByVal Source As String, ByVal Dest As String, Optional ByVal regsvr As Boolean = False, Optional ByVal backup As String = "") As Boolean
    'moves source to dest using backup as temp, regsvr if is activex, and returns rebootflag
    
    If PathExists(Source, True) Then
        Dim sver As String
        
        'MAKE BACKUP
        If backup <> "" Then
            If PathExists(Dest, True) Then
                If Not PathExists(backup, True) Then
                    On Error Resume Next
                    FileCopy Dest, backup
                    If Err Then Err.Clear
                    On Error GoTo 0
                End If
            End If
        End If
        
        'GET SOURCE VERSION
        sver = GetFileVersion(Source)
            
        'TOUCH TO CHECK REBOOT/USE
        CoFreeUnusedLibraries
        Dim fnum As Long
        fnum = FreeFile
        On Error Resume Next
        Open Dest For Binary Access Read Write Lock Read Write As #fnum
        If Err Then
            Err.Clear
            Close #fnum
            MoveLibrary = True 'REBOOT NEED
        Else
            Close #fnum
        End If
        On Error GoTo 0

            If regsvr Then RunProcess "regsvr32", "/s """ & backup & """", , False
            
       MoveFileEx Dest, backup, MOVEFILE_COPY_ALLOWED
            MoveFileEx Source, vbNullString, MOVEFILE_REPLACE_EXISTING + MOVEFILE_COPY_ALLOWED
            MoveFileEx Dest, vbNullString, MOVEFILE_REPLACE_EXISTING + MOVEFILE_COPY_ALLOWED

        MoveFileEx Source, Dest, MOVEFILE_DELAY_UNTIL_REBOOT
            MoveFileEx vbNullString, Dest, MOVEFILE_REPLACE_EXISTING + MOVEFILE_COPY_ALLOWED
            MoveFileEx vbNullString, backup, MOVEFILE_REPLACE_EXISTING + MOVEFILE_COPY_ALLOWED
                        
            If regsvr Then RunProcess "regsvr32", "/s """ & Dest & """", , False
                    
            MoveFileEx vbNullString, Dest, MOVEFILE_WRITE_THROUGH
            MoveFileEx Source, vbNullString, MOVEFILE_WRITE_THROUGH
    
        MoveFileEx Source, Dest, MOVEFILE_REPLACE_EXISTING
        
            MoveFileEx Dest, vbNullString, MOVEFILE_WRITE_THROUGH
            MoveFileEx vbNullString, backup, MOVEFILE_WRITE_THROUGH
            
        MoveFileEx Dest, backup, MOVEFILE_COPY_ALLOWED

            If regsvr Then RunProcess "regsvr32", "/s /u """ & backup & """", , False
            
            CoFreeUnusedLibraries
        
        MoveFileEx backup, vbNullString, MOVEFILE_WRITE_THROUGH
            
        If PathExists(Source, True) And PathExists(backup, True) Then
            Kill Source
            MoveLibrary = False 'REEVAL NEED REBOOT
        End If

    Else
        Err.Raise 53
    End If
End Function

Public Function CheckVersionEqualOrGreater(ByVal VerCheck As String, ByVal VerAgainst As String) As Boolean
    Dim fvi As FILEPROPERTIE
    If PathExists(VerCheck, True) Then
        fvi = GetFileInfo(VerCheck)
    Else
        fvi.FileVersion = VerCheck
    End If
    
    Dim v1 As String
    Dim v2 As String
    Dim v3 As String
    Dim v4 As String
    fvi.FileVersion = Replace(RemoveNextArg(fvi.FileVersion, " "), "'", ".")
    v1 = RemoveNextArg(fvi.FileVersion, ".")
    v2 = RemoveNextArg(fvi.FileVersion, ".")
    v3 = RemoveNextArg(fvi.FileVersion, ".")
    v4 = RemoveNextArg(fvi.FileVersion, ".")
    If Not IsNumeric(v1) Then v1 = "0"
    If Not IsNumeric(v2) Then v2 = "0"
    If Not IsNumeric(v3) Then v3 = "0"
    If Not IsNumeric(v4) Then v4 = "0"


    If PathExists(VerAgainst, True) Then
        fvi = GetFileInfo(VerAgainst)
    Else
        fvi.FileVersion = VerAgainst
    End If
    
    Dim u1 As String
    Dim u2 As String
    Dim u3 As String
    Dim u4 As String
    fvi.FileVersion = Replace(RemoveNextArg(fvi.FileVersion, " "), "'", ".")
    u1 = RemoveNextArg(fvi.FileVersion, ".")
    u2 = RemoveNextArg(fvi.FileVersion, ".")
    u3 = RemoveNextArg(fvi.FileVersion, ".")
    u4 = RemoveNextArg(fvi.FileVersion, ".")
    If Not IsNumeric(u1) Then u1 = "0"
    If Not IsNumeric(u2) Then u2 = "0"
    If Not IsNumeric(u3) Then u3 = "0"
    If Not IsNumeric(u4) Then u4 = "0"
    
    If v1 = u1 And v2 = u2 And v3 = u3 And v4 = u4 Then
        CheckVersionEqualOrGreater = True
    ElseIf v1 > u1 Then
        CheckVersionEqualOrGreater = True
    ElseIf v1 = u1 Then
        If v2 > u2 Then
            CheckVersionEqualOrGreater = True
        ElseIf v2 = u2 Then
            If v3 > u3 Then
                CheckVersionEqualOrGreater = True
            ElseIf v3 = u3 Then
                If v4 > u4 Then
                    CheckVersionEqualOrGreater = True
                End If
            End If
        End If
    End If
End Function


Public Function CheckVersionEqual(ByVal VerCheck As String, ByVal VerAgainst As String) As Boolean
    Dim fvi As FILEPROPERTIE
    If PathExists(VerCheck, True) Then
        fvi = GetFileInfo(VerCheck)
    Else
        fvi.FileVersion = VerCheck
    End If
    
    Dim v1 As String
    Dim v2 As String
    Dim v3 As String
    Dim v4 As String
    fvi.FileVersion = Replace(RemoveNextArg(fvi.FileVersion, " "), "'", ".")
    v1 = RemoveNextArg(fvi.FileVersion, ".")
    v2 = RemoveNextArg(fvi.FileVersion, ".")
    v3 = RemoveNextArg(fvi.FileVersion, ".")
    v4 = RemoveNextArg(fvi.FileVersion, ".")
    If Not IsNumeric(v1) Then v1 = "0"
    If Not IsNumeric(v2) Then v2 = "0"
    If Not IsNumeric(v3) Then v3 = "0"
    If Not IsNumeric(v4) Then v4 = "0"


    If PathExists(VerAgainst, True) Then
        fvi = GetFileInfo(VerAgainst)
    Else
        fvi.FileVersion = VerAgainst
    End If
    
    Dim u1 As String
    Dim u2 As String
    Dim u3 As String
    Dim u4 As String
    fvi.FileVersion = Replace(RemoveNextArg(fvi.FileVersion, " "), "'", ".")
    u1 = RemoveNextArg(fvi.FileVersion, ".")
    u2 = RemoveNextArg(fvi.FileVersion, ".")
    u3 = RemoveNextArg(fvi.FileVersion, ".")
    u4 = RemoveNextArg(fvi.FileVersion, ".")
    If Not IsNumeric(u1) Then u1 = "0"
    If Not IsNumeric(u2) Then u2 = "0"
    If Not IsNumeric(u3) Then u3 = "0"
    If Not IsNumeric(u4) Then u4 = "0"
    
    If v1 = u1 And v2 = u2 And v3 = u3 And v4 = u4 Then
        CheckVersionEqual = True
    End If
End Function

Public Function CheckVersionGreater(ByVal VerCheck As String, ByVal VerAgainst As String) As Boolean
    Dim fvi As FILEPROPERTIE
    If PathExists(VerCheck, True) Then
        fvi = GetFileInfo(VerCheck)
    Else
        fvi.FileVersion = VerCheck
    End If
    
    Dim v1 As String
    Dim v2 As String
    Dim v3 As String
    Dim v4 As String
    fvi.FileVersion = Replace(RemoveNextArg(fvi.FileVersion, " "), "'", ".")
    v1 = RemoveNextArg(fvi.FileVersion, ".")
    v2 = RemoveNextArg(fvi.FileVersion, ".")
    v3 = RemoveNextArg(fvi.FileVersion, ".")
    v4 = RemoveNextArg(fvi.FileVersion, ".")
    If Not IsNumeric(v1) Then v1 = "0"
    If Not IsNumeric(v2) Then v2 = "0"
    If Not IsNumeric(v3) Then v3 = "0"
    If Not IsNumeric(v4) Then v4 = "0"

    If PathExists(VerAgainst, True) Then
        fvi = GetFileInfo(VerAgainst)
    Else
        fvi.FileVersion = VerAgainst
    End If
    
    Dim u1 As String
    Dim u2 As String
    Dim u3 As String
    Dim u4 As String
    fvi.FileVersion = Replace(RemoveNextArg(fvi.FileVersion, " "), "'", ".")
    u1 = RemoveNextArg(fvi.FileVersion, ".")
    u2 = RemoveNextArg(fvi.FileVersion, ".")
    u3 = RemoveNextArg(fvi.FileVersion, ".")
    u4 = RemoveNextArg(fvi.FileVersion, ".")
    If Not IsNumeric(u1) Then u1 = "0"
    If Not IsNumeric(u2) Then u2 = "0"
    If Not IsNumeric(u3) Then u3 = "0"
    If Not IsNumeric(u4) Then u4 = "0"
    
    If v1 > u1 Then
        CheckVersionGreater = True
    ElseIf v1 = u1 Then
        If v2 > u2 Then
            CheckVersionGreater = True
        ElseIf v2 = u2 Then
            If v3 > u3 Then
                CheckVersionGreater = True
            ElseIf v3 = u3 Then
                If v4 > u4 Then
                    CheckVersionGreater = True
                End If
            End If
        End If
    End If
End Function
Public Function GetFileVersion(ByVal FilePath As String) As String
    Dim fvi As FILEPROPERTIE
    fvi = GetFileInfo(FilePath)
    GetFileVersion = Replace(RemoveNextArg(fvi.FileVersion, " "), "'", ".")
End Function
Public Function GetFileInfo(ByVal PathWithFilename As String) As FILEPROPERTIE
    Dim lngBufferlen As Long
    Dim lngDummy As Long
    Dim lngRc As Long
    Dim lngVerPointer As Long
    Dim lngHexNumber As Long
    Dim bytBuffer() As Byte
    Dim bytBuff(255) As Byte
    Dim strBuffer As String
    Dim strLangCharset As String
    Dim strVersionInfo(7) As String
    Dim strTemp As String
    Dim intTemp As Integer
           
    ' size
    lngBufferlen = GetFileVersionInfoSize(PathWithFilename, lngDummy)
    If lngBufferlen > 0 Then
       ReDim bytBuffer(lngBufferlen)
       lngRc = GetFileVersionInfo(PathWithFilename, 0&, lngBufferlen, bytBuffer(0))
       If lngRc <> 0 Then
          lngRc = VerQueryValue(bytBuffer(0), "\VarFileInfo\Translation", _
                   lngVerPointer, lngBufferlen)
          If lngRc <> 0 Then
             'lngVerPointer is a pointer to four 4 bytes of Hex number,
             'first two bytes are language id, and last two bytes are code
             'page. However, strLangCharset needs a  string of
             '4 hex digits, the first two characters correspond to the
             'language id and last two the last two character correspond
             'to the code page id.
             MoveMemory bytBuff(0), lngVerPointer, lngBufferlen
             lngHexNumber = bytBuff(2) + bytBuff(3) * &H100 + _
                    bytBuff(0) * &H10000 + bytBuff(1) * &H1000000
             strLangCharset = Hex(lngHexNumber)
             'now we change the order of the language id and code page
             'and convert it into a string representation.
             'For example, it may look like 040904E4
             'Or to pull it all apart:
             '04------        = SUBLANG_ENGLISH_USA
             '--09----        = LANG_ENGLISH
             ' ----04E4 = 1252 = Codepage for Windows:Multilingual
             Do While Len(strLangCharset) < 8
                 strLangCharset = "0" & strLangCharset
             Loop
             ' assign propertienames
             strVersionInfo(0) = "CompanyName"
             strVersionInfo(1) = "FileDescription"
             strVersionInfo(2) = "FileVersion"
             strVersionInfo(3) = "InternalName"
             strVersionInfo(4) = "LegalCopyright"
             strVersionInfo(5) = "OriginalFileName"
             strVersionInfo(6) = "ProductName"
             strVersionInfo(7) = "ProductVersion"
             ' loop and get fileproperties
             For intTemp = 0 To 7
                strBuffer = String$(255, 0)
                strTemp = "\StringFileInfo\" & strLangCharset _
                   & "\" & strVersionInfo(intTemp)
                lngRc = VerQueryValue(bytBuffer(0), strTemp, _
                      lngVerPointer, lngBufferlen)
                If lngRc <> 0 Then
                   ' get and format data
                   lstrcpy strBuffer, lngVerPointer
                   strBuffer = Mid$(strBuffer, 1, InStr(strBuffer, Chr(0)) - 1)
                   strVersionInfo(intTemp) = strBuffer
                 Else
                   ' property not found
                   strVersionInfo(intTemp) = "?"
                End If
             Next intTemp
          End If
       End If
    End If
    ' assign array to user-defined-type
    GetFileInfo.CompanyName = strVersionInfo(0)
    GetFileInfo.FileDescription = strVersionInfo(1)
    GetFileInfo.FileVersion = strVersionInfo(2)
    GetFileInfo.InternalName = strVersionInfo(3)
    GetFileInfo.LegalCopyright = strVersionInfo(4)
    GetFileInfo.OrigionalFileName = strVersionInfo(5)
    GetFileInfo.ProductName = strVersionInfo(6)
    GetFileInfo.ProductVersion = strVersionInfo(7)
End Function



