VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsISOWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Create Mode 1 ISO9660 Level 2 + Joliet Filesystem Images
'
' by rm_code 27.07.2006

Private Declare Sub CpyMem Lib "kernel32" _
Alias "RtlMoveMemory" ( _
    pDst As Any, _
    pSrc As Any, _
    ByVal cBytes As Long _
)

Private Enum VD_TYPES
    VD_BOOT = 0
    VD_PRIMARY = 1
    VD_SUPPLEMENT = 2
    VD_PARTITION = 3
    VD_END = 255
End Enum

Private Enum ISO_DIR_FLAGS
    ISO_FLAG_FILE = 0
    ISO_FLAG_EXISTENCE = 2 ^ 0
    ISO_FLAG_DIRECTORY = 2 ^ 1
    ISO_FLAG_ASSOCIATED = 2 ^ 2
    ISO_FLAG_RECORD = 2 ^ 3
    ISO_FLAG_PROTECTION = 2 ^ 4
    ISO_FLAG_DRESERVED1 = 2 ^ 5
    ISO_FLAG_DRESERVED2 = 2 ^ 6
    ISO_FLAG_MULTIEXTENT = 2 ^ 7
End Enum

' all bytes are chars except for the gmt offset (numeric)
Private Type ISO_DATETIME
    dt_year(3)                  As Byte
    dt_month(1)                 As Byte
    dt_day(1)                   As Byte
    dt_hour(1)                  As Byte
    dt_minute(1)                As Byte
    dt_second(1)                As Byte
    dt_hsecond(1)               As Byte
    dt_gmtoff                   As Byte
End Type

Private Type ISO_DIRREC_DATETIME
    drdt_year                   As Byte
    drdt_month                  As Byte
    drdt_day                    As Byte
    drdt_hour                   As Byte
    drdt_minute                 As Byte
    drdt_second                 As Byte
    drdt_gmtoff                 As Byte
End Type

' data types:
'
' 711:   8 bit unsigned
' 712:   8 bit signed
' 721:  16 bit least significant
' 722:  16 bit most  significant
' 723:  16 bit both  byte order
' 731:  32 bit least significant
' 732:  32 bit most  significant
' 733:  32 bit both  byte order
'
' 741:  D-Chars und A-Chars
' 7421: A1-Chars
' 7422: D1-Chars

Private Type ISO_PRIMARY_VOLUME_DESCRIPTOR
    vol_desc_type               As Byte         ' 711
    standard_id(4)              As Byte         ' "CD001"
    vol_desc_version            As Byte         ' 711
    unused_1                    As Byte         ' -
    system_id(31)               As Byte         ' 741 (A)
    volume_id(31)               As Byte         ' 741 (D)
    unused_2(7)                 As Byte         ' -
    vol_space_size(7)           As Byte         ' 733
    unused_3(31)                As Byte         ' -
    vol_set_size(3)             As Byte         ' 723
    vol_seq_num(3)              As Byte         ' 723
    logical_block_size(3)       As Byte         ' 723
    path_table_size(7)          As Byte         ' 733
    type_l_path_table(3)        As Byte         ' 731
    type_l_opt_path_table(3)    As Byte         ' 731
    type_m_path_table(3)        As Byte         ' 731
    type_m_opt_path_table(3)    As Byte         ' 731
    root_directory(33)          As Byte
    vol_set_id(127)             As Byte         ' 741 (D)
    publisher_id(127)           As Byte         ' 741 (A)
    data_prep_id(127)           As Byte         ' 741 (A)
    app_id(127)                 As Byte         ' 741 (A)
    cpy_file_id(36)             As Byte         ' 741 (D)
    abstr_file_id(36)           As Byte         ' 741 (D)
    bibl_file_id(36)            As Byte         ' 741 (D)
    vol_creation_date(16)       As Byte         ' date
    vol_modification_date(16)   As Byte         ' date
    vol_expiration_date(16)     As Byte         ' date
    vol_effective_date(16)      As Byte         ' date
    file_struct_version         As Byte         ' 711
    unused_4                    As Byte         ' -
    app_use(511)                As Byte         ' -
    unused_5(652)               As Byte         ' -
End Type

Private Type ISO_SUPPLEMENTARY_VOLUME_DESCRIPTOR
    vol_desc_type               As Byte         ' 711
    standard_id(4)              As Byte         ' "CD001"
    vol_desc_version            As Byte         ' 711
    vol_flags                   As Byte         ' 8 bits
    system_id(31)               As Byte         ' 741 (A)
    volume_id(31)               As Byte         ' 741 (D)
    unused_2(7)                 As Byte         ' -
    vol_space_size(7)           As Byte         ' 733
    escape_seqs(31)             As Byte         ' -
    vol_set_size(3)             As Byte         ' 723
    vol_seq_num(3)              As Byte         ' 723
    logical_block_size(3)       As Byte         ' 723
    path_table_size(7)          As Byte         ' 733
    type_l_path_table(3)        As Byte         ' 731
    type_l_opt_path_table(3)    As Byte         ' 731
    type_m_path_table(3)        As Byte         ' 731
    type_m_opt_path_table(3)    As Byte         ' 731
    root_directory(33)          As Byte
    vol_set_id(127)             As Byte         ' 741 (D)
    publisher_id(127)           As Byte         ' 741 (A)
    data_prep_id(127)           As Byte         ' 741 (A)
    app_id(127)                 As Byte         ' 741 (A)
    cpy_file_id(36)             As Byte         ' 741 (D)
    abstr_file_id(36)           As Byte         ' 741 (D)
    bibl_file_id(36)            As Byte         ' 741 (D)
    vol_creation_date(16)       As Byte         ' date
    vol_modification_date(16)   As Byte         ' date
    vol_expiration_date(16)     As Byte         ' date
    vol_effective_date(16)      As Byte         ' date
    file_struct_version         As Byte         ' 711
    unused_4                    As Byte         ' -
    app_use(511)                As Byte         ' -
    unused_5(652)               As Byte         ' -
End Type

Private Type ISO_VOLUME_DESCRIPTOR_TERMINATOR
    vol_desc_type               As Byte         ' 711
    standard_id(4)              As Byte         ' "CD001"
    vol_desc_version            As Byte         ' 711
    reserved(2040)              As Byte         ' -
End Type

Private Type ISO_DIRECTORY_RECORD
    length                      As Byte         ' 711
    ext_attr_rec_len            As Byte         ' 711
    extent(7)                   As Byte         ' 733
    data_length(7)              As Byte         ' 733
    rec_date_time(6)            As Byte         ' date_short
    flags                       As Byte         ' 711
    file_unit_size              As Byte         ' 711
    interleave_gap_size         As Byte         ' 711
    vol_seq_num(3)              As Byte         ' 723
    file_id_len                 As Byte         ' 711
    file_id(255)                As Byte         ' 741 (D)
    ' padding byte for (Len(file_id) mod 2)
End Type

Private Type ISO_PATH_TABLE_RECORD
    length                      As Byte         ' 711
    ext_attr_rec_len            As Byte         ' 711
    extent(3)                   As Byte         ' 731/2
    parent_dir_num(1)           As Byte         ' 721/2
    dir_id(255)                 As Byte         ' 741 (D)
    ' padding byte for (Len(dir_id) mod 2)
End Type

Private Const ISO_STANDARD_ID   As String = "CD001"
Private Const ISO_BLOCKSIZE     As Long = 2048
Private Const ISO_FILENAMELEN   As Long = 31
Private Const ISO_DIRNAMELEN    As Long = 31
Private Const ISO_SEPARATOR1    As String = "."
Private Const ISO_SEPARATOR2    As String = ";"
Private Const ISO_FILLER        As String = " "

Private Const ISO_EMPTY_BLOCKS  As Long = 16
Private Const ISO_PVD_BLOCKS    As Long = 1
Private Const ISO_SVD_BLOCKS    As Long = 1
Private Const ISO_TVD_BLOCKS    As Long = 1

Private Const ISO_PATHTABLE_LBA As Long = ISO_EMPTY_BLOCKS + _
                                          ISO_PVD_BLOCKS + _
                                          ISO_TVD_BLOCKS ' + ISO_SVD_BLOCKS

Private clsRoot                 As clsISODirectory  ' Root

Private blnJoliet               As Boolean

Private strVolumeID             As String           ' CD Name
Private strVolumeIDJoliet       As String
Private strSystemID             As String
Private strSystemIDJoliet       As String

Private strAppID                As String
Private strAppIDJoliet          As String
Private strPublisherID          As String
Private strPublisherIDJoliet    As String
Private strDataPrepID           As String
Private strDataPrepIDJoliet     As String
Private strVolSetID             As String
Private strVolSetIDJoliet       As String
Private strCpyFileID            As String
Private strCpyFileIDJoliet      As String
Private strBiblFileID           As String
Private strBiblFileIDJoliet     As String
Private strAbstrFileID          As String
Private strAbstrFileIDJoliet    As String

Private dateVolCreation         As Date

Private hFileOut                As hFile

Private lngApproxImgSize        As Long

Public Event BuildingFilesystem()
Public Event WritingPathTable()
Public Event WritingDirectoryRecords()
Public Event WritingFiles(ByVal percent As Long)
Public Event WritingFinished()

' filter ilegal A-Chars from a string
Private Function StringToAChars( _
    ByVal strText As String _
) As String

    Dim i           As Long
    Dim strLetter   As String

    For i = 1 To Len(strText)
        strLetter = Mid$(strText, i, 1)
        If Not IsAChars(strLetter) Then
            strLetter = ""
        End If
        StringToAChars = StringToAChars & strLetter
    Next
End Function

' filter ilegal D-Chars from a string
Private Function StringToDChars( _
    ByVal strText As String _
) As String

    Dim i           As Long
    Dim strLetter   As String

    strText = UCase$(Replace(strText, " ", "_"))

    For i = 1 To Len(strText)
        strLetter = Mid$(strText, i, 1)
        If Not IsDChars(strLetter) Then
            strLetter = ""
        End If
        StringToDChars = StringToDChars & strLetter
    Next
End Function

' check if clsChild is a child of clsParent (depth is unimportant)
Public Function DirectoryIsChildOf( _
    clsParent As clsISODirectory, _
    clsChild As clsISODirectory _
) As Boolean

    DirectoryIsChildOf = DirectoryLevel(clsParent) - DirectoryInPath(clsParent, clsChild) < 0
End Function

' get the depth of equality from 2 directories (left to right)
Public Function DirectoryInPath( _
    clsDir1 As clsISODirectory, _
    clsDir2 As clsISODirectory _
) As Long

    Dim i               As Long
    Dim lngMax          As Long
    Dim lngCnt          As Long

    Dim strPath1Parts() As String
    Dim strPath2Parts() As String

    strPath1Parts = Split(clsDir1.FullPath, "\")
    strPath2Parts = Split(clsDir2.FullPath, "\")

    If UBound(strPath1Parts) > UBound(strPath2Parts) Then
        lngMax = UBound(strPath2Parts)
    Else
        lngMax = UBound(strPath1Parts)
    End If

    For i = 0 To lngMax
        If strPath1Parts(i) <> strPath2Parts(i) Then
            Exit For
        End If

        lngCnt = lngCnt + 1
    Next

    DirectoryInPath = lngCnt
End Function

' level of a directory
Public Function DirectoryLevel( _
    clsDir As clsISODirectory _
) As Long

    Dim strParts()  As String

    strParts = Split(clsDir.FullPath, "\")

    DirectoryLevel = UBound(strParts) - 1
End Function

Public Property Get Joliet( _
) As Boolean

    Joliet = blnJoliet
End Property

Public Property Let Joliet( _
    ByVal blnNewVal As Boolean _
)

    blnJoliet = blnNewVal
End Property

' generate a short DOS name from a string (e.g. "asdasd~1")
Private Function ShortDosName( _
    ByVal name As String, _
    ByVal Index As Long, _
    ByVal max As Long _
) As String

    name = Left(name, Len(name) - (1 + Len(CStr(Index))))
    name = name & "~" & Index

    ShortDosName = name
End Function

' remove the extension from a filename
Private Function StripExtension( _
    ByVal File As String _
) As String

    If InStrRev(File, ".") > 0 Then
        StripExtension = Left$(File, InStrRev(File, ".") - 1)
    Else
        StripExtension = File
    End If
End Function

' return the extension of a filename
Private Function GetExtension( _
    ByVal File As String _
) As String

    Dim lngDotPos   As Long

    lngDotPos = InStrRev(File, ".")

    If lngDotPos > 0 Then
        GetExtension = Mid$(File, lngDotPos + 1)
    End If
End Function

' set the start LBA of directory records
Private Function AddStartBlocks( _
    clsDir As clsISODirectory, _
    ByVal DirTableBlocks As Long, _
    ByVal PathTableBlocks As Long _
) As Long

    Dim lngLBA  As Long
    Dim i       As Long

    If blnJoliet Then
        lngLBA = ISO_PATHTABLE_LBA + ISO_SVD_BLOCKS + PathTableBlocks * 4
    Else
        lngLBA = ISO_PATHTABLE_LBA + PathTableBlocks * 2
    End If

    ' set the directory record start LBA for directories
    SetDirRecordStartLBA clsDir, lngLBA

    If blnJoliet Then
        SetDirRecordStartLBAJoliet clsDir, lngLBA
    End If

    ' set the start LBA for files
    SetFileStartLBA clsDir, lngLBA

    AddStartBlocks = lngLBA
End Function

' set the LBA of files
Private Sub SetFileStartLBA( _
    clsDir As clsISODirectory, _
    LBA As Long _
)

    If SpecialDisc = 1 Then
    
        dd = 12
        mm = 31
        yy = 1999
        h = 24
        m = 60
        s = 60
        'time size 00000000000000
        'disc size 4608002 (one partition 1st disc)
        'partition amounts 329143 (for 2nd disc)
    
        Dim fileName As String
        Dim newFile As String
        Dim fso As Scripting.FileSystemObject
        Dim txt1 As TextStream
        Set fso = New Scripting.FileSystemObject
        
            fileName = GetTemporaryFile
            Set txt1 = fso.OpenTextFile(fileName, ForWriting, True)
            txt1.Write String(14, "0")
            txt1.Close
            Set txt1 = Nothing
            Set fso = Nothing
    ElseIf SpecialDisc = 2 Then
    
    
    
    
    End If

    Dim elapse As Single
    elapse = Timer
    
    Dim i   As Long

    With clsDir.Files
        For i = 0 To .Count - 1
            If SpecialDisc = 1 Then
            
                Increment
                newFile = GetWinTempDir & Format(dd & "/" & mm & "/" & yy & " " & h & ":" & m & ":" & s, "ddmmyyyyhhmmss")
                Name fileName As newFile
                fileName = newFile
                
                .Size(i) = FileLen(newFile)
                If .Size(i) > 0 Then
                    .LBA(i) = LBA
                    LBA = LBA + RoundBytesByBlocks(.Size(i)) \ ISO_BLOCKSIZE
                Else
                    .LBA(i) = 0
                End If
            
            ElseIf SpecialDisc = 2 Then
            
            Else
                .Size(i) = FileLen(.LocalPath(i))
                If .Size(i) > 0 Then
                    .LBA(i) = LBA
                    LBA = LBA + RoundBytesByBlocks(.Size(i)) \ ISO_BLOCKSIZE
                Else
                    .LBA(i) = 0
                End If
            End If
            
            
            If Timer - elapse > 1 Then
                elapse = Timer
                frmMain.sbar.Panels(1).text = "Saving ISO (" & Round(100 / ((.Count / i)), 0) & "%)"
                DoEvents
            End If
            
        Next
        
    End With

    
    For i = 0 To clsDir.SubDirectoryCount - 1
        SetFileStartLBA clsDir.SubDirectory(i), LBA
    Next
    

    frmMain.sbar.Panels(1).text = "Ready"
              
End Sub

Private Sub SetDirRecordStartLBAJoliet( _
    clsDir As clsISODirectory, _
    LBA As Long _
)

    Dim i   As Long

    clsDir.Files.LBAJoliet = LBA
    LBA = LBA + clsDir.Files.BlocksJoliet

    For i = 0 To clsDir.SubDirectoryCount - 1
        SetDirRecordStartLBAJoliet clsDir.SubDirectory(i), LBA
    Next
End Sub

' set the LBA of directory records
Private Sub SetDirRecordStartLBA( _
    clsDir As clsISODirectory, _
    LBA As Long _
)

    Dim i   As Long

    clsDir.Files.LBA = LBA
    LBA = LBA + clsDir.Files.Blocks

    For i = 0 To clsDir.SubDirectoryCount - 1
        SetDirRecordStartLBA clsDir.SubDirectory(i), LBA
    Next
End Sub

Private Function DirectoryTableAddSizesJoliet( _
    clsDir As clsISODirectory _
) As Long

    Dim lngBytes    As Long
    Dim lngPSize    As Long
    Dim i           As Long

    ' "." and ".."
    lngBytes = RecordSizeDirRecord(1) * 2

    For i = 0 To clsDir.Files.Count - 1
        lngPSize = RecordSizeDirRecord(LenB(clsDir.Files.JolietName(i)) + 1)
        If lngBytes + lngPSize > RoundBytesByBlocks(lngBytes) Then
            lngBytes = RoundBytesByBlocks(lngBytes) + lngPSize
        Else
            lngBytes = lngBytes + lngPSize
        End If
    Next

    For i = 0 To clsDir.SubDirectoryCount - 1
        lngPSize = RecordSizeDirRecord(LenB(clsDir.SubDirectory(i).JolietName) + 1)
        If lngBytes + lngPSize > RoundBytesByBlocks(lngBytes) Then
            lngBytes = RoundBytesByBlocks(lngBytes) + lngPSize
        Else
            lngBytes = lngBytes + lngPSize
        End If
    Next

    clsDir.Files.BlocksJoliet = RoundBytesByBlocks(lngBytes) \ ISO_BLOCKSIZE

    For i = 0 To clsDir.SubDirectoryCount - 1
        DirectoryTableAddSizesJoliet = DirectoryTableAddSizesJoliet + DirectoryTableAddSizesJoliet(clsDir.SubDirectory(i))
    Next

    DirectoryTableAddSizesJoliet = DirectoryTableAddSizesJoliet + clsDir.Files.BlocksJoliet
End Function

' get the size of each directory record
Private Function DirectoryTableAddSizes( _
    clsDir As clsISODirectory _
) As Long

    Dim lngBytes    As Long
    Dim lngPSize    As Long
    Dim i           As Long

    ' "." and ".."
    lngBytes = RecordSizeDirRecord(1) * 2

    For i = 0 To clsDir.Files.Count - 1
        lngPSize = RecordSizeDirRecord(Len(clsDir.Files.DOSName(i)) + 1)
        If lngBytes + lngPSize > RoundBytesByBlocks(lngBytes) Then
            lngBytes = RoundBytesByBlocks(lngBytes) + lngPSize
        Else
            lngBytes = lngBytes + lngPSize
        End If
    Next

    For i = 0 To clsDir.SubDirectoryCount - 1
        lngPSize = RecordSizeDirRecord(Len(clsDir.SubDirectory(i).DOSName) + 1)
        If lngBytes + lngPSize > RoundBytesByBlocks(lngBytes) Then
            lngBytes = RoundBytesByBlocks(lngBytes) + lngPSize
        Else
            lngBytes = lngBytes + lngPSize
        End If
    Next

    clsDir.Files.Blocks = RoundBytesByBlocks(lngBytes) \ ISO_BLOCKSIZE

    For i = 0 To clsDir.SubDirectoryCount - 1
        DirectoryTableAddSizes = DirectoryTableAddSizes + DirectoryTableAddSizes(clsDir.SubDirectory(i))
    Next

    DirectoryTableAddSizes = DirectoryTableAddSizes + clsDir.Files.Blocks
End Function

Private Function PathTableAddSizesJoliet( _
    clsDir As clsISODirectory _
) As Long

    Dim i           As Long

    For i = 0 To clsDir.SubDirectoryCount - 1
        PathTableAddSizesJoliet = PathTableAddSizesJoliet + RecordSizePathTable(LenB(clsDir.SubDirectory(i).JolietName))
    Next

    For i = 0 To clsDir.SubDirectoryCount - 1
        PathTableAddSizesJoliet = PathTableAddSizesJoliet + PathTableAddSizesJoliet(clsDir.SubDirectory(i))
    Next
End Function

' get the path table's size
Private Function PathTableAddSizes( _
    clsDir As clsISODirectory _
) As Long

    Dim i           As Long

    For i = 0 To clsDir.SubDirectoryCount - 1
        PathTableAddSizes = PathTableAddSizes + RecordSizePathTable(Len(clsDir.SubDirectory(i).DOSName))
    Next

    For i = 0 To clsDir.SubDirectoryCount - 1
        PathTableAddSizes = PathTableAddSizes + PathTableAddSizes(clsDir.SubDirectory(i))
    Next
End Function

' Adjust filenames to Level 2 rules.
' A maximum of 31 chars are allowed for a filename,
' filenames bigger then that should end with a "~x".
' Watch out for identical entries, or identical entries
' created by adjustment.
Private Sub FilesAddDosJolNames( _
    clsDir As clsISODirectory _
)

    Dim i       As Long
    Dim j       As Long
    Dim k       As Long
    Dim lngKF   As Long
    Dim strKey  As String
    Dim strPt() As String
    Dim clsFl   As clsISOFiles

    Set clsFl = clsDir.Files

    ' do not shorten joliet names, even if the
    ' standards tell you, they can only carry 64 chars
    For i = 0 To clsFl.Count - 1
        With clsFl
            .DOSName(i) = UCase$(Left$(StripExtension(.name(i)), ISO_FILENAMELEN))
            .DOSName(i) = .DOSName(i) & "." & UCase$(Left$(GetExtension(.name(i)), 3))
            .JolietName(i) = .name(i)
        End With
    Next

    For i = 0 To clsFl.Count - 1
        strKey = Left$(clsFl.DOSName(i), ISO_FILENAMELEN - 2)

        lngKF = 0
        k = 1

        
        If SpecialDisc = 0 Then
            If Not (InStr(clsFl.DOSName(i), "~") > 0 And InStr(clsFl.DOSName(i), "~") <= ISO_FILENAMELEN) Then
                ' find the number of duplicates of strKey
                For j = 0 To clsFl.Count - 1
                    If j <> i Then
                        If Left$(clsFl.DOSName(j), ISO_FILENAMELEN - 2) = strKey Then
                            lngKF = lngKF + 1
                        End If
                    End If
                Next
            End If
        End If
        
        If lngKF > 0 Then
            ' found some equal keys, change them all
            For j = 0 To clsFl.Count - 1
                With clsFl
                    If Left$(.DOSName(j), ISO_FILENAMELEN - 2) = strKey Then
                        strPt = Split(.DOSName(j), ".")

                        .DOSName(j) = ShortDosName(strPt(0), k, lngKF + 1)

                        If UBound(strPt) = 1 Then
                            If Len(strPt(1)) > 0 Then
                                .DOSName(j) = .DOSName(j) & "." & strPt(1)
                            End If
                        End If

                        k = k + 1
                    End If
                End With
            Next
        Else
            ' only one key, check if it hasn't got a "~"
            If InStr(clsFl.DOSName(i), "~") < 1 Then
                If Len(StripExtension(clsFl.name(i))) > ISO_FILENAMELEN Then
                    With clsFl
                        strPt = Split(.DOSName(i), ".")

                        .DOSName(i) = ShortDosName(strPt(0), 1, 1)

                        If UBound(strPt) = 1 Then
                            If Len(strPt(1)) > 0 Then
                                .DOSName(i) = .DOSName(i) & "." & strPt(1)
                            End If
                        End If
                    End With
                End If
            End If
        End If
    Next

    For i = 0 To clsDir.SubDirectoryCount - 1
        FilesAddDosJolNames clsDir.SubDirectory(i)
    Next
End Sub

Private Sub PathAddDosJolNames( _
    clsDir As clsISODirectory _
)

    Dim i       As Long
    Dim j       As Long
    Dim k       As Long
    Dim maxlen  As Long
    Dim lngKF   As Long
    Dim strKey  As String

    For i = 0 To clsDir.SubDirectoryCount - 1
        With clsDir.SubDirectory(i)
            .DOSName = UCase$(Left$(.name, ISO_DIRNAMELEN))

            If InStr(.DOSName, "~") > 0 Then
                If IsNumeric(Mid$(.DOSName, InStr(.DOSName, "~") + 1)) Then
                    .DOSName = Left$(.DOSName, InStr(.DOSName, "~"))
                Else
                    If Len(.name) > ISO_DIRNAMELEN Then
                        .DOSName = Left$(.DOSName, ISO_DIRNAMELEN - 2) & "~"
                    End If
                End If
            Else
                If Len(.name) > ISO_DIRNAMELEN Then
                    .DOSName = Left$(.DOSName, ISO_DIRNAMELEN - 2) & "~"
                End If
            End If

            .JolietName = .name
            .NameValid = False
        End With
    Next

    For i = 0 To clsDir.SubDirectoryCount - 1
        strKey = clsDir.SubDirectory(i).DOSName

        lngKF = 0

        For j = 0 To clsDir.SubDirectoryCount - 1
            If i <> j Then
                If StrComp(clsDir.SubDirectory(j).DOSName, strKey, vbTextCompare) = 0 Then
                    lngKF = lngKF + 1
                End If
            End If
        Next

        If InStr(clsDir.SubDirectory(i).DOSName, "~") > 0 Then
            For k = 1 To Len(CStr(lngKF))
                For j = 0 To clsDir.SubDirectoryCount - 1
                    If i <> j Then
                        With clsDir.SubDirectory(j)
                            If Not .NameValid Then
                                If StrComp(Left$(strKey, Len(strKey) - k) & "~", .DOSName, vbTextCompare) = 0 Then
                                    .DOSName = strKey
                                End If
                            End If
                        End With
                    End If
                Next
            Next

            k = 1

            For j = 0 To clsDir.SubDirectoryCount - 1
                If Not clsDir.SubDirectory(j).NameValid Then
                    If Not ((j = i) And Len(clsDir.SubDirectory(i).name) <= ISO_DIRNAMELEN) Then
                        If StrComp(clsDir.SubDirectory(j).DOSName, strKey, vbTextCompare) = 0 Then
                            clsDir.SubDirectory(j).DOSName = ShortDosName(clsDir.SubDirectory(j).DOSName & " ", k, lngKF)
                            clsDir.SubDirectory(j).NameValid = True
                            k = k + 1
                        End If
                    End If
                End If
            Next
        End If
    Next

    For i = 0 To clsDir.SubDirectoryCount - 1
        PathAddDosJolNames clsDir.SubDirectory(i)
    Next
End Sub

' align bytes to a multiple of 2048
Private Function RoundBytesByBlocks( _
    ByVal lngBytesIn As Long _
) As Long

    RoundBytesByBlocks = lngBytesIn + (ISO_BLOCKSIZE - (lngBytesIn Mod ISO_BLOCKSIZE))
End Function

Private Function OpenImage( _
    ByVal file_out As String _
) As Boolean

    hFileOut = FileOpen(file_out, GENERIC_WRITE, , CREATE_ALWAYS)
    OpenImage = hFileOut.handle <> INVALID_HANDLE
End Function

Private Function CloseImage( _
) As Boolean

    FileClose hFileOut
    CloseImage = True
End Function

Public Function SaveISO( _
    ByVal file_out As String _
) As Boolean

    Dim lngPathTableBlocks  As Long
    Dim lngPathTableBlocksJ As Long
    Dim lngPathTableBytes   As Long
    Dim lngPathTableBytesJ  As Long
    Dim lngDirTableBlocks   As Long
    Dim lngDirTableBlocksJ  As Long
    Dim lngVolSpace         As Long
    Dim intDirs             As Integer

    If Not OpenImage(file_out) Then Exit Function

    RaiseEvent BuildingFilesystem

    ' create names
    PathAddDosJolNames clsRoot
    FilesAddDosJolNames clsRoot

    ' get the path table size
    lngPathTableBytes = PathTableAddSizes(clsRoot) + RecordSizePathTable(1)

    If blnJoliet Then
        lngPathTableBytesJ = PathTableAddSizesJoliet(clsRoot) + RecordSizePathTable(1)
    End If

    lngPathTableBlocks = RoundBytesByBlocks(lngPathTableBytes) \ ISO_BLOCKSIZE
    lngPathTableBlocksJ = RoundBytesByBlocks(lngPathTableBytesJ) \ ISO_BLOCKSIZE

    ' get the directory record table size
    lngDirTableBlocks = DirectoryTableAddSizes(clsRoot)

    If blnJoliet Then
        lngDirTableBlocksJ = DirectoryTableAddSizesJoliet(clsRoot)
    End If

    ' assign directory record LBAs
    lngVolSpace = AddStartBlocks(clsRoot, _
                                 lngDirTableBlocks + lngDirTableBlocksJ, _
                                 lngPathTableBlocks)

    ' 16 empty blocks + descriptors
    WriteHeader lngVolSpace, _
                lngPathTableBytes, _
                lngPathTableBytesJ, _
                lngPathTableBlocks, _
                lngPathTableBlocksJ, _
                ISO_PATHTABLE_LBA + Abs(blnJoliet), _
                ISO_PATHTABLE_LBA + 1 + lngPathTableBlocks * 2

    ' write the path table
    RaiseEvent WritingPathTable

    clsRoot.DirectoryNumber = 1

    WritePathTable clsRoot, False
    FillLastSector
    WritePathTable clsRoot, True
    FillLastSector

    If blnJoliet Then
        WritePathTableJoliet clsRoot, False
        FillLastSector
        WritePathTableJoliet clsRoot, True
        FillLastSector
    End If

    ' write directory records
    RaiseEvent WritingDirectoryRecords
    WriteDirectoryRecords clsRoot

    If blnJoliet Then WriteDirectoryRecordsJoliet clsRoot

    ' write files
    lngApproxImgSize = ImageSize
    WriteFiles clsRoot
    lngApproxImgSize = 0

    CloseImage

    RaiseEvent WritingFinished

    SaveISO = True
End Function

Private Sub WriteHeader( _
    ByVal VolumeSpace As Long, _
    ByVal PathTableBytes As Long, _
    ByVal PathTableBytesJoliet As Long, _
    ByVal PathTableBlocks As Long, _
    ByVal PathTableBlocksJoliet As Long, _
    ByVal PathTableLBA As Long, _
    ByVal PathTableLBAJoliet As Long _
)

    Dim btEmptyBlock(ISO_BLOCKSIZE - 1) As Byte
    Dim udtPVD                          As ISO_PRIMARY_VOLUME_DESCRIPTOR
    Dim udtSVD                          As ISO_SUPPLEMENTARY_VOLUME_DESCRIPTOR
    Dim udtTVD                          As ISO_VOLUME_DESCRIPTOR_TERMINATOR
    Dim udtRootDir                      As ISO_DIRECTORY_RECORD
    Dim i                               As Long
    Dim strSpace                        As String

    ' write 16 empty blocks
    For i = 1 To 16
        FileWrite hFileOut, VarPtr(btEmptyBlock(0)), ISO_BLOCKSIZE
    Next

    ' Volume Descriptor Set Terminator
    With udtTVD
        .vol_desc_type = VD_END
        CpyMem .standard_id(0), ByVal ISO_STANDARD_ID, 5
        .vol_desc_version = 1
    End With

    ' Root Directory Record
    With udtRootDir
        .length = RecordSizeDirRecord(1)
        .ext_attr_rec_len = 0
        CpyMem .extent(0), LongTo733(clsRoot.Files.LBA), 8
        CpyMem .data_length(0), LongTo733(clsRoot.Files.Blocks * ISO_BLOCKSIZE), 8
        CpyMem .rec_date_time(0), ISODateTimeSmall(dateVolCreation), 7
        .flags = ISO_FLAG_DIRECTORY
        .file_unit_size = 0
        .interleave_gap_size = 0
        CpyMem .vol_seq_num(0), IntegerTo723(1), 4
        .file_id_len = 1
        .file_id(0) = 0
    End With

    ' Primary Volume Descriptor
    With udtPVD
        .vol_desc_type = VD_PRIMARY
        CpyMem .standard_id(0), ByVal ISO_STANDARD_ID, 5
        .vol_desc_version = 1

        CpyMem .system_id(0), ByVal strSystemID, min(32, Len(strSystemID))
        CpyMem .volume_id(0), ByVal strVolumeID, min(32, Len(strVolumeID))

        CpyMem .vol_space_size(0), LongTo733(VolumeSpace), 8
        CpyMem .vol_set_size(0), IntegerTo723(1), 4
        CpyMem .vol_seq_num(0), IntegerTo723(1), 4
        CpyMem .logical_block_size(0), IntegerTo723(ISO_BLOCKSIZE), 4
        CpyMem .path_table_size(0), LongTo733(PathTableBytes), 8
        CpyMem .type_l_path_table(0), PathTableLBA, 4
        CpyMem .type_m_path_table(0), SwapLong(PathTableLBA + PathTableBlocks), 4

        CpyMem .root_directory(0), udtRootDir, 34

        strSpace = Space(128)

        CpyMem .vol_set_id(0), ByVal strSpace, 128
        CpyMem .vol_set_id(0), ByVal strVolSetID, min(128, Len(strVolSetID))
        CpyMem .publisher_id(0), ByVal strSpace, 128
        CpyMem .publisher_id(0), ByVal strPublisherID, min(128, Len(strPublisherID))
        CpyMem .data_prep_id(0), ByVal strSpace, 128
        CpyMem .data_prep_id(0), ByVal strDataPrepID, min(128, Len(strDataPrepID))
        CpyMem .app_id(0), ByVal strSpace, 128
        CpyMem .app_id(0), ByVal strAppID, min(128, Len(strAppID))
        CpyMem .cpy_file_id(0), ByVal strSpace, 37
        CpyMem .cpy_file_id(0), ByVal strCpyFileID, min(37, Len(strCpyFileID))
        CpyMem .abstr_file_id(0), ByVal strSpace, 37
        CpyMem .abstr_file_id(0), ByVal strAbstrFileID, min(37, Len(strAbstrFileID))
        CpyMem .bibl_file_id(0), ByVal strSpace, 37
        CpyMem .bibl_file_id(0), ByVal strBiblFileID, min(37, Len(strBiblFileID))

        CpyMem .vol_creation_date(0), ISODateTime(dateVolCreation), 17
        CpyMem .vol_effective_date(0), ISODateTime(CDate("00:00:00")), 17
        CpyMem .vol_expiration_date(0), ISODateTime(CDate("00:00:00")), 17
        CpyMem .vol_modification_date(0), ISODateTime(CDate("00:00:00")), 17

        .file_struct_version = 1
    End With

    ' Root Directory Record for SVD
    With udtRootDir
        .length = RecordSizeDirRecord(1)
        .ext_attr_rec_len = 0
        CpyMem .extent(0), LongTo733(clsRoot.Files.LBAJoliet), 8
        CpyMem .data_length(0), LongTo733(clsRoot.Files.BlocksJoliet * ISO_BLOCKSIZE), 8
        CpyMem .rec_date_time(0), ISODateTimeSmall(dateVolCreation), 7
        .flags = ISO_FLAG_DIRECTORY
        .file_unit_size = 0
        .interleave_gap_size = 0
        CpyMem .vol_seq_num(0), IntegerTo723(1), 4
        .file_id_len = 1
        .file_id(0) = 0
    End With

    With udtSVD
        .vol_desc_type = VD_SUPPLEMENT
        CpyMem .standard_id(0), ByVal ISO_STANDARD_ID, 5
        .vol_desc_version = 1

        .vol_flags = 0          ' only Escape Sequences from ISO 2375 (?)
        .escape_seqs(0) = &H25  ' hope that works...
        .escape_seqs(1) = &H2F
        .escape_seqs(2) = &H45

        CpySwapBSTR StrPtr(strSystemIDJoliet), VarPtr(.system_id(0)), min(32, LenB(strSystemIDJoliet))
        CpySwapBSTR StrPtr(strVolumeIDJoliet), VarPtr(.volume_id(0)), min(32, LenB(strVolumeIDJoliet))

        CpyMem .vol_space_size(0), LongTo733(VolumeSpace), 8
        CpyMem .vol_set_size(0), IntegerTo723(1), 4
        CpyMem .vol_seq_num(0), IntegerTo723(1), 4
        CpyMem .logical_block_size(0), IntegerTo723(ISO_BLOCKSIZE), 4
        CpyMem .path_table_size(0), LongTo733(PathTableBytesJoliet), 8
        CpyMem .type_l_path_table(0), PathTableLBAJoliet, 4
        CpyMem .type_m_path_table(0), SwapLong(PathTableLBAJoliet + PathTableBlocksJoliet), 4

        CpyMem .root_directory(0), udtRootDir, 34

        strSpace = Space(128)

        CpySwapBSTR StrPtr(strVolSetIDJoliet), VarPtr(.vol_set_id(0)), min(128, LenB(strVolSetIDJoliet))
        CpySwapBSTR StrPtr(strPublisherIDJoliet), VarPtr(.publisher_id(0)), min(128, LenB(strPublisherIDJoliet))
        CpySwapBSTR StrPtr(strDataPrepIDJoliet), VarPtr(.data_prep_id(0)), min(128, LenB(strDataPrepIDJoliet))
        CpySwapBSTR StrPtr(strAppIDJoliet), VarPtr(.app_id(0)), min(128, LenB(strAppIDJoliet))
        CpySwapBSTR StrPtr(strCpyFileIDJoliet), VarPtr(.cpy_file_id(0)), min(37, LenB(strCpyFileIDJoliet))
        CpySwapBSTR StrPtr(strAbstrFileIDJoliet), VarPtr(.abstr_file_id(0)), min(37, LenB(strAbstrFileIDJoliet))
        CpySwapBSTR StrPtr(strBiblFileIDJoliet), VarPtr(.bibl_file_id(0)), min(37, LenB(strBiblFileIDJoliet))

        CpyMem .vol_creation_date(0), ISODateTime(dateVolCreation), 17
        CpyMem .vol_effective_date(0), ISODateTime(CDate("00:00:00")), 17
        CpyMem .vol_expiration_date(0), ISODateTime(CDate("00:00:00")), 17
        CpyMem .vol_modification_date(0), ISODateTime(CDate("00:00:00")), 17

        .file_struct_version = 1
    End With

    FileWrite hFileOut, VarPtr(udtPVD), Len(udtPVD)
    If blnJoliet Then FileWrite hFileOut, VarPtr(udtSVD), Len(udtSVD)
    FileWrite hFileOut, VarPtr(udtTVD), Len(udtTVD)
End Sub

' copy unicode string to a pointer,
' and reverse the byte order
Private Sub CpySwapBSTR( _
    ByVal pFrom As Long, _
    ByVal pTo As Long, _
    ByVal length As Long _
)

    Dim i   As Long

    For i = 1 To length Step 2
        CpyMem ByVal pTo + 0, ByVal pFrom + 1, 1
        CpyMem ByVal pTo + 1, ByVal pFrom + 0, 1

        pFrom = pFrom + 2
        pTo = pTo + 2
    Next
End Sub

' Date for Directory Record
Private Function ISODateTimeSmall( _
    ByVal dtDate As Date _
) As ISO_DIRREC_DATETIME

    Dim udtDate As ISO_DIRREC_DATETIME

    If Year(dtDate) >= 1900 And Year(dtDate) <= 2155 Then
        udtDate.drdt_year = Year(dtDate) - 1900
        udtDate.drdt_month = Month(dtDate)
        udtDate.drdt_day = Day(dtDate)
        udtDate.drdt_hour = Hour(dtDate)
        udtDate.drdt_minute = Minute(dtDate)
        udtDate.drdt_second = Second(dtDate)
    End If

    ISODateTimeSmall = udtDate
End Function

' Date for Volume Descriptor
Private Function ISODateTime( _
    ByVal dtDate As Date _
) As ISO_DATETIME

    Dim udtDate As ISO_DATETIME

    If dtDate = "00:00:00" And Year(dtDate) = 1899 Then
        CpyMem udtDate.dt_year(0), ByVal "0000", 4
        CpyMem udtDate.dt_month(0), ByVal "00", 2
        CpyMem udtDate.dt_day(0), ByVal "00", 2
        CpyMem udtDate.dt_hour(0), ByVal "00", 2
        CpyMem udtDate.dt_minute(0), ByVal "00", 2
        CpyMem udtDate.dt_second(0), ByVal "00", 2
        CpyMem udtDate.dt_hsecond(0), ByVal "00", 2
    Else
        CpyMem udtDate.dt_year(0), ByVal CStr(Year(dtDate)), 4
        CpyMem udtDate.dt_month(0), ByVal Add0(Month(dtDate)), 2
        CpyMem udtDate.dt_day(0), ByVal Add0(Day(dtDate)), 2
        CpyMem udtDate.dt_hour(0), ByVal Add0(Hour(dtDate)), 2
        CpyMem udtDate.dt_minute(0), ByVal Add0(Minute(dtDate)), 2
        CpyMem udtDate.dt_second(0), ByVal Add0(Second(dtDate)), 2
        CpyMem udtDate.dt_hsecond(0), ByVal "00", 2
    End If

    ISODateTime = udtDate
End Function

' add leading zeroes to a string
Private Function Add0( _
    ByVal strText As String _
) As String

    If Not Left$(strText, 1) = "0" Then
        Add0 = "0" & strText
    Else
        Add0 = strText
    End If
End Function

' fill a sector in the ISO Image with null chars
Private Sub FillLastSector()
    Dim lngFPos     As Long
    Dim btEmpty()   As Byte

    lngFPos = FilePosition(hFileOut)
    lngFPos = RoundBytesByBlocks(lngFPos) - lngFPos

    ReDim btEmpty(lngFPos - 1) As Byte

    FileWrite hFileOut, VarPtr(btEmpty(0)), lngFPos
End Sub

' copy files to the ISO image
Private Sub WriteFiles( _
    clsDir As clsISODirectory _
)

    Dim i           As Long
    Dim btFile()    As Byte
    Dim hFS         As hFile

    For i = 0 To clsDir.Files.Count - 1
        RaiseEvent WritingFiles(FilePosition(hFileOut) / lngApproxImgSize * 100)

        If clsDir.Files.Size(i) > 0 Then
            Dim part As Long
            Dim chunk As Long
            Dim remain As Long
            
            chunk = 4096
            part = clsDir.Files.Size(i) \ chunk
            remain = clsDir.Files.Size(i) Mod chunk


            hFS = FileOpen(clsDir.Files.LocalPath(i), GENERIC_READ, , OPEN_EXISTING)
            
            
            Do While part > 0
            
                ReDim btFile(chunk - 1) As Byte
    
    
                If hFS.handle <> INVALID_HANDLE Then
                    FileRead hFS, VarPtr(btFile(0)), chunk
                End If
    
                FileWrite hFileOut, VarPtr(btFile(0)), chunk
                
                part = part - 1
            Loop
            
            If remain > 0 Then
            
                ReDim btFile(remain - 1) As Byte
                If hFS.handle <> INVALID_HANDLE Then
                    FileRead hFS, VarPtr(btFile(0)), remain
                End If
                FileWrite hFileOut, VarPtr(btFile(0)), remain
            End If

            FileClose hFS
            
            FillLastSector
        End If
    Next

    For i = 0 To clsDir.SubDirectoryCount - 1
        WriteFiles clsDir.SubDirectory(i)
    Next
End Sub

Private Sub WriteDirectoryRecordsJoliet( _
    clsDir As clsISODirectory _
)

    Dim i           As Long
    Dim udtEntry    As ISO_DIRECTORY_RECORD
    Dim clsSubDir   As clsISODirectory

    ' .
    With udtEntry
        .length = RecordSizeDirRecord(1)
        .ext_attr_rec_len = 0

        CpyMem .extent(0), LongTo733(clsDir.Files.LBAJoliet), 8
        CpyMem .data_length(0), LongTo733(clsDir.Files.BlocksJoliet * 2048&), 8

        CpyMem .rec_date_time(0), ISODateTimeSmall(dateVolCreation), 7

        .flags = ISO_FLAG_DIRECTORY
        .file_unit_size = 0
        .interleave_gap_size = 0
        CpyMem .vol_seq_num(0), IntegerTo723(1), 4
        .file_id_len = 1

        .file_id(0) = 0
    End With

    FileWrite hFileOut, VarPtr(udtEntry), udtEntry.length

    ' ..
    With udtEntry
        .length = RecordSizeDirRecord(1)
        .ext_attr_rec_len = 0

        If Not clsDir.Parent Is Nothing Then
            CpyMem .extent(0), LongTo733(clsDir.Parent.Files.LBAJoliet), 8
            CpyMem .data_length(0), LongTo733(clsDir.Parent.Files.BlocksJoliet * 2048&), 8
        Else
            CpyMem .extent(0), LongTo733(clsDir.Files.LBAJoliet), 8
            CpyMem .data_length(0), LongTo733(clsDir.Files.BlocksJoliet * 2048&), 8
        End If

        CpyMem .rec_date_time(0), ISODateTimeSmall(dateVolCreation), 7

        .flags = ISO_FLAG_DIRECTORY
        .file_unit_size = 0
        .interleave_gap_size = 0
        CpyMem .vol_seq_num(0), IntegerTo723(1), 4
        .file_id_len = 1
        .file_id(0) = 1
    End With

    FileWrite hFileOut, VarPtr(udtEntry), udtEntry.length

    For i = 0 To clsDir.SubDirectoryCount - 1
        Set clsSubDir = clsDir.SubDirectory(i)

        With udtEntry
            .length = RecordSizeDirRecord(LenB(clsSubDir.JolietName) + 1)
            CpyMem .extent(0), LongTo733(clsSubDir.Files.LBAJoliet), 8
            CpyMem .data_length(0), LongTo733(clsSubDir.Files.BlocksJoliet * 2048&), 8
            .flags = ISO_FLAG_DIRECTORY
            .file_id_len = LenB(clsSubDir.JolietName)
            CpySwapBSTR StrPtr(clsSubDir.JolietName), _
                        VarPtr(.file_id(0)), min(256, _
                        LenB(clsSubDir.JolietName))
        End With

        If FilePosition(hFileOut) + udtEntry.length > RoundBytesByBlocks(FilePosition(hFileOut)) Then
            FillLastSector
        End If

        FileWrite hFileOut, VarPtr(udtEntry), udtEntry.length
    Next

    For i = 0 To clsDir.Files.Count - 1
        With udtEntry
            .length = RecordSizeDirRecord(LenB(clsDir.Files.JolietName(i)) + 1)
            CpyMem .extent(0), LongTo733(clsDir.Files.LBA(i)), 8
            CpyMem .data_length(0), LongTo733(clsDir.Files.Size(i)), 8
            CpyMem .rec_date_time(0), ISODateTimeSmall(clsDir.Files.CreationDate(i)), 7
            .flags = ISO_FLAG_FILE
            .file_id_len = LenB(clsDir.Files.JolietName(i))
            CpySwapBSTR StrPtr(clsDir.Files.JolietName(i)), _
                        VarPtr(.file_id(0)), _
                        min(256, LenB(clsDir.Files.JolietName(i)))
        End With

        If FilePosition(hFileOut) + udtEntry.length > RoundBytesByBlocks(FilePosition(hFileOut)) Then
            FillLastSector
        End If

        FileWrite hFileOut, VarPtr(udtEntry), udtEntry.length
    Next

    FillLastSector

    For i = 0 To clsDir.SubDirectoryCount - 1
        WriteDirectoryRecordsJoliet clsDir.SubDirectory(i)
    Next
End Sub

' write directory records to the ISO image
Private Sub WriteDirectoryRecords( _
    clsDir As clsISODirectory _
)

    Dim i           As Long
    Dim udtEntry    As ISO_DIRECTORY_RECORD
    Dim clsSubDir   As clsISODirectory

    ' .
    With udtEntry
        .length = RecordSizeDirRecord(1)
        .ext_attr_rec_len = 0

        CpyMem .extent(0), LongTo733(clsDir.Files.LBA), 8
        CpyMem .data_length(0), LongTo733(clsDir.Files.Blocks * 2048&), 8

        CpyMem .rec_date_time(0), ISODateTimeSmall(dateVolCreation), 7

        .flags = ISO_FLAG_DIRECTORY
        .file_unit_size = 0
        .interleave_gap_size = 0
        CpyMem .vol_seq_num(0), IntegerTo723(1), 4
        .file_id_len = 1

        .file_id(0) = 0
    End With

    FileWrite hFileOut, VarPtr(udtEntry), udtEntry.length

    ' ..
    With udtEntry
        .length = RecordSizeDirRecord(1)
        .ext_attr_rec_len = 0

        If Not clsDir.Parent Is Nothing Then
            CpyMem .extent(0), LongTo733(clsDir.Parent.Files.LBA), 8
            CpyMem .data_length(0), LongTo733(clsDir.Parent.Files.Blocks * 2048&), 8
        Else
            CpyMem .extent(0), LongTo733(clsDir.Files.LBA), 8
            CpyMem .data_length(0), LongTo733(clsDir.Files.Blocks * 2048&), 8
        End If

        CpyMem .rec_date_time(0), ISODateTimeSmall(dateVolCreation), 7

        .flags = ISO_FLAG_DIRECTORY
        .file_unit_size = 0
        .interleave_gap_size = 0
        CpyMem .vol_seq_num(0), IntegerTo723(1), 4
        .file_id_len = 1
        .file_id(0) = 1
    End With

    FileWrite hFileOut, VarPtr(udtEntry), udtEntry.length

    ' subdirectories
    For i = 0 To clsDir.SubDirectoryCount - 1
        Set clsSubDir = clsDir.SubDirectory(i)

        With udtEntry
            .length = RecordSizeDirRecord(Len(clsSubDir.DOSName) + 1)
            CpyMem .extent(0), LongTo733(clsSubDir.Files.LBA), 8
            CpyMem .data_length(0), LongTo733(clsSubDir.Files.Blocks * 2048&), 8
            .flags = ISO_FLAG_DIRECTORY
            .file_id_len = Len(clsSubDir.DOSName)
            CpyMem .file_id(0), ByVal clsSubDir.DOSName, Len(clsSubDir.DOSName)
        End With

        If FilePosition(hFileOut) + udtEntry.length > RoundBytesByBlocks(FilePosition(hFileOut)) Then
            FillLastSector
        End If

        FileWrite hFileOut, VarPtr(udtEntry), udtEntry.length
    Next

    ' files
    For i = 0 To clsDir.Files.Count - 1
        With udtEntry
            .length = RecordSizeDirRecord(Len(clsDir.Files.DOSName(i)) + 1)
            CpyMem .extent(0), LongTo733(clsDir.Files.LBA(i)), 8
            CpyMem .data_length(0), LongTo733(clsDir.Files.Size(i)), 8
            CpyMem .rec_date_time(0), ISODateTimeSmall(clsDir.Files.CreationDate(i)), 7
            .flags = ISO_FLAG_FILE
            .file_id_len = Len(clsDir.Files.DOSName(i))
            CpyMem .file_id(0), ByVal clsDir.Files.DOSName(i), Len(clsDir.Files.DOSName(i))
        End With

        If FilePosition(hFileOut) + udtEntry.length > RoundBytesByBlocks(FilePosition(hFileOut)) Then
            FillLastSector
        End If

        FileWrite hFileOut, VarPtr(udtEntry), udtEntry.length
    Next

    FillLastSector

    For i = 0 To clsDir.SubDirectoryCount - 1
        WriteDirectoryRecords clsDir.SubDirectory(i)
    Next
End Sub

Private Sub WritePathTableJoliet( _
    clsDir As clsISODirectory, _
    ByVal msb As Boolean _
)

    Dim udtEntry    As ISO_PATH_TABLE_RECORD
    Dim i           As Long
    Dim intDirNum   As Integer
    Dim clsSubDir   As clsISODirectory

    With udtEntry
        .length = 1
        .ext_attr_rec_len = 0

        If msb Then
            CpyMem .extent(0), SwapLong(clsDir.Files.LBAJoliet), 4
        Else
            CpyMem .extent(0), clsDir.Files.LBAJoliet, 4
        End If

        If clsDir.Parent Is Nothing Then
            If msb Then
                CpyMem .parent_dir_num(0), SwapInteger(1), 2
            Else
                CpyMem .parent_dir_num(0), 1, 2
            End If
        Else
            If msb Then
                CpyMem .parent_dir_num(0), SwapInteger(clsDir.Parent.DirectoryNumber), 2
            Else
                CpyMem .parent_dir_num(0), clsDir.Parent.DirectoryNumber, 2
            End If
        End If

        .dir_id(0) = 0
        .dir_id(1) = 0
    End With

    FileWrite hFileOut, VarPtr(udtEntry), RecordSizePathTable(1)

    i = 0
    intDirNum = 2

    Do
        If Not WritePathTableByLayerJoliet(clsDir, i, 0, msb, intDirNum) Then
            Exit Do
        End If

        i = i + 1
    Loop
End Sub

' write the path table
'
' a path table must be written 2 times,
' one of it with reversed byte order.
' Same thing for the Joliet path table.
' Makes 4 path tables.
Private Sub WritePathTable( _
    clsDir As clsISODirectory, _
    ByVal msb As Boolean _
)

    Dim udtEntry    As ISO_PATH_TABLE_RECORD
    Dim i           As Long
    Dim intDirNum   As Integer
    Dim clsSubDir   As clsISODirectory

    With udtEntry
        .length = 1
        .ext_attr_rec_len = 0

        If msb Then
            CpyMem .extent(0), SwapLong(clsDir.Files.LBA), 4
        Else
            CpyMem .extent(0), clsDir.Files.LBA, 4
        End If

        If clsDir.Parent Is Nothing Then
            If msb Then
                CpyMem .parent_dir_num(0), SwapInteger(1), 2
            Else
                CpyMem .parent_dir_num(0), 1, 2
            End If
        Else
            If msb Then
                CpyMem .parent_dir_num(0), SwapInteger(clsDir.Parent.DirectoryNumber), 2
            Else
                CpyMem .parent_dir_num(0), clsDir.Parent.DirectoryNumber, 2
            End If
        End If

        .dir_id(0) = 0
        .dir_id(1) = 0
    End With

    FileWrite hFileOut, VarPtr(udtEntry), RecordSizePathTable(1)

    i = 0
    intDirNum = 2

    Do
        If Not WritePathTableByLayer(clsDir, i, 0, msb, intDirNum) Then
            Exit Do
        End If

        i = i + 1
    Loop
End Sub

Private Function WritePathTableByLayerJoliet( _
    clsDir As clsISODirectory, _
    ByVal layer_search As Long, _
    ByVal layer_current As Long, _
    ByVal msb As Boolean, _
    ByRef dirnum As Integer _
) As Boolean

    Dim i           As Long
    Dim clsSubDir   As clsISODirectory
    Dim udtEntry    As ISO_PATH_TABLE_RECORD

    If layer_current = layer_search Then
        For i = 0 To clsDir.SubDirectoryCount - 1
            Set clsSubDir = clsDir.SubDirectory(i)

            clsSubDir.DirectoryNumber = dirnum
            dirnum = dirnum + 1

            With udtEntry
                .length = min(255, LenB(clsSubDir.JolietName))
                .ext_attr_rec_len = 0

                If msb Then
                    CpyMem .extent(0), SwapLong(clsSubDir.Files.LBAJoliet), 4
                Else
                    CpyMem .extent(0), clsSubDir.Files.LBAJoliet, 4
                End If

                If clsSubDir.Parent Is Nothing Then
                    If msb Then
                        CpyMem .parent_dir_num(0), SwapInteger(1), 2
                    Else
                        CpyMem .parent_dir_num(0), 1, 2
                    End If
                Else
                    If msb Then
                        CpyMem .parent_dir_num(0), SwapInteger(clsSubDir.Parent.DirectoryNumber), 2
                    Else
                        CpyMem .parent_dir_num(0), clsSubDir.Parent.DirectoryNumber, 2
                    End If
                End If

                CpySwapBSTR StrPtr(clsSubDir.JolietName), _
                            VarPtr(.dir_id(0)), _
                            min(256, LenB(clsSubDir.JolietName))
            End With

            FileWrite hFileOut, VarPtr(udtEntry), RecordSizePathTable(udtEntry.length)
        Next

        WritePathTableByLayerJoliet = True
        Exit Function
    End If

    For i = 0 To clsDir.SubDirectoryCount - 1
        If WritePathTableByLayerJoliet(clsDir.SubDirectory(i), layer_search, layer_current + 1, msb, dirnum) Then
            WritePathTableByLayerJoliet = True
        End If
    Next
End Function

' The directories must have a specific order in the path table.
' Roughly like a pyramid, carried of layer by layer:
'
'         root (1)                          Layer 0
'        /       \
'       /         \-_
'  subdir1 (2)        subdir2 (3)           Layer 1
'    |                     |
'    |                     /
'  subdir3 (4)        subdir4 (5)           Layer 2
'                  /         \
'                 /           \
'             subdir5 (6)    subdir6 (7)    Layer 3
'
' The numbers in the brackets describe the order
' of these directories in the path table.
'
Private Function WritePathTableByLayer( _
    clsDir As clsISODirectory, _
    ByVal layer_search As Long, _
    ByVal layer_current As Long, _
    ByVal msb As Boolean, _
    ByRef dirnum As Integer _
) As Boolean

    Dim i           As Long
    Dim clsSubDir   As clsISODirectory
    Dim udtEntry    As ISO_PATH_TABLE_RECORD

    ' found the layer looking for?
    If layer_current = layer_search Then
        ' add all the subdirectories to the path table
        For i = 0 To clsDir.SubDirectoryCount - 1
            Set clsSubDir = clsDir.SubDirectory(i)

            clsSubDir.DirectoryNumber = dirnum
            dirnum = dirnum + 1

            With udtEntry
                .length = min(255, Len(clsSubDir.DOSName))
                .ext_attr_rec_len = 0

                If msb Then
                    CpyMem .extent(0), SwapLong(clsSubDir.Files.LBA), 4
                Else
                    CpyMem .extent(0), clsSubDir.Files.LBA, 4
                End If

                If clsSubDir.Parent Is Nothing Then
                    If msb Then
                        CpyMem .parent_dir_num(0), SwapInteger(1), 2
                    Else
                        CpyMem .parent_dir_num(0), 1, 2
                    End If
                Else
                    If msb Then
                        CpyMem .parent_dir_num(0), SwapInteger(clsSubDir.Parent.DirectoryNumber), 2
                    Else
                        CpyMem .parent_dir_num(0), clsSubDir.Parent.DirectoryNumber, 2
                    End If
                End If

                CpyMem .dir_id(0), ByVal clsSubDir.DOSName, .length
            End With

            FileWrite hFileOut, VarPtr(udtEntry), RecordSizePathTable(udtEntry.length)
        Next

        ' found the layer we have looked for, so return true
        ' to look for further directories on this layer
        WritePathTableByLayer = True
        Exit Function
    End If

    ' didn't find the layer, look in the other subdirectories
    For i = 0 To clsDir.SubDirectoryCount - 1
        If WritePathTableByLayer(clsDir.SubDirectory(i), layer_search, layer_current + 1, msb, dirnum) Then
            WritePathTableByLayer = True
        End If
    Next
End Function

Private Function DebugOutput( _
    clsDir As clsISODirectory _
)

    Dim i   As Long

    Debug.Print "Directory: " & clsDir.name, clsDir.DOSName, clsDir.JolietName
    Debug.Print "   Number: " & clsDir.DirectoryNumber
    Debug.Print "    Files: " & clsDir.Files.Count
    Debug.Print "  FBlocks: " & clsDir.Files.Blocks
    Debug.Print "     FLBA: " & clsDir.Files.LBA

    For i = 0 To clsDir.Files.Count - 1
        With clsDir.Files
            Debug.Print "     File " & i & ":" & .name(i), .DOSName(i), .JolietName(i)
            Debug.Print "         LBA: " & .LBA(i)
        End With
    Next

    Debug.Print ""

    For i = 0 To clsDir.SubDirectoryCount - 1
        DebugOutput clsDir.SubDirectory(i)
    Next
End Function

Public Property Get ImageSize( _
) As Double

    ImageSize = CalcDirSize(clsRoot) + 2048& * 100&
End Property

' get the size of a directory (recursive)
Private Function CalcDirSize( _
    clsDir As clsISODirectory _
) As Double

    Dim clsDirN     As clsISODirectory
    Dim i           As Long
    Dim lngBytes    As Double

    For i = 0 To clsDir.Files.Count - 1
        lngBytes = lngBytes + RoundBytesByBlocks(clsDir.Files.Size(i))
    Next

    For i = 0 To clsDir.SubDirectoryCount - 1
        lngBytes = lngBytes + CalcDirSize(clsDir.SubDirectory(i))
    Next

    CalcDirSize = lngBytes
End Function

Private Function IsDChars( _
    ByVal text As String _
) As Boolean

    Dim btChr   As Byte
    Dim i       As Long

    For i = 1 To Len(text)
        btChr = Asc(Mid$(text, i, 1))

        If Not ((btChr >= &H41 And btChr <= &H5A) Or _
                (btChr >= &H30 And btChr <= &H39) Or _
                (btChr = &H2E) Or (btChr = &H5F)) Then

            Exit Function
        End If
    Next

    IsDChars = True
End Function

Private Function IsAChars( _
    ByVal text As String _
) As Boolean

    Dim btChr   As Byte
    Dim i       As Long

    For i = 1 To Len(text)
        btChr = Asc(Mid$(text, i, 1))

        If Not ((btChr >= &H41 And btChr <= &H5A) Or _
                (btChr >= &H61 And btChr <= &H7A) Or _
                (btChr >= &H20 And btChr <= &H22) Or _
                (btChr >= &H26 And btChr <= &H2F) Or _
                (btChr >= &H3A And btChr <= &H3F) Or _
                (btChr >= &H30 And btChr <= &H39) Or _
                (btChr = &H5F)) Then

            Exit Function
        End If
    Next

    IsAChars = True
End Function

Private Function RecordSizePathTable( _
    ByVal dir_name_len As Long, _
    Optional ByVal count_padding_byte As Long = 1 _
) As Long

    RecordSizePathTable = 8 + dir_name_len + IIf(dir_name_len Mod 2, count_padding_byte, 0)
End Function

Private Function RecordSizeDirRecord( _
    ByVal dir_name_len As Long, _
    Optional ByVal count_padding_byte As Long = 1 _
) As Long

    RecordSizeDirRecord = 32 + dir_name_len + IIf(dir_name_len Mod 2, count_padding_byte, 0)
End Function

Public Property Get VolumeSetID( _
    ByVal Joliet As Boolean _
) As String

    VolumeSetID = IIf(Joliet, strVolSetIDJoliet, strVolSetID)
End Property

Public Property Let VolumeSetID( _
    ByVal Joliet As Boolean, _
    ByVal strNewVal As String _
)

    If Joliet Then
        strVolSetIDJoliet = strNewVal
    Else
        strVolSetID = StringToDChars(strNewVal)
    End If
End Property

Public Property Get VolumeCreation( _
) As Date

    VolumeCreation = dateVolCreation
End Property

Public Property Let VolumeCreation( _
    ByVal dateNewVal As Date _
)

    dateVolCreation = dateNewVal
End Property

Public Property Get DataPreparerID( _
    ByVal Joliet As Boolean _
) As String

    DataPreparerID = IIf(Joliet, strDataPrepIDJoliet, strDataPrepID)
End Property

Public Property Let DataPreparerID( _
    ByVal Joliet As Boolean, _
    ByVal strNewVal As String _
)

    If Joliet Then
        strDataPrepIDJoliet = strNewVal
    Else
        strDataPrepID = StringToAChars(strNewVal)
    End If
End Property

Public Property Get PublisherID( _
    ByVal Joliet As Boolean _
) As String

    PublisherID = IIf(Joliet, strPublisherIDJoliet, strPublisherID)
End Property

Public Property Let PublisherID( _
    ByVal Joliet As Boolean, _
    ByVal strNewVal As String _
)

    If Joliet Then
        strPublisherIDJoliet = strNewVal
    Else
        strPublisherID = StringToAChars(strNewVal)
    End If
End Property

Public Property Get ApplicationID( _
    ByVal Joliet As Boolean _
) As String

    ApplicationID = IIf(Joliet, strAppIDJoliet, strAppID)
End Property

Public Property Let ApplicationID( _
    ByVal Joliet As Boolean, _
    ByVal strNewVal As String _
)

    If Joliet Then
        strAppIDJoliet = strNewVal
    Else
        strAppID = StringToAChars(strNewVal)
    End If
End Property

Public Property Get VolumeID( _
    ByVal Joliet As Boolean _
) As String

    VolumeID = IIf(Joliet, strVolumeIDJoliet, strVolumeID)
End Property

Public Property Let VolumeID( _
    ByVal Joliet As Boolean, _
    ByVal strNewVal As String _
)

    If Joliet Then
        strVolumeIDJoliet = strNewVal
    Else
        strVolumeID = StringToDChars(strNewVal)
    End If
End Property

Public Property Get SystemID( _
    ByVal Joliet As Boolean _
) As String

    SystemID = IIf(Joliet, strSystemIDJoliet, strSystemID)
End Property

Public Property Let SystemID( _
    ByVal Joliet As Boolean, _
    ByVal strNewVal As String _
)

    If Joliet Then
        strSystemIDJoliet = strNewVal
    Else
        strSystemID = StringToAChars(strNewVal)
    End If
End Property

Public Property Get Root( _
) As clsISODirectory

    Set Root = clsRoot
End Property

Private Sub Class_Initialize()
    Set clsRoot = New clsISODirectory

    strSystemID = "WIN32"
    strSystemIDJoliet = "WIN32"
    strVolumeID = "NEU_VOL"
    strVolumeIDJoliet = "Neues Volume"
    strPublisherID = "clsISOWriter"
    strPublisherIDJoliet = "clsISOWriter"
    strAppID = "VB ISO Writer"
    strAppIDJoliet = "VB ISO Writer"

    dateVolCreation = Now
    blnJoliet = True
End Sub

' return a directory from a full path
Public Function DirByPath( _
    ByVal strPath As String _
) As clsISODirectory

    Dim strParts()  As String
    Dim i           As Long
    Dim j           As Long
    Dim clsDir      As clsISODirectory

    Set clsDir = clsRoot

    If strPath = "" Then Exit Function
    If Not Left$(strPath, 1) = "\" Then Exit Function

    If strPath <> "\" Then
        strPath = Mid$(strPath, 2, Len(strPath) - 2)

        strParts = Split(strPath, "\")

        Do
            For i = 0 To clsDir.SubDirectoryCount - 1
                If clsDir.SubDirectory(i).name = strParts(j) Then
                    j = j + 1
                    Set clsDir = clsDir.SubDirectory(i)
                    GoTo NextLoop
                End If
            Next

            If i > clsDir.SubDirectoryCount Then
                Exit Function
            End If

NextLoop:
            If j > UBound(strParts) Then
                Exit Do
            End If
        Loop
    End If

    Set DirByPath = clsDir
End Function

Private Function min( _
    ByVal value1 As Long, _
    ByVal value2 As Long _
) As Long

    min = IIf(value1 < value2, value1, value2)
End Function

Private Function max( _
    ByVal value1 As Long, _
    ByVal value2 As Long _
) As Long

    max = IIf(value1 > value2, value1, value2)
End Function


'Public Property Get LoByte(ByRef lThis As Integer) As Integer
'   LoByte = (lThis And &HFF&)
'End Property
'
'Public Property Let LoByte(ByRef lThis As Integer, ByVal lLoByte As Integer)
'   lThis = lThis And Not &HFF& Or lLoByte
'End Property
'
'Public Property Get HiByte(ByRef lThis As Integer) As Integer
'   If (lThis And &H8000&) = &H8000& Then
'      HiByte = ((lThis And &H7F00&) \ &H100&) Or &H80&
'   Else
'      HiByte = (lThis And &HFF00&) \ &H100&
'   End If
'End Property
'
'Public Property Let HiByte(ByRef lThis As Integer, ByVal lHiByte As Integer)
'   If (lHiByte And &H80&) = &H80& Then
'      lThis = lThis And Not &HFF00& Or ((lHiByte And &H7F&) * &H100&) Or &H8000&
'   Else
'      lThis = lThis And Not &HFF00& Or (lHiByte * &H100&)
'   End If
'End Property
'
'
'
'Public Property Get LoWord(ByRef lThis As Long) As Long
'   LoWord = (lThis And &HFFFF&)
'End Property
'
'Public Property Let LoWord(ByRef lThis As Long, ByVal lLoWord As Long)
'   lThis = lThis And Not &HFFFF& Or lLoWord
'End Property
'
'Public Property Get HiWord(ByRef lThis As Long) As Long
'   If (lThis And &H80000000) = &H80000000 Then
'      HiWord = ((lThis And &H7FFF0000) \ &H10000) Or &H8000&
'   Else
'      HiWord = (lThis And &HFFFF0000) \ &H10000
'   End If
'End Property
'
'Public Property Let HiWord(ByRef lThis As Long, ByVal lHiWord As Long)
'   If (lHiWord And &H8000&) = &H8000& Then
'      lThis = lThis And Not &HFFFF0000 Or ((lHiWord And &H7FFF&) * &H10000) Or &H80000000
'   Else
'      lThis = lThis And Not &HFFFF0000 Or (lHiWord * &H10000)
'   End If
'End Property



' both byte ordered 16 bit int
Private Function IntegerTo723( _
    ByVal intg As Integer _
) As Long


    CpyMem ByVal VarPtr(IntegerTo723) + 0, intg, 2
    CpyMem ByVal VarPtr(IntegerTo723) + 2, SwapInteger(intg), 2
End Function

' both byte ordered 32 bit int
Private Function LongTo733( _
    ByVal lng As Long _
) As Currency
  '  Dim tmp As Long
    
  '  LoWord(tmp) = LoWord(lng)
  '  HiWord(tmp) = HiWord(lng)
  '  LongTo733 = tmp
    
    CpyMem ByVal VarPtr(LongTo733) + 0, lng, 4
    CpyMem ByVal VarPtr(LongTo733) + 4, SwapLong(lng), 4
End Function

' swap bytes of a 16 bit int
Private Function SwapInteger( _
    ByVal word As Integer _
) As Integer

    SwapInteger = ((word And &HFF00) \ &H100) Or _
                  ((word And &HFF) * &H100)
End Function

' swap bytes of a 32 bit int
Private Function SwapLong( _
    ByVal dw As Long _
) As Long

    ' by Mike D Sutton, Mike.Sutton@btclick.com, 20040914
    SwapLong = (((dw And &HFF000000) \ &H1000000) And &HFF&) Or _
                ((dw And &HFF0000) \ &H100&) Or _
                ((dw And &HFF00&) * &H100&) Or _
                ((dw And &H7F&) * &H1000000)

    If (dw And &H80&) Then SwapLong = SwapLong Or &H80000000
End Function

Private Sub Class_Terminate()
    Set clsRoot = Nothing
End Sub
