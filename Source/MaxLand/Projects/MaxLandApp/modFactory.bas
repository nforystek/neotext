Attribute VB_Name = "modFactory"
'Private Declare Function ExeNew Lib "msvbvm60" Alias "__vbaNew" (lpObjectInfo As Any) As IUnknown

Option Explicit

'
' This must be a standard (BAS) module and it MUST be named "modFactory" if things are to work correctly.
'
#Const ALSO_USERCONTROLS = False ' Not tested.
'
Private Declare Function EbExecuteLine Lib "vba6.dll" (ByVal pStringToExec As Long, ByVal Foo1 As Long, ByVal Foo2 As Long, ByVal fCheckOnly As Long) As Long
'
Private Declare Function lstrcmpi Lib "kernel32" Alias "lstrcmpiA" (ByVal s1 As String, ByVal s2 As Long) As Long
Private Declare Function ExeNew Lib "msvbvm60" Alias "__vbaNew" (lpObjectInfo As Any) As IUnknown
Private Declare Function AryPtr Lib "msvbvm60" Alias "VarPtr" (ary() As Any) As Long
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal lpAddress As Long, dst As Any)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal lpAddress As Long, ByVal nv As Long)
'
Private Type EXEPROJECTINFO
    Signature                       As Long
    RuntimeVersion                  As Integer
    BaseLanguageDll(0 To 13)        As Byte
    ExtLanguageDll(0 To 13)         As Byte
    RuntimeRevision                 As Integer
    BaseLangiageDllLCID             As Long
    ExtLanguageDllLCID              As Long
    lpSubMain                       As Long
    lpProjectData                   As Long
    '< There are other fields, but not declared, not needed. >
End Type
'
Private Type ProjectData
    Version                         As Long
    lpModuleDescriptorsTableHeader  As Long
    '< There are other fields, but not declared, not needed. >
End Type
'
Private Type MODDESCRTBL_HEADER
    Reserved0                       As Long
    lpProjectObject                 As Long
    lpProjectExtInfo                As Long
    Reserved1                       As Long
    Reserved2                       As Long
    lpProjectData                   As Long
    GUID(0 To 15)                   As Byte
    Reserved3                       As Integer
    TotalModuleCount                As Integer
    CompiledModuleCount             As Integer
    UsedModuleCount                 As Integer
    lpFirstDescriptor               As Long
    '< There are other fields, but not declared, not needed. >
End Type
'
Private Enum MODFLAGS
    mfBasic = 1
    mfNonStatic = 2
    mfUserControl = &H42000
End Enum
'
Private Type MODDESCRTBL_ENTRY
    lpObjectInfo                    As Long
    FullBits                        As Long
    Placeholder0(0 To 15)           As Byte
    lpszName                        As Long
    MethodsCount                    As Long
    lpMethodNamesArray              As Long
    Placeholder1                    As Long
    ModuleType                      As MODFLAGS
    Placeholder2                    As Long
End Type

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uiAction As Long, ByVal uiParam As Long, pvParam As Any, ByVal fWInIni As Long) As Boolean


Private Const SPI_GETSCREENSAVEACTIVE As Long = &H10
Private Const SPI_GETSCREENSAVERRUNNING As Long = &H72

'Public LastCreateCall As String

Public Function IsScreenSaverActive() As Boolean
    Dim bActive As Boolean
    SystemParametersInfo SPI_GETSCREENSAVEACTIVE, 0, bActive, False
    IsScreenSaverActive = bActive
End Function

Public Function IsScreenSaverRunning() As Boolean
    Dim bRunning As Boolean
    SystemParametersInfo SPI_GETSCREENSAVERRUNNING, 0, bRunning, False
    IsScreenSaverRunning = bRunning
End Function
Public Function NewObject(ByVal Class As String) As IUnknown
    '
    ' When you work in the compiled form and the different mechanisms will be used by the IDE.
#If VBIDE = -1 Then
 '   If InIDE Then

        Set NewObject = IdeCreateInstance(Class)
  '  Else
#Else
        Set NewObject = ExeCreateInstance(Class)
#End If
'#If VBIDE = -1 Then
 '   End If
'#End If
End Function

Private Function IdeCreateInstance(ByVal Class As String) As IUnknown
    ' Only for IDE.
    '
    ' IMPORTANT: The module this is in MUST be named modFactory.
    EbExecuteLine StrPtr("modFactory.OneCellQueue New " & UCase(Left(Class, 1)) & LCase(Mid(Class, 2))), 0, 0, 0
    
    'LastCreateCall = UCase(Left(Class, 1)) & LCase(Mid(Class, 2))
    'EbExecuteLine StrPtr("modFactory.OneCellQueue New " & LastCreateCall), 0, 0, 0
    '
    Set IdeCreateInstance = OneCellQueue(Nothing)
    If IdeCreateInstance Is Nothing Then
        Err.Raise 8, , "Specified class '" + Class + "' is not defined."
        Exit Function
    End If
     EbExecuteLine StrPtr("modFactory.OneCellQueue Nothing"), 0, 0, 0 'reset not to keep hanging references
End Function

Private Function OneCellQueue(ByVal refIn As IUnknown) As IUnknown
    ' Returns what was "previously" passed in as refIn,
    ' and then stores the current refIn for return next time.
    '
    Static o As IUnknown
    '
    Set OneCellQueue = o
    Set o = refIn
End Function

Private Function ExeCreateInstance(ByVal Class As String) As IUnknown
    ' Only for Executable.
    '
    Dim lpObjectInformation As Long
    '
    ' Get the address of a block of information about the class.
    ' And then create an instance of this class.
    ' If a class is not found, generated an error.
    '
    If Not GetOiOfClass(Class, lpObjectInformation) Then
        Err.Raise 8, , "Specified class '" + Class + "' does not defined."
        Exit Function
    End If
    '
    Set ExeCreateInstance = ExeNew(ByVal lpObjectInformation)
End Function

Private Function GetOiOfClass(ByVal Class As String, ByRef lpObjInfo As Long) As Boolean
    ' Only for Executable.
    '
    ' lpObjInfo is a returned argument.
    ' Function returns true if successful.
    '
    Static Modules()        As modFactory.MODDESCRTBL_ENTRY
    Static bModulesSet      As Boolean
    Dim i                   As Long
    '
    #If ALSO_USERCONTROLS Then
        Const mfBadFlags As Long = mfUserControl
    #Else
        Const mfBadFlags As Long = 0
    #End If
    '
    If Not bModulesSet Then
        ReDim Modules(0)
        If LoadDescriptorsTable(Modules) Then
            bModulesSet = True
        Else
            Exit Function
        End If
    End If
    '
    ' We are looking for a descriptor corresponding to the specified class.
    For i = LBound(Modules) To UBound(Modules)
        With Modules(i)
        If lstrcmpi(Class, .lpszName) = 0 And CBool(.ModuleType And mfNonStatic) And Not CBool(.ModuleType And mfBadFlags) Then
                lpObjInfo = .lpObjectInfo
                GetOiOfClass = True
                Exit Function
            End If
        End With
    Next i
End Function

Private Function LoadDescriptorsTable(dt() As MODDESCRTBL_ENTRY) As Boolean
    ' Only for Executable.
    '
    Dim lpEPI               As Long
    Dim EPI(0)              As modFactory.EXEPROJECTINFO
    Dim ProjectData(0)      As modFactory.ProjectData
    Dim ModDescrTblHdr(0)   As modFactory.MODDESCRTBL_HEADER
    '
    ' This procedure is called only once for the project.
    ' Get the address of the EPI.
    '
    If Not FindEpiSimple(lpEPI) Then
        Err.Raise 17, , "Failed to locate EXEPROJECTINFO structure in process module image."
        Exit Function
    End If
    '
    ' From EPI find location PROJECTDATA, from PROJECTDATA obtain location
    ' of Table header tags, the title tags, and obtain the number of address sequence.
    '
    SaMap AryPtr(EPI), lpEPI
    SaMap AryPtr(ProjectData), EPI(0).lpProjectData: SaUnmap AryPtr(EPI)
    SaMap AryPtr(ModDescrTblHdr), ProjectData(0).lpModuleDescriptorsTableHeader: SaUnmap AryPtr(ProjectData)
    SaMap AryPtr(dt), ModDescrTblHdr(0).lpFirstDescriptor, ModDescrTblHdr(0).TotalModuleCount: SaUnmap AryPtr(ModDescrTblHdr)
    '
    LoadDescriptorsTable = True
End Function

Private Function FindEpiSimple(ByRef lpEPI As Long) As Boolean
    ' Only for Executable.
    '
    Dim DWords()            As Long: ReDim DWords(0)
    Dim PotentionalEPI(0)   As modFactory.EXEPROJECTINFO
    Dim PotentionalPD(0)    As modFactory.ProjectData
    Dim i                   As Long
    '
    Const EPI_Signature     As Long = &H21354256 ' "VB5/6!"
    Const PD_Version        As Long = &H1F4
    '
    ' We are trying to get a pointer to a structure EXEPROJECTINFO. The address is not stored anywhere.
    ' Therefore the only way to find the structure - find its signature.
    '
    ' Current research implementation simply disgusting: it is looking for signatures from the
    ' very beginning of the image, including those places where it can not be known. And find out
    ' Behind the border of the image, if you find a signature within the virtual image failed.
    ' This will likely result in AV-exclusion. But its (implementation) is compact.
    '
    SaMap AryPtr(DWords), App.hInstance
    Do
        If DWords(i) = EPI_Signature Then
            SaMap AryPtr(PotentionalEPI), VarPtr(DWords(i))
            SaMap AryPtr(PotentionalPD), PotentionalEPI(0).lpProjectData
            If PotentionalPD(0).Version = PD_Version Then
                lpEPI = VarPtr(DWords(i))
                FindEpiSimple = True
            End If
            SaUnmap AryPtr(PotentionalPD)
            SaUnmap AryPtr(PotentionalEPI)
            If FindEpiSimple Then Exit Do
        End If
        i = i + 1
    Loop
    SaUnmap AryPtr(DWords)
End Function

Private Sub SaMap(ByVal ppSA As Long, ByVal pMemory As Long, Optional ByVal NewSize As Long = -1)
    Dim pSA As Long: GetMem4 ppSA, pSA:
    PutMem4 pSA + 12, ByVal pMemory: PutMem4 pSA + 16, ByVal NewSize
End Sub

Private Sub SaUnmap(ByVal ppSA As Long)
    Dim pSA As Long: GetMem4 ppSA, pSA
    PutMem4 pSA + 12, ByVal 0: PutMem4 pSA + 16, ByVal 0
End Sub

Public Function InIDE() As Boolean
#If VBIDE = 0 Then
    InIDE = False
    Exit Function
#End If
    On Error GoTo InTheIDE

    Debug.Print 1 / 0
    Exit Function

InTheIDE:
    InIDE = True

End Function
