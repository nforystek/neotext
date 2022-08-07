Attribute VB_Name = "modDialog"
#Const modDialog = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Const GWL_HINSTANCE = (-6)
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOACTIVATE = &H10
Private Const HCBT_ACTIVATE = 5
Private Const WH_CBT = 5

Private hHook As Long

Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function ChooseColor Lib "comdlg32" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLORS) As Long
Private Declare Function CommDlgExtendedError Lib "comdlg32" () As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function ChooseFont Lib "comdlg32" Alias "ChooseFontA" (pChoosefont As CHOOSEFONTS) As Long
Private Declare Function PrintDlg Lib "comdlg32" Alias "PrintDlgA" (pPrintdlg As PRINTDLGS) As Long

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_EXPLORER = &H80000
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_NOLONGNAMES = &H40000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_READONLY = &H1
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHAREWARN = 0
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHOWHELP = &H10
Public Const OFS_MAXPATHNAME = 256

Public Const LF_FACESIZE = 32

Public Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_CREATEPROMPT Or OFN_NODEREFERENCELINKS Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
Public Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY

Public Type OPENFILENAME
    nStructSize As Long
    hwndOwner As Long
    hInstance As Long
    sFilter As String
    sCustomFilter As String
    nCustFilterSize As Long
    nFilterIndex As Long
    sFile As String
    nFileSize As Long
    sFileTitle As String
    nTitleSize As Long
    sInitDir As String
    sDlgTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExt As Integer
    sDefFileExt As String
    nCustDataSize As Long
    fnHook As Long
    sTemplateName As String
End Type

Private Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type

Private Type OFNOTIFY
        hdr As NMHDR
        lpOFN As OPENFILENAME
        pszFile As String
End Type

Private Type CHOOSECOLORS
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Public Type CHOOSEFONTS
    lStructSize As Long
    hwndOwner As Long
    hdc As Long
    lpLogFont As Long
    iPointSize As Long
    flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type

Public Const CC_RGBINIT = &H1
Public Const CC_FULLOPEN = &H2
Public Const CC_PREVENTFULLOPEN = &H4
Public Const CC_SHOWHELP = &H8
Public Const CC_ENABLEHOOK = &H10
Public Const CC_ENABLETEMPLATE = &H20
Public Const CC_ENABLETEMPLATEHANDLE = &H40
Public Const CC_SOLIDCOLOR = &H80
Public Const CC_ANYCOLOR = &H100

Public Const COLOR_FLAGS = CC_FULLOPEN Or CC_ANYCOLOR Or CC_RGBINIT

Public Const CF_SCREENFONTS = &H1
Public Const CF_PRINTERFONTS = &H2
Public Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Public Const CF_SHOWHELP = &H4&
Public Const CF_ENABLEHOOK = &H8&
Public Const CF_ENABLETEMPLATE = &H10&
Public Const CF_ENABLETEMPLATEHANDLE = &H20&
Public Const CF_INITTOLOGFONTSTRUCT = &H40&
Public Const CF_USESTYLE = &H80&
Public Const CF_EFFECTS = &H100&
Public Const CF_APPLY = &H200&
Public Const CF_ANSIONLY = &H400&
Public Const CF_SCRIPTSONLY = CF_ANSIONLY
Public Const CF_NOVECTORFONTS = &H800&
Public Const CF_NOOEMFONTS = CF_NOVECTORFONTS
Public Const CF_NOSIMULATIONS = &H1000&
Public Const CF_LIMITSIZE = &H2000&
Public Const CF_FIXEDPITCHONLY = &H4000&
Public Const CF_WYSIWYG = &H8000
Public Const CF_FORCEFONTEXIST = &H10000
Public Const CF_SCALABLEONLY = &H20000
Public Const CF_TTONLY = &H40000
Public Const CF_NOFACESEL = &H80000
Public Const CF_NOSTYLESEL = &H100000
Public Const CF_NOSIZESEL = &H200000
Public Const CF_SELECTSCRIPT = &H400000
Public Const CF_NOSCRIPTSEL = &H800000
Public Const CF_NOVERTFONTS = &H1000000

Public Const SIMULATED_FONTTYPE = &H8000
Public Const PRINTER_FONTTYPE = &H4000
Public Const SCREEN_FONTTYPE = &H2000
Public Const BOLD_FONTTYPE = &H100
Public Const ITALIC_FONTTYPE = &H200
Public Const REGULAR_FONTTYPE = &H400

Public Const LBSELCHSTRING = "commdlg_LBSelChangedNotify"
Public Const SHAREVISTRING = "commdlg_ShareViolation"
Public Const FILEOKSTRING = "commdlg_FileNameOK"
Public Const COLOROKSTRING = "commdlg_ColorOK"
Public Const SETRGBSTRING = "commdlg_SetRGBColor"
Public Const HELPMSGSTRING = "commdlg_help"
Public Const FINDMSGSTRING = "commdlg_FindReplace"

Public Const CD_LBSELNOITEMS = -1
Public Const CD_LBSELCHANGE = 0
Public Const CD_LBSELSUB = 1
Public Const CD_LBSELADD = 2

Private Type PRINTDLGS
        lStructSize As Long
        hwndOwner As Long
        hDevMode As Long
        hDevNames As Long
        hdc As Long
        flags As Long
        nFromPage As Integer
        nToPage As Integer
        nMinPage As Integer
        nMaxPage As Integer
        nCopies As Integer
        hInstance As Long
        lCustData As Long
        lpfnPrintHook As Long
        lpfnSetupHook As Long
        lpPrintTemplateName As String
        lpSetupTemplateName As String
        hPrintTemplate As Long
        hSetupTemplate As Long
End Type

Public Const PD_ALLPAGES = &H0
Public Const PD_SELECTION = &H1
Public Const PD_PAGENUMS = &H2
Public Const PD_NOSELECTION = &H4
Public Const PD_NOPAGENUMS = &H8
Public Const PD_COLLATE = &H10
Public Const PD_PRINTTOFILE = &H20
Public Const PD_PRINTSETUP = &H40
Public Const PD_NOWARNING = &H80
Public Const PD_RETURNDC = &H100
Public Const PD_RETURNIC = &H200
Public Const PD_RETURNDEFAULT = &H400
Public Const PD_SHOWHELP = &H800
Public Const PD_ENABLEPRINTHOOK = &H1000
Public Const PD_ENABLESETUPHOOK = &H2000
Public Const PD_ENABLEPRINTTEMPLATE = &H4000
Public Const PD_ENABLESETUPTEMPLATE = &H8000
Public Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
Public Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
Public Const PD_USEDEVMODECOPIES = &H40000
Public Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000
Public Const PD_DISABLEPRINTTOFILE = &H80000
Public Const PD_HIDEPRINTTOFILE = &H100000
Public Const PD_NONETWORKBUTTON = &H200000

Private Type DEVNAMES
        wDriverOffset As Integer
        wDeviceOffset As Integer
        wOutputOffset As Integer
        wDefault As Integer
End Type

Public Const DN_DEFAULTPRN = &H1

Public Type SelectedFile
    nFilesSelected As Integer
    sFiles() As String
    sLastDirectory As String
    bCanceled As Boolean
End Type

Public Type SelectedColor
    oSelectedColor As Variant
    bCanceled As Boolean
End Type

Public Type SelectedFont
    sSelectedFont As String
    bCanceled As Boolean
    bBold As Boolean
    bItalic As Boolean
    nSize As Integer
    bUnderline As Boolean
    bStrikeOut As Boolean
    lColor As Long
    sFaceName As String
End Type

Public FileDialog As OPENFILENAME
Public ColorDialog As CHOOSECOLORS
Public FontDialog As CHOOSEFONTS
Public PrintDialog As PRINTDLGS
Private ParenthWnd As Long

Public Function ShowOpen(ByVal hwnd As Long, Optional ByVal centerForm As Boolean = True) As SelectedFile
    Dim Ret As Long
    Dim Count As Integer
    Dim fileNameHolder As String
    Dim LastCharacter As Integer
    Dim NewCharacter As Integer
    Dim tempFiles(1 To 200) As String
    Dim hInst As Long
    Dim Thread As Long
    
    ParenthWnd = hwnd
    FileDialog.nStructSize = Len(FileDialog)
    FileDialog.hwndOwner = hwnd
    FileDialog.sFileTitle = Space$(2048)
    FileDialog.nTitleSize = Len(FileDialog.sFileTitle)
    FileDialog.sFile = FileDialog.sFile & Space$(2047) & Chr$(0)
    FileDialog.nFileSize = Len(FileDialog.sFile)
    
    If Not CBool(FileDialog.flags And OFS_FILE_OPEN_FLAGS) Then
        FileDialog.flags = FileDialog.flags Or OFS_FILE_OPEN_FLAGS
    End If
    
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    Ret = GetOpenFileName(FileDialog)

    If Ret Then
        If Trim$(Replace(FileDialog.sFileTitle, Chr(0), "")) = "" Then
            LastCharacter = 0
            Count = 0
            While ShowOpen.nFilesSelected = 0
                NewCharacter = InStr(LastCharacter + 1, FileDialog.sFile, Chr$(0), vbTextCompare)
                If Count > 0 Then
                    tempFiles(Count) = Mid(FileDialog.sFile, LastCharacter + 1, NewCharacter - LastCharacter - 1)
                Else
                    ShowOpen.sLastDirectory = Mid(FileDialog.sFile, LastCharacter + 1, NewCharacter - LastCharacter - 1)
                End If
                Count = Count + 1
                If InStr(NewCharacter + 1, FileDialog.sFile, Chr$(0), vbTextCompare) = InStr(NewCharacter + 1, FileDialog.sFile, Chr$(0) & Chr$(0), vbTextCompare) Then
                    tempFiles(Count) = Mid(FileDialog.sFile, NewCharacter + 1, InStr(NewCharacter + 1, FileDialog.sFile, Chr$(0) & Chr$(0), vbTextCompare) - NewCharacter - 1)
                    ShowOpen.nFilesSelected = Count
                End If
                LastCharacter = NewCharacter
            Wend
            ReDim ShowOpen.sFiles(1 To ShowOpen.nFilesSelected)
            For Count = 1 To ShowOpen.nFilesSelected
                ShowOpen.sFiles(Count) = tempFiles(Count)
            Next
        Else
            ReDim ShowOpen.sFiles(1 To 1)
            ShowOpen.sLastDirectory = Left$(FileDialog.sFile, FileDialog.nFileOffset)
            ShowOpen.nFilesSelected = 1
            ShowOpen.sFiles(1) = Mid(FileDialog.sFile, FileDialog.nFileOffset + 1, InStr(1, FileDialog.sFile, Chr$(0), vbTextCompare) - FileDialog.nFileOffset - 1)
        End If
        ShowOpen.bCanceled = False
        Exit Function
    Else
        ShowOpen.sLastDirectory = ""
        ShowOpen.nFilesSelected = 0
        ShowOpen.bCanceled = True
        Erase ShowOpen.sFiles
        Exit Function
    End If
End Function

Public Function ShowSave(ByVal hwnd As Long, Optional ByVal centerForm As Boolean = True) As SelectedFile
    Dim Ret As Long
    Dim hInst As Long
    Dim Thread As Long
    
    ParenthWnd = hwnd
    FileDialog.nStructSize = Len(FileDialog)
    FileDialog.hwndOwner = hwnd
    FileDialog.sFileTitle = Space$(2048)
    FileDialog.nTitleSize = Len(FileDialog.sFileTitle)
    FileDialog.sFile = Space$(2047) & Chr$(0)
    FileDialog.nFileSize = Len(FileDialog.sFile)
    
    If Not CBool(FileDialog.flags And OFS_FILE_SAVE_FLAGS) Then
        FileDialog.flags = FileDialog.flags Or OFS_FILE_SAVE_FLAGS
    End If

    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    Ret = GetSaveFileName(FileDialog)
    ReDim ShowSave.sFiles(1)

    If Ret Then
        ShowSave.sLastDirectory = Left$(FileDialog.sFile, FileDialog.nFileOffset)
        ShowSave.nFilesSelected = 1
        ShowSave.sFiles(1) = Trim(Replace(FileDialog.sFile, Chr(0), ""))
        ShowSave.bCanceled = False
        Exit Function
    Else
        ShowSave.sLastDirectory = ""
        ShowSave.nFilesSelected = 0
        ShowSave.bCanceled = True
        Erase ShowSave.sFiles
        Exit Function
    End If
End Function

Public Function ShowColor(ByVal hwnd As Long, Optional ByVal centerForm As Boolean = True) As SelectedColor
    Dim customcolors() As Byte
    Dim i As Integer
    Dim Ret As Long
    Dim hInst As Long
    Dim Thread As Long

    ParenthWnd = hwnd
    If ColorDialog.lpCustColors = "" Then
        ReDim customcolors(0 To 16 * 4 - 1) As Byte
    
        For i = LBound(customcolors) To UBound(customcolors)
          customcolors(i) = 254
        Next i
        
        ColorDialog.lpCustColors = StrConv(customcolors, vbUnicode)
    End If
    
    ColorDialog.hwndOwner = hwnd
    ColorDialog.lStructSize = Len(ColorDialog)
    ColorDialog.flags = COLOR_FLAGS

    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    Ret = ChooseColor(ColorDialog)
    If Ret Then
        ShowColor.bCanceled = False
        ShowColor.oSelectedColor = ColorDialog.rgbResult
        Exit Function
    Else
        ShowColor.bCanceled = True
        ShowColor.oSelectedColor = &H0&
        Exit Function
    End If
End Function

Public Function ShowFont(ByVal hwnd As Long, ByVal startingFontName As String, Optional ByVal centerForm As Boolean = True) As SelectedFont
    Dim Ret As Long
    Dim lfLogFont As LOGFONT
    Dim hInst As Long
    Dim Thread As Long
    Dim i As Integer
    
    ParenthWnd = hwnd
    FontDialog.nSizeMax = 0
    FontDialog.nSizeMin = 0
    FontDialog.nFontType = Screen.FontCount
    FontDialog.hwndOwner = hwnd
    FontDialog.hdc = 0
    FontDialog.lpfnHook = 0
    FontDialog.lCustData = 0
    FontDialog.lpLogFont = VarPtr(lfLogFont)
    If FontDialog.iPointSize = 0 Then
        FontDialog.iPointSize = 10 * 10
    End If
    FontDialog.lpTemplateName = Space$(2048)
    FontDialog.rgbColors = RGB(0, 255, 255)
    FontDialog.lStructSize = Len(FontDialog)
    
    If FontDialog.flags = 0 Then
        FontDialog.flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT 'Or CF_EFFECTS
    End If
    
    For i = 0 To Len(startingFontName) - 1
        lfLogFont.lfFaceName(i) = Asc(Mid(startingFontName, i + 1, 1))
    Next

    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    Ret = ChooseFont(FontDialog)
        
    If Ret Then
        ShowFont.bCanceled = False
        ShowFont.bBold = IIf(lfLogFont.lfWeight > 400, 1, 0)
        ShowFont.bItalic = lfLogFont.lfItalic
        ShowFont.bStrikeOut = lfLogFont.lfStrikeOut
        ShowFont.bUnderline = lfLogFont.lfUnderline
        ShowFont.lColor = FontDialog.rgbColors
        ShowFont.nSize = FontDialog.iPointSize / 10
        For i = 0 To 31
            ShowFont.sSelectedFont = ShowFont.sSelectedFont + Chr(lfLogFont.lfFaceName(i))
        Next
    
        ShowFont.sSelectedFont = Mid(ShowFont.sSelectedFont, 1, InStr(1, ShowFont.sSelectedFont, Chr(0)) - 1)
        Exit Function
    Else
        ShowFont.bCanceled = True
        Exit Function
    End If
End Function
Public Function ShowPrinter(ByVal hwnd As Long, Optional ByVal centerForm As Boolean = True) As Long
    Dim hInst As Long
    Dim Thread As Long
    
    ParenthWnd = hwnd
    PrintDialog.hwndOwner = hwnd
    PrintDialog.lStructSize = Len(PrintDialog)

    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    ShowPrinter = PrintDlg(PrintDialog)
End Function
Private Function WinProcCenterScreen(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim rectForm As RECT, rectMsg As RECT
    Dim X As Long, Y As Long
    If lMsg = HCBT_ACTIVATE Then

        GetWindowRect wParam, rectMsg
        X = Screen.Width / Screen.TwipsPerPixelX / 2 - (rectMsg.Right - rectMsg.Left) / 2
        Y = Screen.Height / Screen.TwipsPerPixelY / 2 - (rectMsg.Bottom - rectMsg.Top) / 2

        SetWindowPos wParam, 0, X, Y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE

        UnhookWindowsHookEx hHook
    End If
    WinProcCenterScreen = False
End Function

Private Function WinProcCenterForm(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim rectForm As RECT, rectMsg As RECT
    Dim X As Long, Y As Long

    If lMsg = HCBT_ACTIVATE Then

        GetWindowRect ParenthWnd, rectForm
        GetWindowRect wParam, rectMsg
        X = (rectForm.Left + (rectForm.Right - rectForm.Left) / 2) - ((rectMsg.Right - rectMsg.Left) / 2)
        Y = (rectForm.Top + (rectForm.Bottom - rectForm.Top) / 2) - ((rectMsg.Bottom - rectMsg.Top) / 2)

        SetWindowPos wParam, 0, X, Y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE

        UnhookWindowsHookEx hHook
     End If
     WinProcCenterForm = False
End Function





