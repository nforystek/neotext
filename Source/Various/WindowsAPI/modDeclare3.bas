Attribute VB_Name = "modDeclare3"
Option Explicit
' ------------------------------------------------------------------------
'
'    WIN32API.TXT -- Win32 API Declarations for Visual Basic
'
'              Copyright (C) 1994-98 Microsoft Corporation
'
'  This file is required for the Visual Basic 6.0 version of the APILoader.
'  Older versions of this file will not work correctly with the version
'  6.0 APILoader.  This file is backwards compatible with previous releases
'  of the APILoader with the exception that Constants are no longer declared
'  as Global or Public in this file.
'
'  This file contains only the Const, Type,
'  and Public Declare statements for  Win32 APIs.
'
'  You have a royalty-free right to use, modify, reproduce and distribute
'  this file (and/or any modified version) in any way you find useful,
'  provided that you agree that Microsoft has no warranty, obligation or
'  liability for its contents.  Refer to the Microsoft Windows Programmer's
'  Reference for further information.
'
' ------------------------------------------------------------------------

' Locale Independent APIs

Public Declare Function GetStringTypeA Lib "kernel32" (ByVal lcid As Long, ByVal dwInfoType As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, lpCharType As Long) As Long
Public Declare Function FoldString Lib "kernel32" Alias "FoldStringA" (ByVal dwMapFlags As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, ByVal lpDestStr As String, ByVal cchDest As Long) As Long
Public Declare Function EnumTimeFormats Lib "kernel32" (ByVal lpTimeFmtEnumProc As Long, ByVal Locale As Long, ByVal dwFlags As Long) As Long
Public Declare Function EnumDateFormats Lib "kernel32" (ByVal lpDateFmtEnumProc As Long, ByVal Locale As Long, ByVal dwFlags As Long) As Long
Public Declare Function IsValidLocale Lib "kernel32" (ByVal Locale As Long, ByVal dwFlags As Long) As Long
Public Declare Function ConvertDefaultLocale Lib "kernel32" (ByVal Locale As Long) As Long
Public Declare Function GetThreadLocale Lib "kernel32" () As Long
Public Declare Function EnumSystemLocales Lib "kernel32" (ByVal lpLocaleEnumProc As Long, ByVal dwFlags As Long) As Long
Public Declare Function EnumSystemCodePages Lib "kernel32" (ByVal lpCodePageEnumProc As Long, ByVal dwFlags As Long) As Long
Public Declare Function ReadConsoleOutput Lib "kernel32" Alias "ReadConsoleOutputA" (ByVal hConsoleOutput As Long, lpBuffer As CHAR_INFO, dwBufferSize As COORD, dwBufferCoord As COORD, lpReadRegion As SMALL_RECT) As Long
Public Declare Function WriteConsoleOutput Lib "kernel32" Alias "WriteConsoleOutputA" (ByVal hConsoleOutput As Long, lpBuffer As CHAR_INFO, dwBufferSize As COORD, dwBufferCoord As COORD, lpWriteRegion As SMALL_RECT) As Long
Public Declare Function ReadConsoleOutputCharacter Lib "kernel32" Alias "ReadConsoleOutputCharacterA" (ByVal hConsoleOutput As Long, ByVal lpCharacter As String, ByVal nLength As Long, dwReadCoord As COORD, lpNumberOfCharsRead As Long) As Long
Public Declare Function ReadConsoleOutputAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, lpAttribute As Long, ByVal nLength As Long, dwReadCoord As COORD, lpNumberOfAttrsRead As Long) As Long
Public Declare Function WriteConsoleOutputCharacter Lib "kernel32" Alias "WriteConsoleOutputCharacterA" (ByVal hConsoleOutput As Long, ByVal lpCharacter As String, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfCharsWritten As Long) As Long
Public Declare Function WriteConsoleOutputAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, lpAttribute As Integer, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfAttrsWritten As Long) As Long
Public Declare Function FillConsoleOutputCharacter Lib "kernel32" Alias "FillConsoleOutputCharacterA" (ByVal hConsoleOutput As Long, ByVal cCharacter As Byte, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfCharsWritten As Long) As Long
Public Declare Function FillConsoleOutputAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttribute As Long, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfAttrsWritten As Long) As Long
Public Declare Function GetConsoleMode Lib "kernel32" (ByVal hConsoleHandle As Long, lpMode As Long) As Long
Public Declare Function GetNumberOfConsoleInputEvents Lib "kernel32" (ByVal hConsoleInput As Long, lpNumberOfEvents As Long) As Long
Public Declare Function GetConsoleScreenBufferInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long
Public Declare Function GetLargestConsoleWindowSize Lib "kernel32" (ByVal hConsoleOutput As Long) As COORD
Public Declare Function GetConsoleCursorInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
Public Declare Function GetNumberOfConsoleMouseButtons Lib "kernel32" (lpNumberOfMouseButtons As Long) As Long
Public Declare Function SetConsoleMode Lib "kernel32" (ByVal hConsoleHandle As Long, ByVal dwMode As Long) As Long
Public Declare Function SetConsoleActiveScreenBuffer Lib "kernel32" (ByVal hConsoleOutput As Long) As Long
Public Declare Function FlushConsoleInputBuffer Lib "kernel32" (ByVal hConsoleInput As Long) As Long
Public Declare Function SetConsoleScreenBufferSize Lib "kernel32" (ByVal hConsoleOutput As Long, dwSize As COORD) As Long
Public Declare Function SetConsoleCursorPosition Lib "kernel32" (ByVal hConsoleOutput As Long, dwCursorPosition As COORD) As Long
Public Declare Function SetConsoleCursorInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
Public Declare Function ScrollConsoleScreenBuffer Lib "kernel32" Alias "ScrollConsoleScreenBufferA" (ByVal hConsoleOutput As Long, lpScrollRectangle As SMALL_RECT, lpClipRectangle As SMALL_RECT, dwDestinationOrigin As COORD, lpFill As CHAR_INFO) As Long
Public Declare Function SetConsoleWindowInfo Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal bAbsolute As Long, lpConsoleWindow As SMALL_RECT) As Long
Public Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
Public Declare Function SetConsoleCtrlHandler Lib "kernel32" (ByVal HandlerRoutine As Long, ByVal Add As Long) As Long
Public Declare Function GenerateConsoleCtrlEvent Lib "kernel32" (ByVal dwCtrlEvent As Long, ByVal dwProcessGroupId As Long) As Long
Public Declare Function AllocConsole Lib "kernel32" () As Long
Public Declare Function FreeConsole Lib "kernel32" () As Long
Public Declare Function GetConsoleTitle Lib "kernel32" Alias "GetConsoleTitleA" (ByVal lpConsoleTitle As String, ByVal nSize As Long) As Long
Public Declare Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
Public Declare Function ReadConsole Lib "kernel32" Alias "ReadConsoleA" (ByVal hConsoleInput As Long, lpBuffer As Any, ByVal nNumberOfCharsToRead As Long, lpNumberOfCharsRead As Long, lpReserved As Any) As Long
Public Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Public Declare Function CreateConsoleScreenBuffer Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwFlags As Long, lpScreenBufferData As Any) As Long
Public Declare Function GetConsoleCP Lib "kernel32" () As Long
Public Declare Function SetConsoleCP Lib "kernel32" (ByVal wCodePageID As Long) As Long
Public Declare Function GetConsoleOutputCP Lib "kernel32" () As Long
Public Declare Function SetConsoleOutputCP Lib "kernel32" (ByVal wCodePageID As Long) As Long
Public Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Public Declare Function AnimatePalette Lib "gdi32" Alias "AnimatePaletteA" (ByVal hPalette As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteColors As PALETTEENTRY) As Long
Public Declare Function Arc Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CancelDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function Chord Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Public Declare Function CloseMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CopyMetaFile Lib "gdi32" Alias "CopyMetaFileA" (ByVal hMF As Long, ByVal lpFileName As String) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateDiscardableBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As Long
Public Declare Function CreateDIBitmap Lib "gdi32" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function CreateDIBPatternBrush Lib "gdi32" (ByVal hPackedDIB As Long, ByVal wUsage As Long) As Long
Public Declare Function CreateDIBPatternBrushPt Lib "gdi32" (lpPackedDIB As Any, ByVal iUsage As Long) As Long
Public Declare Function CreateEllipticRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Public Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Public Declare Function CreateIC Lib "gdi32" Alias "CreateICA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As Long
Public Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
Public Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Public Declare Function CreatePolyPolygonRgn Lib "gdi32" (lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function CreateScalableFontResource Lib "gdi32" Alias "CreateScalableFontResourceA" (ByVal fHidden As Long, ByVal lpszResourceFile As String, ByVal lpszFontFile As String, ByVal lpszCurrentPath As String) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
Public Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, lpDevMode As DEVMODE) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function EqualRgn Lib "gdi32" (ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long) As Long
Public Declare Function Escape Lib "gdi32" (ByVal hdc As Long, ByVal nEscape As Long, ByVal nCount As Long, ByVal lpInData As String, lpOutData As Any) As Long
Public Declare Function ExtEscape Lib "gdi32" (ByVal hdc As Long, ByVal nEscape As Long, ByVal cbInput As Long, ByVal lpszInData As String, ByVal cbOutput As Long, ByVal lpszOutData As String) As Long
Public Declare Function DrawEscape Lib "gdi32" (ByVal hdc As Long, ByVal nEscape As Long, ByVal cbInput As Long, ByVal lpszInData As String) As Long
Public Declare Function ExcludeClipRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function ExtCreateRegion Lib "gdi32" (lpXform As xform, ByVal nCount As Long, lpRgnData As RgnData) As Long
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Public Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Public Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetROP2 Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetAspectRatioFilterEx Lib "gdi32" (ByVal hdc As Long, lpAspectRatio As Size) As Long
Public Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetBkMode Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Public Declare Function GetBitmapDimensionEx Lib "gdi32" (ByVal hBitmap As Long, lpDimension As Size) As Long
Public Declare Function GetBoundsRect Lib "gdi32" (ByVal hdc As Long, lprcBounds As RECT, ByVal Flags As Long) As Long
Public Declare Function GetCharWidth Lib "gdi32" Alias "GetCharWidthA" (ByVal hdc As Long, ByVal wFirstChar As Long, ByVal wLastChar As Long, lpBuffer As Long) As Long
Public Declare Function GetCharWidth32 Lib "gdi32" Alias "GetCharWidth32A" (ByVal hdc As Long, ByVal iFirstChar As Long, ByVal iLastChar As Long, lpBuffer As Long) As Long
Public Declare Function GetCharWidthFloat Lib "gdi32" Alias "GetCharWidthFloatA" (ByVal hdc As Long, ByVal iFirstChar As Long, ByVal iLastChar As Long, pxBuffer As Double) As Long
Public Declare Function GetCharABCWidths Lib "gdi32" Alias "GetCharABCWidthsA" (ByVal hdc As Long, ByVal uFirstChar As Long, ByVal uLastChar As Long, lpabc As ABC) As Long
Public Declare Function GetCharABCWidthsFloat Lib "gdi32" Alias "GetCharABCWidthsFloatA" (ByVal hdc As Long, ByVal iFirstChar As Long, ByVal iLastChar As Long, lpABCF As ABCFLOAT) As Long
Public Declare Function GetClipBox Lib "gdi32" (ByVal hdc As Long, lpRect As RECT) As Long
Public Declare Function GetClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Public Declare Function GetMetaRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Public Declare Function GetCurrentObject Lib "gdi32" (ByVal hdc As Long, ByVal uObjectType As Long) As Long
Public Declare Function GetCurrentPositionEx Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function GetFontData Lib "gdi32" Alias "GetFontDataA" (ByVal hdc As Long, ByVal dwTable As Long, ByVal dwOffset As Long, lpvBuffer As Any, ByVal cbData As Long) As Long
Public Declare Function GetGlyphOutline Lib "gdi32" Alias "GetGlyphOutlineA" (ByVal hdc As Long, ByVal uChar As Long, ByVal fuFormat As Long, lpgm As GLYPHMETRICS, ByVal cbBuffer As Long, lpBuffer As Any, lpmat2 As MAT2) As Long
Public Declare Function GetGraphicsMode Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetMapMode Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetMetaFileBitsEx Lib "gdi32" (ByVal hMF As Long, ByVal nSize As Long, lpvData As Any) As Long
Public Declare Function GetMetaFile Lib "gdi32" Alias "GetMetaFileA" (ByVal lpFileName As String) As Long
Public Declare Function GetNearestColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function GetNearestPaletteIndex Lib "gdi32" (ByVal hPalette As Long, ByVal crColor As Long) As Long
Public Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Public Declare Function GetOutlineTextMetrics Lib "gdi32" Alias "GetOutlineTextMetricsA" (ByVal hdc As Long, ByVal cbData As Long, lpotm As OUTLINETEXTMETRIC) As Long
Public Declare Function GetPaletteEntries Lib "gdi32" (ByVal hPalette As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetPolyFillMode Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetRasterizerCaps Lib "gdi32" (lpraststat As RASTERIZER_STATUS, ByVal cb As Long) As Long
Public Declare Function GetRegionData Lib "gdi32" Alias "GetRegionDataA" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As RgnData) As Long
Public Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Public Declare Function GetStretchBltMode Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Public Declare Function GetSystemPaletteUse Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetTextCharacterExtra Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetTextAlign Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetTextExtentPoint Lib "gdi32" Alias "GetTextExtentPointA" (ByVal hdc As Long, ByVal lpszString As String, ByVal cbString As Long, lpSize As Size) As Long
Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Public Declare Function GetTextExtentExPoint Lib "gdi32" Alias "GetTextExtentExPointA" (ByVal hdc As Long, ByVal lpszStr As String, ByVal cchString As Long, ByVal nMaxExtent As Long, lpnFit As Long, alpDx As Long, lpSize As Size) As Long
Public Declare Function GetViewportExtEx Lib "gdi32" (ByVal hdc As Long, lpSize As Size) As Long
Public Declare Function GetViewportOrgEx Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetWindowExtEx Lib "gdi32" (ByVal hdc As Long, lpSize As Size) As Long
Public Declare Function GetWindowOrgEx Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI) As Long
Public Declare Function IntersectClipRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function InvertRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MaskBlt Lib "gdi32" (ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal hbmMask As Long, ByVal xMask As Long, ByVal yMask As Long, ByVal dwRop As Long) As Long
Public Declare Function PlgBlt Lib "gdi32" (ByVal hdcDest As Long, lpPoint As POINTAPI, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hbmMask As Long, ByVal xMask As Long, ByVal yMask As Long) As Long
Public Declare Function OffsetClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function Pie Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Public Declare Function PlayMetaFile Lib "gdi32" (ByVal hdc As Long, ByVal hMF As Long) As Long
Public Declare Function PaintRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Public Declare Function PolyPolygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long) As Long
Public Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function PtVisible Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function RectVisible Lib "gdi32" (ByVal hdc As Long, lpRect As RECT) As Long
Public Declare Function RectInRegion Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal nSavedDC As Long) As Long
Public Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hdc As Long, lpInitData As DEVMODE) As Long
Public Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long
Public Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function ResizePalette Lib "gdi32" (ByVal hPalette As Long, ByVal nNumEntries As Long) As Long
Public Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Public Declare Function ExtSelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal fnMode As Long) As Long
Public Declare Function SetMetaRgn Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Public Declare Function SetBoundsRect Lib "gdi32" (ByVal hdc As Long, lprcBounds As RECT, ByVal Flags As Long) As Long
Public Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function SetMapperFlags Lib "gdi32" (ByVal hdc As Long, ByVal dwFlag As Long) As Long
Public Declare Function SetGraphicsMode Lib "gdi32" (ByVal hdc As Long, ByVal iMode As Long) As Long
Public Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Public Declare Function SetMetaFileBitsEx Lib "gdi32" (ByVal nSize As Long, lpData As Byte) As Long
Public Declare Function SetPaletteEntries Lib "gdi32" (ByVal hPalette As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function SetPolyFillMode Lib "gdi32" (ByVal hdc As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetRectRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Public Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Public Declare Function SetSystemPaletteUse Lib "gdi32" (ByVal hdc As Long, ByVal wUsage As Long) As Long
Public Declare Function SetTextCharacterExtra Lib "gdi32" Alias "SetTextCharacterExtraA" (ByVal hdc As Long, ByVal nCharExtra As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetTextAlign Lib "gdi32" (ByVal hdc As Long, ByVal wFlags As Long) As Long
Public Declare Function SetTextJustification Lib "gdi32" (ByVal hdc As Long, ByVal nBreakExtra As Long, ByVal nBreakCount As Long) As Long
Public Declare Function UpdateColors Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function PlayMetaFileRecord Lib "gdi32" (ByVal hdc As Long, lpHandletable As HANDLETABLE, lpMetaRecord As METARECORD, ByVal nHandles As Long) As Long
Public Declare Function CloseEnhMetaFile Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As Long, ByVal lpszFile As String) As Long
Public Declare Function CreateEnhMetaFile Lib "gdi32" Alias "CreateEnhMetaFileA" (ByVal hdcRef As Long, ByVal lpFileName As String, lpRect As RECT, ByVal lpDescription As String) As Long
Public Declare Function DeleteEnhMetaFile Lib "gdi32" (ByVal hemf As Long) As Long
Public Declare Function EnumEnhMetaFile Lib "gdi32" (ByVal hdc As Long, ByVal hemf As Long, ByVal lpEnhMetaFunc As Long, lpData As Any, lpRect As RECT) As Long
Public Declare Function GetEnhMetaFile Lib "gdi32" Alias "GetEnhMetaFileA" (ByVal lpszMetaFile As String) As Long
Public Declare Function GetEnhMetaFileBits Lib "gdi32" (ByVal hemf As Long, ByVal cbBuffer As Long, lpbBuffer As Byte) As Long
Public Declare Function GetEnhMetaFileDescription Lib "gdi32" Alias "GetEnhMetaFileDescriptionA" (ByVal hemf As Long, ByVal cchBuffer As Long, ByVal lpszDescription As String) As Long
Public Declare Function GetEnhMetaFileHeader Lib "gdi32" (ByVal hemf As Long, ByVal cbBuffer As Long, lpemh As ENHMETAHEADER) As Long
Public Declare Function GetEnhMetaFilePaletteEntries Lib "gdi32" (ByVal hemf As Long, ByVal cEntries As Long, lppe As PALETTEENTRY) As Long
Public Declare Function GetWinMetaFileBits Lib "gdi32" (ByVal hemf As Long, ByVal cbBuffer As Long, lpbBuffer As Byte, ByVal fnMapMode As Long, ByVal hdcRef As Long) As Long
Public Declare Function PlayEnhMetaFile Lib "gdi32" (ByVal hdc As Long, ByVal hemf As Long, lpRect As RECT) As Long
Public Declare Function PlayEnhMetaFileRecord Lib "gdi32" (ByVal hdc As Long, lpHandletable As HANDLETABLE, lpEnhMetaRecord As ENHMETARECORD, ByVal nHandles As Long) As Long
Public Declare Function SetEnhMetaFileBits Lib "gdi32" (ByVal cbBuffer As Long, lpData As Byte) As Long
Public Declare Function SetWinMetaFileBits Lib "gdi32" (ByVal cbBuffer As Long, lpbBuffer As Byte, ByVal hdcRef As Long, lpmfp As METAFILEPICT) As Long
Public Declare Function GdiComment Lib "gdi32" (ByVal hdc As Long, ByVal cbSize As Long, lpData As Byte) As Long
Public Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Public Declare Function AngleArc Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dwRadius As Long, ByVal eStartAngle As Double, ByVal eSweepAngle As Double) As Long
Public Declare Function PolyPolyline Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, lpdwPolyPoints As Long, ByVal cCount As Long) As Long
Public Declare Function GetWorldTransform Lib "gdi32" (ByVal hdc As Long, lpXform As xform) As Long
Public Declare Function SetWorldTransform Lib "gdi32" (ByVal hdc As Long, lpXform As xform) As Long
Public Declare Function ModifyWorldTransform Lib "gdi32" (ByVal hdc As Long, lpXform As xform, ByVal iMode As Long) As Long
Public Declare Function CombineTransform Lib "gdi32" (lpxformResult As xform, lpxform1 As xform, lpxform2 As xform) As Long
Public Declare Function SetColorAdjustment Lib "gdi32" (ByVal hdc As Long, lpca As ColorAdjustment) As Long
Public Declare Function GetColorAdjustment Lib "gdi32" (ByVal hdc As Long, lpca As ColorAdjustment) As Long
Public Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function StartDoc Lib "gdi32" Alias "StartDocA" (ByVal hdc As Long, lpdi As DOCINFO) As Long
Public Declare Function StartPage Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function EndPage Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function EndDoc Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function AbortDoc Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function AbortPath Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function ArcTo Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Public Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CloseFigure Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function FillPath Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function FlattenPath Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetPath Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, lpTypes As Byte, ByVal nSize As Long) As Long
Public Declare Function PathToRegion Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function PolyDraw Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, lpbTypes As Byte, ByVal cCount As Long) As Long
Public Declare Function SelectClipPath Lib "gdi32" (ByVal hdc As Long, ByVal iMode As Long) As Long
Public Declare Function SetArcDirection Lib "gdi32" (ByVal hdc As Long, ByVal ArcDirection As Long) As Long
Public Declare Function SetMiterLimit Lib "gdi32" (ByVal hdc As Long, ByVal eNewLimit As Double, peOldLimit As Double) As Long
Public Declare Function StrokeAndFillPath Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function StrokePath Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function WidenPath Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function ExtCreatePen Lib "gdi32" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, lplb As LOGBRUSH, ByVal dwStyleCount As Long, lpStyle As Long) As Long
Public Declare Function GetMiterLimit Lib "gdi32" (ByVal hdc As Long, peLimit As Double) As Long
Public Declare Function GetArcDirection Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Public Declare Function PolyTextOut Lib "gdi32" Alias "PolyTextOutA" (ByVal hdc As Long, pptxt As POLYTEXT, cStrings As Long) As Long
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function DPtoLP Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function LPtoDP Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function PolyBezier Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, ByVal cPoints As Long) As Long
Public Declare Function PolyBezierTo Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, ByVal cCount As Long) As Long
Public Declare Function PolylineTo Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, ByVal cCount As Long) As Long
Public Declare Function SetViewportExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
Public Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Public Declare Function SetWindowOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Public Declare Function SetWindowExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
Public Declare Function OffsetViewportOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Public Declare Function OffsetWindowOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Public Declare Function ScaleWindowExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nXnum As Long, ByVal nXdenom As Long, ByVal nYnum As Long, ByVal nYdenom As Long, lpSize As Size) As Long
Public Declare Function ScaleViewportExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nXnum As Long, ByVal nXdenom As Long, ByVal nYnum As Long, ByVal nYdenom As Long, lpSize As Size) As Long
Public Declare Function SetBitmapDimensionEx Lib "gdi32" (ByVal hbm As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
Public Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
Public Declare Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal hdc As Long, ByVal nCount As Long, ByVal lpFacename As String) As Long
Public Declare Function GetKerningPairs Lib "gdi32" Alias "GetKerningPairsA" (ByVal hdc As Long, ByVal cPairs As Long, lpkrnpair As KERNINGPAIR) As Long
Public Declare Function GetDCOrgEx Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI) As Long
Public Declare Function UnrealizeObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GdiFlush Lib "gdi32" () As Long
Public Declare Function GdiSetBatchLimit Lib "gdi32" (ByVal dwLimit As Long) As Long
Public Declare Function GdiGetBatchLimit Lib "gdi32" () As Long
Public Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hdc As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumFontFamiliesEx Lib "gdi32" Alias "EnumFontFamiliesExA" (ByVal hdc As Long, lpLogFont As LOGFONT, ByVal lpEnumFontProc As Long, ByVal lParam As Long, ByVal dw As Long) As Long
Public Declare Function GetTextCharset Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetTextCharsetInfo Lib "gdi32" (ByVal hdc As Long, lpSig As FONTSIGNATURE, ByVal dwFlags As Long) As Long
Public Declare Function TranslateCharsetInfo Lib "gdi32" (lpSrc As Long, lpcs As CHARSETINFO, ByVal dwFlags As Long) As Long
Public Declare Function GetFontLanguageInfo Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetCharacterPlacement Lib "gdi32" Alias " GetCharacterPlacementA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n1 As Long, ByVal n2 As Long, lpGcpResults As GCP_RESULTS, ByVal dw As Long) As Long
Public Declare Function SetICMMode Lib "gdi32" (ByVal hdc As Long, ByVal n As Long) As Long
Public Declare Function CheckColorsInGamut Lib "gdi32" (ByVal hdc As Long, lpv As Any, lpv2 As Any, ByVal dw As Long) As Long
Public Declare Function GetLogColorSpace Lib "gdi32" Alias "GetLogColorSpaceA" (ByVal hcolorspace As Long, ByVal lplogcolorspace As LOGCOLORSPACE, ByVal dw As Long) As Long
Public Declare Function GetColorSpace Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateColorSpace Lib "gdi32" Alias "CreateColorSpaceA" (lplogcolorspace As LOGCOLORSPACE) As Long
Public Declare Function SetColorSpace Lib "gdi32" (ByVal hdc As Long, ByVal hcolorspace As Long) As Long
Public Declare Function DeleteColorSpace Lib "gdi32" (ByVal hcolorspace As Long) As Long
Public Declare Function GetICMProfile Lib "gdi32" Alias "GetICMProfileA" (ByVal hdc As Long, ByVal dw As Long, ByVal lpStr As String) As Long
Public Declare Function SetICMProfile Lib "gdi32" Alias "SetICMProfileA" (ByVal hdc As Long, ByVal lpStr As String) As Long
Public Declare Function GetDeviceGammaRamp Lib "gdi32" (ByVal hdc As Long, lpv As Any) As Long
Public Declare Function SetDeviceGammaRamp Lib "gdi32" (ByVal hdc As Long, lpv As Any) As Long
Public Declare Function ColorMatchToTarget Lib "gdi32" (ByVal hdc As Long, ByVal hdc2 As Long, ByVal dw As Long) As Long
Public Declare Function EnumICMProfiles Lib "gdi32" Alias "EnumICMProfilesA" (ByVal hdc As Long, ByVal icmEnumProc As Long, ByVal lParam As Long) As Long
Public Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal Flags As Long) As Long
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal Flags As Long) As Long
Public Declare Function UnloadKeyboardLayout Lib "user32" (ByVal hkl As Long) As Long
Public Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Public Declare Function GetThreadDesktop Lib "user32" (ByVal dwThread As Long) As Long
Public Declare Function GetProcessWindowStation Lib "user32" () As Long
Public Declare Function SetUserObjectSecurity Lib "user32" (ByVal hObj As Long, pSIRequested As Long, pSd As SECURITY_DESCRIPTOR) As Long
Public Declare Function GetUserObjectSecurity Lib "user32" (ByVal hObj As Long, pSIRequested As Long, pSd As SECURITY_DESCRIPTOR, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
' Message Function Templates

Public Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Public Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Public Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long) As Long
Public Declare Function ExitWindows Lib "user32" (ByVal dwReserved As Long, ByVal uReturnCode As Long) As Long
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long
Public Declare Function GetMessagePos Lib "user32" () As Long
Public Declare Function GetMessageTime Lib "user32" () As Long
Public Declare Function GetMessageExtraInfo Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Public Declare Function SendNotifyMessage Lib "user32" Alias "SendNotifyMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageCallback Lib "user32" Alias "SendMessageCallbackA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal lpResultCallBack As Long, ByVal dwData As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function PostThreadMessage Lib "user32" Alias "PostThreadMessageA" (ByVal idThread As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Public Declare Function ReplyMessage Lib "user32" (ByVal lReply As Long) As Long
Public Declare Function WaitMessage Lib "user32" () As Long
Public Declare Function WaitForInputIdle Lib "user32" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Public Declare Function InSendMessage Lib "user32" () As Long
Public Declare Function GetDoubleClickTime Lib "user32" () As Long
Public Declare Function SetDoubleClickTime Lib "user32" (ByVal wCount As Long) As Long
Public Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Public Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Public Declare Function GetClassInfo Lib "user32" Alias "GetClassInfoA" (ByVal hInstance As Long, ByVal lpClassName As String, lpWndClass As WNDCLASS) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function IsChild Lib "user32" (ByVal hwndParent As Long, ByVal hWnd As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Public Declare Function ShowOwnedPopups Lib "user32" (ByVal hWnd As Long, ByVal fShow As Long) As Long
Public Declare Function OpenIcon Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function SetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function BeginDeferWindowPos Lib "user32" (ByVal nNumWindows As Long) As Long
Public Declare Function DeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long, ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function EndDeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function AnyPopup Lib "user32" () As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function EndDialog Lib "user32" (ByVal hDlg As Long, ByVal nResult As Long) As Long
Public Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Public Declare Function SetDlgItemInt Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wValue As Long, ByVal bSigned As Long) As Long
Public Declare Function GetDlgItemInt Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpTranslated As Long, ByVal bSigned As Long) As Long
Public Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String) As Long
Public Declare Function GetDlgItemText Lib "user32" Alias "GetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Public Declare Function CheckDlgButton Lib "user32" Alias "CheckDLGButtonA" (ByVal hDlg As Long, ByVal nIDButton As Long, ByVal wCheck As Long) As Long
Public Declare Function CheckRadioButton Lib "user32" Alias "CheckRadioButtonA" (ByVal hDlg As Long, ByVal nIDFirstButton As Long, ByVal nIDLastButton As Long, ByVal nIDCheckButton As Long) As Long
Public Declare Function IsDlgButtonChecked Lib "user32" (ByVal hDlg As Long, ByVal nIDButton As Long) As Long
Public Declare Function SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetNextDlgGroupItem Lib "user32" (ByVal hDlg As Long, ByVal hCtl As Long, ByVal bPrevious As Long) As Long
Public Declare Function GetNextDlgTabItem Lib "user32" (ByVal hDlg As Long, ByVal hCtl As Long, ByVal bPrevious As Long) As Long
Public Declare Function GetDlgCtrlID Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetDialogBaseUnits Lib "user32" () As Long
Public Declare Function DefDlgProc Lib "user32" Alias "DefDlgProcA" (ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CallMsgFilter Lib "user32" Alias "CallMsgFilterA" (lpMsg As Msg, ByVal nCode As Long) As Long
' Clipboard Manager Functions

Public Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function GetClipboardOwner Lib "user32" () As Long
Public Declare Function SetClipboardViewer Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetClipboardViewer Lib "user32" () As Long
Public Declare Function ChangeClipboardChain Lib "user32" (ByVal hWnd As Long, ByVal hWndNext As Long) As Long
Public Declare Function SetClipboardData Lib "user32" Alias "SetClipboardDataA" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function GetClipboardData Lib "user32" Alias "GetClipboardDataA" (ByVal wFormat As Long) As Long
Public Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Public Declare Function CountClipboardFormats Lib "user32" () As Long
Public Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
Public Declare Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Public Declare Function GetPriorityClipboardFormat Lib "user32" (lpPriorityList As Long, ByVal nCount As Long) As Long
Public Declare Function GetOpenClipboardWindow Lib "user32" () As Long
Public Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Public Declare Function OemToChar Lib "user32" Alias "OemToCharA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Public Declare Function CharToOemBuff Lib "user32" Alias "CharToOemBuffA" (ByVal lpszSrc As String, ByVal lpszDst As String, ByVal cchDstLength As Long) As Long
Public Declare Function OemToCharBuff Lib "user32" Alias "OemToCharBuffA" (ByVal lpszSrc As String, ByVal lpszDst As String, ByVal cchDstLength As Long) As Long
Public Declare Function CharUpper Lib "user32" Alias "CharUpperA" (ByVal lpsz As String) As String
Public Declare Function CharUpperBuff Lib "user32" Alias "CharUpperBuffA" (ByVal lpsz As String, ByVal cchLength As Long) As Long
Public Declare Function CharLower Lib "user32" Alias "CharLowerA" (ByVal lpsz As String) As String
Public Declare Function CharLowerBuff Lib "user32" Alias "CharLowerBuffA" (ByVal lpsz As String, ByVal cchLength As Long) As Long
Public Declare Function CharNext Lib "user32" Alias "CharNextA" (ByVal lpsz As String) As String
Public Declare Function CharPrev Lib "user32" Alias "CharPrevA" (ByVal lpszStart As String, ByVal lpszCurrent As String) As String
' Language dependent Routines

Public Declare Function IsCharAlpha Lib "user32" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long
Public Declare Function IsCharAlphaNumeric Lib "user32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
Public Declare Function IsCharUpper Lib "user32" Alias "IsCharUpperA" (ByVal cChar As Byte) As Long
Public Declare Function IsCharLower Lib "user32" Alias "IsCharLowerA" (ByVal cChar As Byte) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
' Keyboard Information Routines

Public Declare Function GetKBCodePage Lib "user32" () As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Public Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Public Declare Function GetKeyboardType Lib "user32" (ByVal nTypeFlag As Long) As Long
Public Declare Function GetKeyNameText Lib "user32" Alias "GetKeyNameTextA" (ByVal lParam As Long, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpbKeyState As Byte, lpwTransKey As Long, ByVal fuState As Long) As Long
Public Declare Function ToUnicode Lib "user32" (ByVal wVirtKey As Long, ByVal wScanCode As Long, lpKeyState As Byte, ByVal pwszBuff As String, ByVal cchBuff As Long, ByVal wFlags As Long) As Long
Public Declare Function OemKeyScan Lib "user32" (ByVal wOemChar As Long) As Long
Public Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Public Declare Function GetInputState Lib "user32" () As Long
Public Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function MsgWaitForMultipleObjects Lib "user32" (ByVal nCount As Long, pHandles As Long, ByVal fWaitAll As Long, ByVal dwMilliseconds As Long, ByVal dwWakeMask As Long) As Long
' Windows Functions

Public Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function IsWindowUnicode Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function LoadAccelerators Lib "user32" Alias "LoadAcceleratorsA" (ByVal hInstance As Long, ByVal lpTableName As String) As Long
Public Declare Function CreateAcceleratorTable Lib "user32" Alias "CreateAcceleratorTableA" (lpaccl As ACCEL, ByVal cEntries As Long) As Long
Public Declare Function DestroyAcceleratorTable Lib "user32" (ByVal haccel As Long) As Long
Public Declare Function CopyAcceleratorTable Lib "user32" Alias "CopyAcceleratorTableA" (ByVal hAccelSrc As Long, lpAccelDst As ACCEL, ByVal cAccelEntries As Long) As Long
Public Declare Function TranslateAccelerator Lib "user32" Alias "TranslateAcceleratorA" (ByVal hWnd As Long, ByVal hAccTable As Long, lpMsg As Msg) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function LoadMenu Lib "user32" Alias "LoadMenuA" (ByVal hInstance As Long, ByVal lpString As String) As Long
Public Declare Function LoadMenuIndirect Lib "user32" Alias "LoadMenuIndirectA" (ByVal lpMenuTemplate As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long) As Long
Public Declare Function HiliteMenuItem Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal wIDHiliteItem As Long, ByVal wHilite As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function CreateMenu Lib "user32" () As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function CheckMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
Public Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Public Declare Function GetMenuCheckMarkDimensions Lib "user32" () As Long
Public Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprc As RECT) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function TabbedTextOut Lib "user32" Alias "TabbedTextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long, ByVal nTabPositions As Long, lpnTabStopPositions As Long, ByVal nTabOrigin As Long) As Long
Public Declare Function GetTabbedTextExtent Lib "user32" Alias "GetTabbedTextExtentA" (ByVal hdc As Long, ByVal lpString As String, ByVal nCount As Long, ByVal nTabPositions As Long, lpnTabStopPositions As Long) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function WindowFromDC Lib "user32" (ByVal hdc As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetDCEx Lib "user32" (ByVal hWnd As Long, ByVal hrgnclip As Long, ByVal fdwOptions As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Public Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Public Declare Function GetUpdateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Declare Function GetUpdateRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal fErase As Long) As Long
Public Declare Function ExcludeUpdateRgn Lib "user32" (ByVal hdc As Long, ByVal hWnd As Long) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Declare Function ValidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function InvalidateRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bErase As Long) As Long
Public Declare Function ValidateRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long) As Long
Public Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

