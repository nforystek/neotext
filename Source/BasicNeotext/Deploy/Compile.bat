

"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make "C:\Development\Neotext\BasicNeotext\Projects\BasicNeotext_DLL.vbp" /d VBIDE=0:modRegistry=-1
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make "C:\Development\Neotext\BasicNeotext\Projects\BasicNeotext_EXE.vbp" /d VBIDE=0:modRegistry=-1
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make "C:\Development\Neotext\BasicNeotext\Projects\BasicNotable.vbp" /d VBIDE=0


rem "C:\Program Files\Microsoft Visual Studio\VB98\VBN.EXE" /signonly "C:\Development\Neotext\BasicNeotext\Binary\VBN.DLL"
rem "C:\Program Files\Microsoft Visual Studio\VB98\VBN.EXE" /signonly "C:\Development\Neotext\BasicNeotext\Binary\VBN.EXE"
rem "C:\Program Files\Microsoft Visual Studio\VB98\VBN.EXE" /signonly "C:\Development\Neotext\BasicNeotext\Binary\SINK.EXE"
rem "C:\Program Files\Microsoft Visual Studio\VB98\VBN.EXE" /signonly "C:\Development\Neotext\BasicNeotext\Binary\BasicService.DLL"


rem "C:\Program Files\Microsoft Visual Studio\VB98\VBN.EXE" /timeonly "C:\Development\Neotext\BasicNeotext\Binary\VBN.DLL"
rem "C:\Program Files\Microsoft Visual Studio\VB98\VBN.EXE" /timeonly "C:\Development\Neotext\BasicNeotext\Binary\VBN.EXE"
rem "C:\Program Files\Microsoft Visual Studio\VB98\VBN.EXE" /timeonly "C:\Development\Neotext\BasicNeotext\Binary\SINK.EXE"
rem "C:\Program Files\Microsoft Visual Studio\VB98\VBN.EXE" /timeonly "C:\Development\Neotext\BasicNeotext\Binary\BasicService.DLL"

"C:\Program Files\NSIS\makensis.exe" /V1 "C:\Development\Neotext\BasicNeotext\Deploy\BasicNeotext v3.0.0.nsi"

rem "C:\Program Files\Microsoft Visual Studio\VB98\VBN.EXE" /signonly  "C:\Development\Neotext\BasicNeotext\Deploy\BasicNeotext v3.0.0.exe"
rem "C:\Program Files\Microsoft Visual Studio\VB98\VBN.EXE" /timeonly  "C:\Development\Neotext\BasicNeotext\Deploy\BasicNeotext v3.0.0.exe"

"C:\Program Files\Microsoft Visual Studio\VB98\Uninstall.exe" /S

"C:\Development\Neotext\BasicNeotext\Deploy\BasicNeotext v3.0.0.exe" /S

