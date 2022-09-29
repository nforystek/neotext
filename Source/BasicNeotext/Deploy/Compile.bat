

"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make "C:\Development\Neotext\BasicNeotext\Projects\BasicNeotext_DLL.vbp" /d VBIDE=0:modRegistry=-1
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make "C:\Development\Neotext\BasicNeotext\Projects\BasicNeotext_EXE.vbp" /d VBIDE=0:modRegistry=-1
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make "C:\Development\Neotext\BasicNeotext\Projects\BasicNotable.vbp" /d VBIDE=0

"C:\Program Files\NSIS\makensis.exe" /V1 "C:\Development\Neotext\BasicNeotext\Deploy\BasicNeotext v3.0.0.nsi"

"C:\Program Files\Microsoft Visual Studio\VB98\Uninstall.exe" /S

"C:\Development\Neotext\BasicNeotext\Deploy\BasicNeotext v3.0.0.exe" /S

