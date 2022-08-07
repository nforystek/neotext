
"C:\Development\Neotext\Max-FTP\Binary\MaxUtility.exe" stop
"C:\Development\Neotext\RemindMe\Binary\Utility.exe" stop
"C:\Program Files\Ident Protocol Service\Reload.exe" stop
"C:\Program Files\CrayonStill\CrayonStall.exe" stop
"C:\Development\Neotext\CrayonStill\Binary\CrayonStall.exe" stop

regsvr32.exe /u /s "C:\WINDOWS\system32\NTControls22.ocx"
regsvr32.exe /u /s "C:\WINDOWS\system32\NTAdvFTP61.dll"
regsvr32.exe /u /s "C:\WINDOWS\system32\NTService20.dll"
regsvr32.exe /u /s "C:\WINDOWS\system32\NTSchedule20.dll"
regsvr32.exe /u /s "C:\WINDOWS\system32\NTSound20.dll"
regsvr32.exe /u /s "C:\WINDOWS\system32\NTShell22.dll"
regsvr32.exe /u /s "C:\WINDOWS\system32\NTNodes10.dll"
regsvr32.exe /u /s "C:\WINDOWS\system32\NTPopup21.dll"
regsvr32.exe /u /s "C:\WINDOWS\system32\NTCipher10.dll"


erase "C:\WINDOWS\system32\NTControls22.ocx"
erase "C:\WINDOWS\system32\NTAdvFTP61.dll"
erase "C:\WINDOWS\system32\NTService20.dll"
erase "C:\WINDOWS\system32\NTSchedule20.dll"
erase "C:\WINDOWS\system32\NTSound20.dll"
erase "C:\WINDOWS\system32\NTNodes10.dll"
erase "C:\WINDOWS\system32\NTShell22.dll"
erase "C:\WINDOWS\system32\NTPopup21.dll"
erase "C:\WINDOWS\system32\NTCipher10.dll"

regsvr32.exe /u /s "C:\Development\Neotext\Common\Binary\NTControls22.ocx"
regsvr32.exe /u /s "C:\Development\Neotext\Common\Binary\NTAdvFTP61.dll"
regsvr32.exe /u /s "C:\Development\Neotext\Common\Binary\NTService20.dll"
regsvr32.exe /u /s "C:\Development\Neotext\Common\Binary\NTSchedule20.dll"
regsvr32.exe /u /s "C:\Development\Neotext\Common\Binary\NTSound20.dll"
regsvr32.exe /u /s "C:\Development\Neotext\Common\Binary\NTShell22.dll"
regsvr32.exe /u /s "C:\Development\Neotext\Common\Binary\NTNodes10.dll"
regsvr32.exe /u /s "C:\Development\Neotext\Common\Binary\NTPopup21.dll"
regsvr32.exe /u /s "C:\Development\Neotext\Common\Binary\NTCipher10.dll"

"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make "C:\Development\Neotext\Common\Projects\NTCipher10.vbp" /d VBIDE=0
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make "C:\Development\Neotext\Common\Projects\NTPopup21.vbp" /d VBIDE=0
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make "C:\Development\Neotext\Common\Projects\NTShell22.vbp" /d VBIDE=0
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make "C:\Development\Neotext\Common\Projects\NTSound20.vbp" /d VBIDE=0
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make "C:\Development\Neotext\Common\Projects\NTSchedule20.vbp" /d VBIDE=0
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make "C:\Development\Neotext\Common\Projects\NTService20.vbp" /d VBIDE=0
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make "C:\Development\Neotext\Common\Projects\NTNodes10.vbp" /d VBIDE=0
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make "C:\Development\Neotext\Common\Projects\NTAdvFTP61.vbp" /d VBIDE=0
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make "C:\Development\Neotext\Common\Projects\NTControls22.vbp" /d VBIDE=0

"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make "C:\Development\Neotext\Max-FTP\Projects\MaxIDE.vbp" /d VBIDE=0
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make "C:\Development\Neotext\Max-FTP\Projects\MaxFTP.vbp" /d VBIDE=0
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make "C:\Development\Neotext\Max-FTP\Projects\MaxService.vbp" /d VBIDE=0
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make "C:\Development\Neotext\Max-FTP\Projects\MaxUtility.vbp" /d VBIDE=0


cd \Development\Neotext\InstallerStar\Binary
C:\Development\Neotext\Max-FTP\Binary\MaxUtility.exe /setupreset
Wizard.exe /compile Max-FTP

"C:\Development\Neotext\Max-FTP\Deploy\Max-FTP v6.1.0.exe" /Q