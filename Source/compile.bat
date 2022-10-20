@echo off


echo Stopping applications and services using shared libaries...
"C:\Development\Neotext\Max-FTP\Binary\MaxUtility.exe" stop
"C:\Development\Neotext\RemindMe\Binary\Utility.exe" stop
"C:\Program Files\IdentAuth\Reload.exe" stop
net stop Securities
rem "C:\Program Files\CrayonStill\CrayonStall.exe" stop
rem "C:\Development\Neotext\CrayonStill\Binary\CrayonStall.exe" stop


rem echo Copying system DLL's from cache to the system folder...
rem copy /y "c:\windows\system32\dllcache\comctl32.ocx" "c:\windows\system32"
rem copy /y "c:\windows\system32\dllcache\comdlg32.ocx" "c:\windows\system32"
rem copy /y "c:\windows\system32\dllcache\dx7vb.dll" "c:\windows\system32"
rem copy /y "c:\windows\system32\dllcache\dx8vb.dll" "c:\windows\system32"
rem copy /y "c:\windows\system32\dllcache\msado15.dll" "c:\windows\system32"
rem copy /y "c:\windows\system32\dllcache\mscomct2.ocx" "c:\windows\system32"
rem copy /y "c:\windows\system32\dllcache\mscomctl.ocx" "c:\windows\system32"
rem copy /y "c:\windows\system32\dllcache\msjro.dll" "c:\windows\system32"
rem copy /y "c:\windows\system32\dllcache\msscript.ocx" "c:\windows\system32"
rem copy /y "c:\windows\system32\dllcache\msvbvm50.dll" "c:\windows\system32"
rem copy /y "c:\windows\system32\dllcache\msvbvm60.dll" "c:\windows\system32"
rem copy /y "c:\windows\system32\dllcache\mswinsck.ocx" "c:\windows\system32"
rem copy /y "c:\windows\system32\dllcache\msxml2.dll" "c:\windows\system32"
rem copy /y "c:\windows\system32\dllcache\msxml6.dll" "c:\windows\system32"
rem copy /y "c:\windows\system32\dllcache\msxml.dll" "c:\windows\system32"
rem copy /y "c:\windows\system32\dllcache\ntsvc.ocx" "c:\windows\system32"
rem copy /y "c:\windows\system32\dllcache\richtx32.ocx" "c:\windows\system32"
rem copy /y "c:\windows\system32\dllcache\scrrun.dll" "c:\windows\system32"
rem copy /y "c:\windows\system32\dllcache\shdocvw.dll" "c:\windows\system32"


rem echo Registering Microsoft shared ActiveX DLL's in system32...
rem regsvr32 /s "c:\windows\system32\comctl32.ocx"
rem regsvr32 /s "c:\windows\system32\comdlg32.ocx"
rem regsvr32 /s "c:\windows\system32\dx7vb.dll"
rem regsvr32 /s "c:\windows\system32\dx8vb.dll"
rem regsvr32 /s "c:\windows\system32\msado15.dll"
rem regsvr32 /s "c:\windows\system32\mscomct2.ocx"
rem regsvr32 /s "c:\windows\system32\mscomctl.ocx"
rem regsvr32 /s "c:\windows\system32\msjro.dll"
rem regsvr32 /s "c:\windows\system32\msscript.ocx"
rem regsvr32 /s "c:\windows\system32\msvbvm50.dll"
rem regsvr32 /s "c:\windows\system32\msvbvm60.dll"
rem regsvr32 /s "c:\windows\system32\mswinsck.ocx"
rem regsvr32 /s "c:\windows\system32\msxml2.dll"
rem regsvr32 /s "c:\windows\system32\msxml6.dll"
rem regsvr32 /s "c:\windows\system32\msxml.dll"
rem regsvr32 /s "c:\windows\system32\ntsvc.ocx"
rem regsvr32 /s "c:\windows\system32\richtx32.ocx"
rem regsvr32 /s "c:\windows\system32\scrrun.dll"
rem regsvr32 /s "c:\windows\system32\shdocvw.dll"


rem echo Unregistering Neotext shared libraries at the windows system path...
regsvr32.exe /u /s "C:\WINDOWS\system32\NTControls30.ocx"
regsvr32.exe /u /s "C:\WINDOWS\system32\NTControls22.ocx"
regsvr32.exe /u /s "C:\WINDOWS\system32\NTImaging10.ocx"
regsvr32.exe /u /s "C:\WINDOWS\system32\NTAdvFTP61.dll"
regsvr32.exe /u /s "C:\WINDOWS\system32\NTService20.dll"
regsvr32.exe /u /s "C:\WINDOWS\system32\NTSchedule20.dll"
regsvr32.exe /u /s "C:\WINDOWS\system32\NTNodes10.dll"
regsvr32.exe /u /s "C:\WINDOWS\system32\NTCipher10.dll"
regsvr32.exe /u /s "C:\WINDOWS\system32\NTSoSweet.dll"
regsvr32.exe /u /s "C:\WINDOWS\system32\NTSound20.dll"
regsvr32.exe /u /s "C:\WINDOWS\system32\NTShell22.dll"
regsvr32.exe /u /s "C:\WINDOWS\system32\NTPopup21.dll"
regsvr32.exe /u /s "C:\WINDOWS\system32\NTSmpFTP30.dll"
regsvr32.exe /u /s "C:\WINDOWS\system32\NTSMTP23.dll"


rem echo Erasing Neotext shared librarie files at the windows system path...
erase "C:\WINDOWS\system32\NTControls30.ocx"
erase "C:\WINDOWS\system32\NTControls22.ocx"
erase "C:\WINDOWS\system32\NTImaging10.ocx"
erase "C:\WINDOWS\system32\NTAdvFTP61.dll"
erase "C:\WINDOWS\system32\NTService20.dll"
erase "C:\WINDOWS\system32\NTSchedule20.dll"
erase "C:\WINDOWS\system32\NTNodes10.dll"
erase "C:\WINDOWS\system32\NTCipher10.dll"
erase "C:\WINDOWS\system32\NTSoSweet.dll"
erase "C:\WINDOWS\system32\NTSound20.dll"
erase "C:\WINDOWS\system32\NTShell22.dll"
erase "C:\WINDOWS\system32\NTPopup21.dll"
erase "C:\Windows\System32\NTSmpFTP30.dll"
erase "C:\Windows\System32\NTSMTP23.dll"
erase "C:\Windows\System32\MaxLandLib.dll"


rem echo Unregistering Neotext shared libraries at the compile deploy path...
regsvr32.exe /u /s "C:\Development\Neotext\Common\Binary\NTControls30.ocx"
regsvr32.exe /u /s "C:\Development\Neotext\Common\Binary\NTControls22.ocx"
regsvr32.exe /u /s "C:\Development\Neotext\Common\Binary\NTImaging10.ocx"
regsvr32.exe /u /s "C:\Development\Neotext\Common\Binary\NTAdvFTP61.dll"
regsvr32.exe /u /s "C:\Development\Neotext\Common\Binary\NTService20.dll"
regsvr32.exe /u /s "C:\Development\Neotext\Common\Binary\NTSchedule20.dll"
regsvr32.exe /u /s "C:\Development\Neotext\Common\Binary\NTNodes10.dll"
regsvr32.exe /u /s "C:\Development\Neotext\Common\Binary\NTCipher10.dll"
regsvr32.exe /u /s "C:\Development\Neotext\Common\Binary\NTSoSweet.dll"
regsvr32.exe /u /s "C:\Development\Neotext\Common\Binary\NTSound20.dll"
regsvr32.exe /u /s "C:\Development\Neotext\Common\Binary\NTShell22.dll"
regsvr32.exe /u /s "C:\Development\Neotext\Common\Binary\NTPopup21.dll"
regsvr32.exe /u /s "C:\Development\Neotext\Common\Binary\NTSmpFTP30.dll"
regsvr32.exe /u /s "C:\Development\Neotext\Common\Binary\NTSMTP23.dll"



echo Sum compiling shared libraries with debug environment set off...
echo|set /p="1. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\Common\Projects\NTCipher10.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\Common\Projects\NTCipher10.vbp"
echo|set /p="2. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\Common\Projects\NTNodes10.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\Common\Projects\NTNodes10.vbp"
echo|set /p="3. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\Common\Projects\NTSchedule20.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\Common\Projects\NTSchedule20.vbp"
echo|set /p="4. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\Common\Projects\NTPopup21.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\Common\Projects\NTPopup21.vbp"
echo|set /p="5. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\Common\Projects\NTShell22.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\Common\Projects\NTShell22.vbp"
echo|set /p="6. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\Common\Projects\NTSound20.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\Common\Projects\NTSound20.vbp"
echo|set /p="7. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\Common\Projects\NTSoSweet.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\Common\Projects\NTSoSweet.vbp"
echo|set /p="8. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\Common\Projects\NTService20.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\Common\Projects\NTService20.vbp"
echo|set /p="9. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\Common\Projects\NTAdvFTP61.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\Common\Projects\NTAdvFTP61.vbp"
echo|set /p="10. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\Common\Projects\NTControls22.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\Common\Projects\NTControls22.vbp"
echo|set /p="11. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\Common\Projects\NTControls30.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\Common\Projects\NTControls30.vbp"
echo|set /p="12. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\Common\Projects\NTImaging10.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\Common\Projects\NTImaging10.vbp"
echo Done. 


echo Sum compiling project programs with debug environment set off...
echo|set /p="1. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\Blacklawn\Projects\Blacklawn.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\Blacklawn\Projects\Blacklawn.vbp"
echo|set /p="2. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\Blacklawn\Projects\BlkLServer.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\Blacklawn\Projects\BlkLServer.vbp"
echo|set /p="3. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\To-Doster\Projects\ToDoster.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\To-Doster\Projects\ToDoster.vbp"
echo|set /p="4. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\HouseOfGlass\Projects\HouseOfGlass.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\HouseOfGlass\Projects\HouseOfGlass.vbp"
echo|set /p="5. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\Max-FTP\Projects\MaxIDE.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\Max-FTP\Projects\MaxIDE.vbp"
echo|set /p="6. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\Max-FTP\Projects\MaxFTP.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\Max-FTP\Projects\MaxFTP.vbp"
echo|set /p="7. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\Max-FTP\Projects\MaxService.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\Max-FTP\Projects\MaxService.vbp"
echo|set /p="8. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\Max-FTP\Projects\MaxUtility.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\Max-FTP\Projects\MaxUtility.vbp"
echo|set /p="9. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\MaxLand\Projects\MaxLandApp.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\MaxLand\Projects\MaxLandApp.vbp"
echo|set /p="10. 
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\RemindMe\Projects\RmdMeSrv.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\RemindMe\Projects\RmdMeSrv.vbp"
echo|set /p="11. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\RemindMe\Projects\RemindMe.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\RemindMe\Projects\RemindMe.vbp"
echo|set /p="12. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\RemindMe\Projects\Utility.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\RemindMe\Projects\Utility.vbp"
echo|set /p="13. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\Creata-Tree\Projects\CreataTree.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\Creata-Tree\Projects\CreataTree.vbp"
echo|set /p="14. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\Sequencer\Projects\SoSouiXSeq.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\Sequencer\Projects\SoSouiXSeq.vbp"
echo|set /p="15. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\IdentAuth\Projects\Reload.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\IdentAuth\Projects\Reload.vbp"
echo|set /p="16. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\IdentAuth\Projects\Ident.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\IdentAuth\Projects\Ident.vbp"
echo|set /p="17. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\InstallerStar\Projects\Wizard.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\InstallerStar\Projects\Wizard.vbp"
echo|set /p="18. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\InstallerStar\Projects\Remove.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\InstallerStar\Projects\Remove.vbp"
rem echo|set /p="16. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\Schematical\Projects\Schematical.vbp"
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\Schematical\Projects\Schematical.vbp"
rem echo|set /p="18. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\CrayonStill\Projects\CrayonStill.vbp"
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0:CrayonStill=-1 /make "C:\Development\Neotext\CrayonStill\Projects\CrayonStill.vbp"
rem echo|set /p="19. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\CrayonStill\Projects\CrayonStall.vbp"
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0:CrayonStall=-1 /make "C:\Development\Neotext\CrayonStill\Projects\CrayonStall.vbp"
rem echo|set /p="20. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\CrayonStill\Projects\CrayonStiff.vbp"
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0:CrayonStiff=-1 /make "C:\Development\Neotext\CrayonStill\Projects\CrayonStiff.vbp"
rem echo|set /p="23. "
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /open "C:\Development\Neotext\KadPatch\Projects\KadPatch.vbp"
rem "C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /d VBIDE=0 /make "C:\Development\Neotext\KadPatch\Projects\KadPatch.vbp"
echo Done. 


echo Sesetting databses to clear test information for packing...
"C:\Development\Neotext\Blacklawn\Binary\Blacklawn.exe" /setupreset
"C:\Development\Neotext\HouseOfGlass\Binary\HouseOfGlass.exe" /setupreset
"C:\Development\Neotext\Creata-Tree\Binary\CreataTree.exe" /setupreset
"C:\Development\Neotext\Max-FTP\Binary\MaxUtility.exe" /setupreset
"C:\Development\Neotext\MaxLand\Binary\MaxLandApp.exe" /setupreset
"C:\Development\Neotext\RemindMe\Binary\Utility.exe" /setupreset
rem "C:\Development\Neotext\KadPatch\Binary\KadPatch.exe" /setupreset


echo Sign and time stamping all EXE's and DLL's to be packed...

"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\Common\Projects\NTCipher10.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\Common\Projects\NTNodes10.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\Common\Projects\NTSchedule20.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\Common\Projects\NTPopup21.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\Common\Projects\NTShell22.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\Common\Projects\NTSound20.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\Common\Projects\NTSoSweet.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\Common\Projects\NTService20.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\Common\Projects\NTAdvFTP61.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\Common\Projects\NTControls22.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\Common\Projects\NTControls30.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\Common\Projects\NTImaging10.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\Blacklawn\Projects\Blacklawn.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\Blacklawn\Projects\BlkLServer.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\To-Doster\Projects\ToDoster.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\HouseOfGlass\Projects\HouseOfGlass.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\Max-FTP\Projects\MaxIDE.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\Max-FTP\Projects\MaxFTP.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\Max-FTP\Projects\MaxService.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\Max-FTP\Projects\MaxUtility.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\MaxLand\Projects\MaxLandApp.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\RemindMe\Projects\RmdMeSrv.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\RemindMe\Projects\RemindMe.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\RemindMe\Projects\Utility.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\Creata-Tree\Projects\CreataTree.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\Sequencer\Projects\SoSouiXSeq.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\IdentAuth\Projects\Reload.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\IdentAuth\Projects\Ident.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\CrayonStill\Projects\CrayonStall.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\CrayonStill\Projects\CrayonStiff.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\CrayonStill\Projects\CrayonStill.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\InstallerStar\Projects\Wizard.vbp"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\InstallerStar\Projects\Remove.vbp"



"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make "C:\Development\Neotext\BasicNeotext\Projects\BasicNeotext_DLL.vbp" /d VBIDE=0:modRegistry=-1
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make "C:\Development\Neotext\BasicNeotext\Projects\BasicNeotext_EXE.vbp" /d VBIDE=0:modRegistry=-1
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /make "C:\Development\Neotext\BasicNeotext\Projects\BasicNotable.vbp" /d VBIDE=0



echo Compiling install packages to the development location...

echo 1. Packing Max-Ftp 
cd \Development\Neotext\InstallerStar\Binary
Wizard.exe /compile Max-FTP


echo 2. Packing Blacklawn 
cd \Development\Neotext\InstallerStar\Binary
Wizard.exe /compile Blacklawn


echo 3. Packing HouseOfGlass 
cd \Development\Neotext\InstallerStar\Binary
Wizard.exe /compile HouseOfGlass


echo 4. Packing Creata-Tree 
cd \Development\Neotext\InstallerStar\Binary
Wizard.exe /compile Creata-Tree


echo 5. Packing MaxLand 
cd \Development\Neotext\InstallerStar\Binary
Wizard.exe /compile MaxLand


echo 6. Packing RemindMe 
cd \Development\Neotext\InstallerStar\Binary
Wizard.exe /compile RemindMe


echo 7. Packing To-Doster 
cd \Development\Neotext\InstallerStar\Binary
Wizard.exe /compile To-Doster


echo 8. Packing Sequencer 
cd \Development\Neotext\InstallerStar\Binary
Wizard.exe /compile Sequencer


echo 9. Packing Ident 
cd \Development\Neotext\InstallerStar\Binary
Wizard.exe /compile IdentAuth


"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\Max-FTP\Deploy\Max-FTP v6.1.0.exe"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\Blacklawn\Deploy\Blacklawn v1.1.0.exe"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\HouseOfGlass\Deploy\HouseOfGlass v1.0.0.exe"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\Creata-Tree\Deploy\Creata-Tree v3.1.0.exe"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\MaxLand\Deploy\MaxLand v2.2.0.exe"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\RemindMe\Deploy\RemindMe v2.1.0.exe"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\To-Doster\Deploy\To-Doster v1.2.0.exe"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\Sequencer\Deploy\Sequencer v5.0.5.exe"
"C:\Program Files\Microsoft Visual Studio\VB98\vbn.exe" /sign "C:\Development\Neotext\IdentAuth\Deploy\IdentAuth v7.1.0.exe"


echo Done. 

"C:\Program Files\NSIS\makensis.exe" /V1 "C:\Development\Neotext\BasicNeotext\Deploy\BasicNeotext v3.0.0.nsi"


"C:\Program Files\Microsoft Visual Studio\VB98\Uninstall.exe" /S

"C:\Development\Neotext\BasicNeotext\Deploy\BasicNeotext v3.0.0.exe" /S


echo Installing local packages with the silent enabled for testing...
echo 1. Installing Blacklawn 
"C:\Development\Neotext\Blacklawn\Deploy\Blacklawn v1.1.0.exe" /Q
echo 2. Installing Max-FTP 
"C:\Development\Neotext\Max-FTP\Deploy\Max-FTP v6.1.0.exe" /Q
echo 3. Installing HouseOfGlass 
"C:\Development\Neotext\HouseOfGlass\Deploy\HouseOfGlass v1.0.0.exe" /Q
echo 4. Installing Creata-Tree 
"C:\Development\Neotext\Creata-Tree\Deploy\Creata-Tree v3.1.0.exe" /Q
echo 5. Installing MaxLand 
"C:\Development\Neotext\MaxLand\Deploy\MaxLand v2.2.0.exe" /Q
echo 6. Installing RemindMe 
"C:\Development\Neotext\RemindMe\Deploy\RemindMe v2.1.0.exe" /Q
echo 7. Installing To-Doster 
"C:\Development\Neotext\To-Doster\Deploy\To-Doster v1.2.0.exe" /Q
echo 8. Installing Sequencer 
"C:\Development\Neotext\Sequencer\Deploy\Sequencer v5.0.5.exe" /Q
echo 9. Installing Ident 
"C:\Development\Neotext\IdentAuth\Deploy\IdentAuth v7.1.0.exe" /Q

rem echo 9. Installing KadPatch
rem "C:\Development\Neotext\KadPatch\Deploy\KadPatch v1.0.0.exe" /Q
rem echo 10.Installing CrayonStill 
rem "C:\Development\Neotext\CrayonStill\Deploy\CrayonStill v0.0.0.exe" /Q




echo Done. 




