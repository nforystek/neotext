
;NSIS 2.24
!define APPPATH "C:\Development\Neotext"
!define APPNAME "BasicNeotext"
!define APPVER "3.0.0"

;!define SIGNCODE

Name "${APPNAME}"
SetCompressor /SOLID bzip2
SetCompress force


;!ifdef SIGNCODE
;	!ifdef INNER
;		OutFile "$%TEMP%\signinst.exe"
;		SetCompress off
;	!else
;		SetCompressor /SOLID lzma
;		!system "$\"${NSISDIR}\makensis$\" /V1 /DINNER $\"${APPPATH}\${APPNAME}\Deploy\${APPNAME} v${APPVER}.nsi$\"" = 0
;		!system "$%TEMP%\signinst.exe" = 2
;		!system "$\"C:\Program Files\Microsoft Visual Studio\VB98\VBN.EXE$\" /sign $\"$%TEMP%\Uninstall.exe$\""=0
;		OutFile "${APPNAME} v${APPVER}.exe"
;	!endif
;!else
	OutFile "${APPNAME} v${APPVER}.exe"
;!endif

InstallDir "$PROGRAMFILES\${APPNAME}"
InstallDirRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "InstallLocation"
BrandingText "Built on ${__DATE__} at ${__TIME__}"

InstProgressFlags smooth

UninstallIcon "${APPPATH}\Source\Media\win-uninstall.ico"
CheckBitmap "${APPPATH}\Source\Media\classic.bmp"


VIProductVersion "${APPVER}.0"
VIAddVersionKey "ProductName" "${APPNAME}"
VIAddVersionKey "Comments" "http://www.neotext.org"
VIAddVersionKey "CompanyName" "Neotext"
VIAddVersionKey "LegalTrademarks" ""
VIAddVersionKey "LegalCopyright" "© 1999,2006,2013 by Nicholas Forystek"
VIAddVersionKey "FileDescription" "${APPNAME} Install Package"
VIAddVersionKey "FileVersion" "${APPVER}.0"

VIAddVersionKey "VersionLegalTrademarks" ""
VIAddVersionKey "VersionLegalCopyright" "© 1999,2006,2013 by Nicholas Forystek"
VIAddVersionKey "VersionFileDescription" "${APPNAME} Install Package"
VIAddVersionKey "VersionFileVersion" "${APPVER}.0"
VIAddVersionKey "VersionProductName" "${APPNAME}"

RequestExecutionLevel admin
InstallColors 000000 FFFFFF
WindowIcon on


!macro SignedUninstaller
	!ifdef WROTEUNINSTALLER
		!undef WROTEUNINSTALLER
;		!ifdef SIGNCODE
;			!ifdef INNER		
;				WriteUninstaller "$%TEMP%\Uninstall.exe"
;				Quit
;			!else		
;				ReserveFile $%TEMP%\Uninstall.exe
;				!system "erase $%TEMP%\signinst.exe" = 0
;				!system "erase $%TEMP%\Uninstall.exe" = 0
;			!endif
;		!endif
	!else
		!define WROTEUNINSTALLER
		WriteUninstaller "$INSTDIR\Uninstall.exe"
;		!ifdef SIGNCODE
;			!ifndef INNER
;				SetOutPath $INSTDIR
;				SetOverwrite on
;				File $%TEMP%\Uninstall.exe
;			!endif
;		!endif
	!endif
!macroend

var PGPKEY 
var PGPEND

!macro PGPPrivateBlock
	StrCpy $PGPKEY "-----BEGIN PGP PRIVATE KEY BLOCK-----$\nVersion: GnuPG v1"
	StrCpy $PGPEND "lQO+BFdFHikBCADpE7CntHnSrox2qPU3lPgZyOuAW+taG4utbmYmyqywyozJwjrW$\nNHZA25CA/RkPHZWTh55KBl7Dg+vns+CMck+aSVIgea+uuornQf4P9T9BNMi+sjsQ$\nDruOhPl1US6c0jM8r2xD1kgjiuc9P9rZ6rURkjcHFVGmgvZxj+lKu65ffOwV0N4x$\nZ50+JWNc6CyeGDnxK9xjuXfW98kECzasvnrqsH0Drk4SIq2n4hyj6pIeQopZ7Xq1$\nmrF+CZ/GYV6+Z0cmCiMl7O4ypIczR5L3MBFqvZ+qCa8bsdQ3xq7LuoM9RNkrNuYP$\nwC7AxQYgWCdFOi/sDyuMq93I/u67UzXsLJRBABEBAAH+AgMCe9GpHbv37H5gSCth$\ngzjuhdWdWnm2iPfMImJrtaS3EARQ2wh2cnobg8U/wrYObuZzI0jWGcTMxLfQ9LEP$\nZkIxgQx3YAWMjkrmRIa2obb2kbtOT7FBT5T8Lr562m3ouWwkI31Loq/ADRcDNchI$\nb0/pKR1JxJ0am0SjWDkKm1DiLyT/0IicwAnusjJmeTUVe65AM6giMC8RM32TueFP$\nASyL36wbxeldfFfwFxMFN+YfibKF5a9i4LncuIRlBWDSLdOvgIKP8/LAY5pdo6xJ$\n4GsTiIEN6Jel093c0/FqEw5lFDEtKQ7zu0edIEB00/3+H4sMr+wyU2nhkbmde0e4$\nWfUaqgcvzXdKkPnlw4Ou7v2zDrDo/VBMDli53wfwA4svnN35HUNxcW2ONkKUYeCj$\ngVhyKuyfgwEmYBGzUb0mzHHRKHNTsebenuYiqSneBIdLoSLeZeUAdgY1hPiSalk8$\n51YyxF6Mp75NkzVgKbeM1oYy9+ScVsN8nFhoBuueMAz0IBRFwXtbtupR1oSLY6fc$\nbY8WXdnlLkx9CaqOUkh27fex+WJH+/sCIyXKyvpSYcSZLXN5UEg+KTyWramerklU$\nqsPLTkZQ28oE20aSzdZqJi06lqFmPU7lhrXSn6gY3a30/NjSW4pQZZ73biokhiYi$\n1dUhEw2UiQRQwWzscZQh/lSNXsAXXtTRmPBeiU7z4mvUJBCNpVX7fK7oHH1PtLY0$\neAjgV2OICwT0tUK48/GPmTZVhUqYw/cuURhm0QLzRm9bIJNor+YFBQoUUg9wfEfa$\n9BAuihOgQhvDoek6QI03Q9NlV2nV7jEmXuIxlwxBCGxyaKj+d+w2qwGXmfSX7pNQ$\n5wDY0pdSMTuB9KlE8mLR68fdwozd8uTZ4gzkY0PYrUbTEBvO/f3pAxbWaeeYCtXE$\nkLRGTmljaG9sYXMgUmFuZGFsbCBGb3J5c3RlayAoQ29tcHV0ZXIgU2NpZW50aXN0$\nKSA8bmZvcnlzdGVrQG5lb3RleHQub3JnPokBOAQTAQIAIgUCV0UeKQIbAwYLCQgH$\nAwIGFQgCCQoLBBYCAwECHgECF4AACgkQ9qogsfEJZVlc6Qf/XGK9tyVFjcJKcuBr$\nQDalnlbNz+fgD6aveE5ZNww5mNdH/r625mZxiljWcquDAPQMS1T5WFTEDJkmuxS7$\nLwlNDftGi5QNmGC4IpxcxVxoHlvDjDBUUhciEyo9QoNrqxaz64YmMeZn5fddYm4e$\ntVKt9HFMjRX2E7Wxkj21rSV8AvFPaLOSfxTHPofEcBHj40E7V4F0PpZqLwL5nhmz$\nACvDWDoaFsE56E6AvYOlyr+etlOS2+vqA1qn4Fz/Cy2zh9UqQwA8mdf1B/ioyUld$\n6CeuJhpjoK6195bOCfITvIMlSg2D1OtcRikFRJcISLm1T12p2j+xD1cruKIkHW1J$\nLlg13g==$\n=Qb07$\n-----END PGP PRIVATE KEY BLOCK-----"
	!insertmacro SwapVariables $PGPKEY $PGPEND
!macroend

!macro SwapVariables varone vartwo
	Push ${varone}
	Push ${vartwo}
	Pop ${varone}
	Push ${vartwo}
	Pop ${varone}
	Pop ${vartwo}
!macroend

!include "${NSISDIR}\Include\VB6RunTime.nsh"


Page components

 PageEx license
   LicenseText "Please read the information below and agree to use$\nat your own risk by clicking the $\"I Agree$\" button."
   LicenseData "C:\Development\Neotext\BasicNeotext\Media\WARNING.rtf"
	Caption ": Use at your own risk"
 PageExEnd


Page directory
Page instfiles

UninstPage uninstConfirm
UninstPage instfiles

Var AlreadyInstalled

!macro InstallLibaries
IfFileExists "$INSTDIR\*.exe" 0 new_installation ;Replace MyApp.exe with your application filename
StrCpy $AlreadyInstalled 1
new_installation:
!insertmacro VB6RunTimeInstall C:\Development\Neotext\Windows\Runtime $AlreadyInstalled
SetOverwrite ifdiff
!insertmacro InstallLib REGDLL $AlreadyInstalled REBOOT_PROTECTED "C:\Development\Neotext\BasicNeotext\Binary\BasicService.DLL" "$SYSDIR\BasicService.DLL" "$SYSDIR"
!insertmacro InstallLib DLL    $AlreadyInstalled REBOOT_PROTECTED "${APPPATH}\Windows\System\user32.dll" "$SYSDIR\user32.dll" "$SYSDIR"
!insertmacro InstallLib DLL    $AlreadyInstalled REBOOT_PROTECTED "${APPPATH}\Windows\System\kernel32.dll" "$SYSDIR\kernel32.dll" "$SYSDIR"
!insertmacro InstallLib DLL    $AlreadyInstalled REBOOT_PROTECTED "${APPPATH}\Windows\System\advapi32.dll" "$SYSDIR\advapi32.dll" "$SYSDIR"
!insertmacro InstallLib DLL    $AlreadyInstalled REBOOT_PROTECTED "${APPPATH}\Windows\ActiveX\ole32.dll" "$SYSDIR\ole32.dll" "$SYSDIR"
!insertmacro InstallLib DLL    $AlreadyInstalled REBOOT_PROTECTED "${APPPATH}\Windows\System\gdi32.dll" "$SYSDIR\gdi32.dll" "$SYSDIR"
!insertmacro InstallLib REGDLL $AlreadyInstalled REBOOT_PROTECTED "${APPPATH}\Windows\ActiveX\scrrun.dll" "$SYSDIR\scrrun.dll" "$SYSDIR"
!insertmacro InstallLib DLL    $AlreadyInstalled REBOOT_PROTECTED "${APPPATH}\Windows\ActiveX\shell32.dll" "$SYSDIR\shell32.dll" "$SYSDIR"
!insertmacro InstallLib REGDLL $AlreadyInstalled REBOOT_PROTECTED "${APPPATH}\Windows\ActiveX\comdlg32.ocx" "$SYSDIR\comdlg32.ocx" "$SYSDIR"
!insertmacro InstallLib REGDLL $AlreadyInstalled REBOOT_PROTECTED "${APPPATH}\Windows\ActiveX\richtx32.ocx" "$SYSDIR\richtx32.ocx" "$SYSDIR"
!insertmacro InstallLib REGDLL $AlreadyInstalled REBOOT_PROTECTED "${APPPATH}\Common\Binary\NTControls22.ocx" "$SYSDIR\NTControls22.ocx" "$SYSDIR"
!insertmacro InstallLib REGDLL $AlreadyInstalled REBOOT_PROTECTED "${APPPATH}\Windows\ActiveX\ntsvc.ocx" "$SYSDIR\ntsvc.ocx" "$SYSDIR"
!insertmacro InstallLib DLL    $AlreadyInstalled REBOOT_PROTECTED "${APPPATH}\Windows\System\ws2_32.dll" "$SYSDIR\ws2_32.dll" "$SYSDIR"
!insertmacro InstallLib DLL    $AlreadyInstalled REBOOT_PROTECTED "${APPPATH}\Windows\System\wsock32.dll" "$SYSDIR\wsock32.dll" "$SYSDIR"
!macroend
!macro UninstallLibaries
;!insertmacro VB6RunTimeUnInstall
;!insertmacro UnInstallLib REGDLL SHARED NOREBOOT_PROTECTED "$SYSDIR\BasicService.DLL"
;!insertmacro UnInstallLib DLL    SHARED REBOOT_PROTECTED "$SYSDIR\user32.dll"
;!insertmacro UnInstallLib DLL    SHARED REBOOT_PROTECTED "$SYSDIR\kernel32.dll"
;!insertmacro UnInstallLib DLL    SHARED REBOOT_PROTECTED "$SYSDIR\advapi32.dll"
;!insertmacro UnInstallLib REGDLL SHARED NOREBOOT_PROTECTED "$SYSDIR\ole32.dll"
;!insertmacro UnInstallLib DLL    SHARED REBOOT_PROTECTED "$SYSDIR\gdi32.dll"
;!insertmacro UnInstallLib REGDLL SHARED NOREBOOT_PROTECTED "$SYSDIR\scrrun.dll"
;!insertmacro UnInstallLib REGDLL SHARED NOREBOOT_PROTECTED "$SYSDIR\shell32.dll"
;!insertmacro UnInstallLib REGDLL SHARED NOREBOOT_PROTECTED "$SYSDIR\comdlg32.ocx"
;!insertmacro UnInstallLib REGDLL SHARED NOREBOOT_PROTECTED "$SYSDIR\richtx32.ocx"
;!insertmacro UnInstallLib REGDLL SHARED NOREBOOT_NOTPROTECTED "$SYSDIR\NTControls22.ocx"
;!insertmacro UnInstallLib REGDLL SHARED NOREBOOT_PROTECTED "$SYSDIR\ntsvc.ocx"
;!insertmacro UnInstallLib DLL    SHARED REBOOT_PROTECTED "$SYSDIR\ws2_32.dll"
;!insertmacro UnInstallLib DLL    SHARED REBOOT_PROTECTED "$SYSDIR\wsock32.dll"
!macroend


;Section "Create Restore Point"
;	IfFileExists "$SYSDIR\wbem\wmic.exe" +1 +2
;	ExecWait "$SYSDIR\wbem\wmic.exe /Namespace:\\root\default Path SystemRestore Call CreateRestorePoint $\"Before Installation of ${APPNAME}$\", 100, 12"
;SectionEnd


Icon "${APPPATH}\${APPNAME}\Media\MS VB.ico"
UninstallText "This will uninstall ${APPNAME},  Click next to continue."


ComponentText "This will install ${APPNAME} on your computer."
DirText "Please provide the directory where VB6.EXE is located to continue:"

AllowRootDirInstall false
RequestExecutionLevel Admin

Section "CodeSign Switch"


	IfFileExists "$INSTDIR\Uninstall.exe" +1 +3
	ExecWait "$INSTDIR\Uninstall.exe /S"
	Goto +7
	IfSilent +6
	IfFileExists "$INSTDIR\BasicNT-Uninstall.exe" +3
	IfFileExists "$INSTDIR\PRE.EXE" +2
	IfFileExists "$INSTDIR\VBE.EXE" +1 +3
	MessageBox MB_OK "Please uninstall any previous version of VB6 Neotext Basic Enhancements AddIn."
	Abort

	!insertmacro InstallLibaries

	SectionIn RO
	SetOverwrite on

	SetOutPath $INSTDIR

	File "${APPPATH}\BasicNeotext\Binary\README.txt"
	
	File "${APPPATH}\${APPNAME}\Binary\VBN.EXE"
	ExecWait "$INSTDIR\VBN.exe /regserver"
	ExecWait "$INSTDIR\VBN.exe /install"

	;File "${APPPATH}\${APPNAME}\Binary\Link.exe"
	;File "${APPPATH}\${APPNAME}\Binary\C2.exe"

	;WriteRegStr HKLM "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers" "$INSTDIR\LINK.exe" "~ RUNASADMIN"
	;WriteRegStr HKLM "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers" "$INSTDIR\C2.exe" "~ RUNASADMIN"

	WriteRegDWORD HKCU "Software\Microsoft\Visual Basic\6.0\Addins\BasicNeotext.Connect" "CommandLineSafe" 0x00000000
	WriteRegStr HKCU "Software\Microsoft\Visual Basic\6.0\Addins\BasicNeotext.Connect" "Description" "Enhancements for Visual Basic 6.0"
	WriteRegStr HKCU "Software\Microsoft\Visual Basic\6.0\Addins\BasicNeotext.Connect" "FriendlyName" "VB 6 Neotext Basic"
	WriteRegDWORD HKCU "Software\Microsoft\Visual Basic\6.0\Addins\BasicNeotext.Connect" "LoadBehavior" 0x00000003

	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayName" "${APPNAME}"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "UninstallString" '"$INSTDIR\Uninstall.exe"'
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "Publisher" "SoSouiX"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "ProductVersion" '"${APPVER}.0.0"'
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "ModifyPath" '"$INSTDIR\Uninstall.exe"'
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayIcon" '"$INSTDIR\VBN.exe"'
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "InstallLocation" '"$INSTDIR"'
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "HelpLink" '"https://www.neotext.org"'
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "HelpTelephone" '"+1-952-457-9224"'
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayVersion" '"${APPNAME} v${APPVER}"'
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "VersionMajor" '"1"'
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "VersionMinor" '"0"'
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "ServicePackState" "0"


	WriteRegStr HKLM "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers" "$INSTDIR\VBN.exe" "~ RUNASADMIN"
	WriteRegStr HKLM "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers" "$INSTDIR\Uninstall.exe" "~ RUNASADMIN"

	!insertmacro SignedUninstaller

SectionEnd

Section "Mouse Wheel Fix"

	SetOverwrite on
	SetOutPath $INSTDIR
	WriteRegDWORD HKCU "Software\Microsoft\Visual Basic\6.0\Addins\BasicNeotext.Scroller" "CommandLineSafe" 0x00000000
	WriteRegStr HKCU "Software\Microsoft\Visual Basic\6.0\Addins\BasicNeotext.Scroller" "Description" "Enhancements for Visual Basic 6.0"
	WriteRegStr HKCU "Software\Microsoft\Visual Basic\6.0\Addins\BasicNeotext.Scroller" "FriendlyName" "VB 6 Mouse Wheel"
	WriteRegDWORD HKCU "Software\Microsoft\Visual Basic\6.0\Addins\BasicNeotext.Scroller" "LoadBehavior" 0x00000003
	File "${APPPATH}\${APPNAME}\Binary\VBN.DLL"
	RegDLL "$INSTDIR\VBN.DLL"
SectionEnd

;Section "API Export Template"
;	CreateDirectory "$INSTDIR\Template\Projects"
;	SetOutPath "$INSTDIR\Template\Projects"
;	File "${APPPATH}\${APPNAME}\Binary\Template\Projects\Library.vbp"
;	File "${APPPATH}\${APPNAME}\Binary\Template\Projects\Library.vbw"
;	File "${APPPATH}\${APPNAME}\Binary\Template\Projects\Exports.def"
;	File "${APPPATH}\${APPNAME}\Binary\Template\Projects\Exports.bas"
;	File "${APPPATH}\${APPNAME}\Binary\Template\Projects\Exports.cls"
;	File "${APPPATH}\${APPNAME}\Binary\Template\Projects\README.TXT"
;SectionEnd

Section "Service Template"

	SetOverwrite on
	CreateDirectory "$INSTDIR\Template\Projects"
	SetOutPath "$INSTDIR\Template\Projects"
	File "${APPPATH}\${APPNAME}\Binary\Template\Projects\Service.vbp"
	File "${APPPATH}\${APPNAME}\Binary\Template\Projects\Service.vbw"
	File "${APPPATH}\${APPNAME}\Binary\Template\Projects\Module.bas"
	File "${APPPATH}\${APPNAME}\Binary\Template\Projects\Class.cls"
	File "${APPPATH}\${APPNAME}\Binary\Template\Projects\README.TXT"
	SetOutPath $SYSDIR
	File "${APPPATH}\${APPNAME}\Binary\BasicService.DLL"
	RegDLL "$SYSDIR\BasicService.DLL"
SectionEnd

;Section /o "Notable Ink Pad"
;
;	SetOutPath $INSTDIR
;
;	SetOverwrite on
;	
;	File "${APPPATH}\${APPNAME}\Binary\SINK.exe"
;
;	WriteRegStr HKCR ".ink" "" "BasicNeotext.Ink"
;	WriteRegStr HKCR "BasicNeotext.Ink" "" "Batch Ink File"
;	WriteRegStr HKCR "BasicNeotext.Ink\DefaultIcon" "" "$INSTDIR\SINK.exe,0"
;	WriteRegStr HKCR "BasicNeotext.Ink\shell\Run\command" "" '"$INSTDIR\SINK.exe" /run "%1"'
;	WriteRegStr HKCR "BasicNeotext.Ink\shell\Open\command" "" '"$INSTDIR\SINK.exe" "%1"'
;	WriteRegStr HKCR "BasicNeotext.Ink\shell\Run Exit\command" "" '"$INSTDIR\SINK.exe" /runexit "%1"'
;	WriteRegStr HKCR "BasicNeotext.Ink\shell\Run Hidden\command" "" '"$INSTDIR\SINK.exe" /runhide "%1"'
;	Push $0
;
;	ReadRegStr $0 HKCR "batfile\shell\Open\command" ""
;	StrCmp $0 "" +1 +2
;	WriteRegStr HKCR "batfile\shell\Open\command" "" '"%1" %*'
;
;	ReadRegStr $0 HKCR "batfile\shell\Edit\command" ""
;	StrCmp $0 "" +1 +2
;	WriteRegStr HKCR "batfile\shell\Edit\command" "" '"$WINDIR\notepad.exe" "%1"'
;
;	ReadRegStr $0 HKCR ".bat" ""
;	StrCmp $0 "batfile" +2 +1
;	WriteRegStr HKCR ".bat" "" 'batfile'
;
;	ReadRegStr $0 HKCR ".bat\Shell" ""
;	StrCmp $0 "Open" +2 +1
;	WriteRegStr HKCR ".bat\Shell" "" 'Open'
;
;	ReadRegStr $0 HKCR ".bat\PersistentHandler" ""
;	StrCmp $0 "Open" +2 +1
;	WriteRegStr HKCR ".bat\PersistentHandler" "" '{5e941d80-bf96-11cd-b579-08002b30bfeb}'
;	Pop $0
;
;	WriteRegStr HKCR ".bat\shell\Edit in Notable" "" "Edit in Notable"
;	WriteRegStr HKCR ".bat\shell\Exec in Notable" "" "Exec in Notable"
;	WriteRegStr HKCR ".bat\shell\Edit in Notable\command" "" '"$INSTDIR\SINK.exe" "%1"'
;	WriteRegStr HKCR ".bat\shell\Exec in Notable\command" "" '"$INSTDIR\SINK.exe" /runexit "%1"'
;
;	WriteRegStr HKCR ".bat\OpenWithProgids" "BasicNeotext.Ink" ""
;	WriteRegStr HKLM "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers" "$INSTDIR\SINK.exe" "~ RUNASADMIN"
;
;SectionEnd


Section /o "Help Documents"
	CreateDirectory "$INSTDIR\BasicNeotext"

	SetOutPath $INSTDIR\BasicNeotext"
	
        CreateDirectory "$INSTDIR\BasicNeotext\Media"
        SetOutPath "$INSTDIR\BasicNeotext\Media"
        CreateDirectory "$INSTDIR\BasicNeotext\Media\Bullet Icons"
        SetOutPath "$INSTDIR\BasicNeotext\Media\Bullet Icons"
        File "${APPPATH}\${APPNAME}\Binary\Help\Media\Bullet Icons\bookClosed.gif"
        File "${APPPATH}\${APPNAME}\Binary\Help\Media\Bullet Icons\bookOpen.gif"
        File "${APPPATH}\${APPNAME}\Binary\Help\Media\Bullet Icons\overview.gif"
        File "${APPPATH}\${APPNAME}\Binary\Help\Media\Bullet Icons\topic.gif"
        File "${APPPATH}\${APPNAME}\Binary\Help\Media\Bullet Icons\world.gif"
        CreateDirectory "$INSTDIR\BasicNeotext\Media\PlusMinus"
        SetOutPath "$INSTDIR\BasicNeotext\Media\PlusMinus"
        CreateDirectory "$INSTDIR\BasicNeotext\Media\PlusMinus\Black"
        SetOutPath "$INSTDIR\BasicNeotext\Media\PlusMinus\Black"
        File "${APPPATH}\${APPNAME}\Binary\Help\Media\PlusMinus\Black\minus.gif"
        File "${APPPATH}\${APPNAME}\Binary\Help\Media\PlusMinus\Black\plus.gif"
        CreateDirectory "$INSTDIR\BasicNeotext\Media\Treelines"
        SetOutPath "$INSTDIR\BasicNeotext\Media\Treelines"
        CreateDirectory "$INSTDIR\BasicNeotext\Media\Treelines\Black"
        SetOutPath "$INSTDIR\BasicNeotext\Media\Treelines\Black"
        File "${APPPATH}\${APPNAME}\Binary\Help\Media\Treelines\Black\btm.gif"
        File "${APPPATH}\${APPNAME}\Binary\Help\Media\Treelines\Black\hline.gif"
        File "${APPPATH}\${APPNAME}\Binary\Help\Media\Treelines\Black\mid.gif"
        File "${APPPATH}\${APPNAME}\Binary\Help\Media\Treelines\Black\top.gif"
        File "${APPPATH}\${APPNAME}\Binary\Help\Media\Treelines\Black\vline.gif"
        File "${APPPATH}\${APPNAME}\Binary\Help\Media\202.gif"
        CreateDirectory "$INSTDIR\BasicNeotext\Overview_files"
        SetOutPath "$INSTDIR\BasicNeotext\Overview_files"
        File "${APPPATH}\${APPNAME}\Binary\Help\AdditionalAdd-Ins.htm"
        File "${APPPATH}\${APPNAME}\Binary\Help\BasicNeotextHelp.htm"
        File "${APPPATH}\${APPNAME}\Binary\Help\BitImbalances.htm"
        File "${APPPATH}\${APPNAME}\Binary\Help\CodeSigningBuilds.htm"
        File "${APPPATH}\${APPNAME}\Binary\Help\CommandLineSwitches.htm"
        File "${APPPATH}\${APPNAME}\Binary\Help\Contents.htm"
        File "${APPPATH}\${APPNAME}\Binary\Help\custom.js"
        File "${APPPATH}\${APPNAME}\Binary\Help\Enhancements.htm"
        File "${APPPATH}\${APPNAME}\Binary\Help\index.htm"
        File "${APPPATH}\${APPNAME}\Binary\Help\InstallandUninstall.htm"
        File "${APPPATH}\${APPNAME}\Binary\Help\Overview.htm"
        File "${APPPATH}\${APPNAME}\Binary\Help\ProjectTemplates.htm"
        File "${APPPATH}\${APPNAME}\Binary\Help\ReadMe.htm"
        File "${APPPATH}\${APPNAME}\Binary\Help\ReleaseExecutables.htm"
        File "${APPPATH}\${APPNAME}\Binary\Help\SecurityFeatures.htm"
        File "${APPPATH}\${APPNAME}\Binary\Help\SettingsPreservation.htm"
        File "${APPPATH}\${APPNAME}\Binary\Help\tree.html"
        File "${APPPATH}\${APPNAME}\Binary\Help\vbblue.png"
        File "${APPPATH}\${APPNAME}\Binary\Help\VBMouseWheelFix.htm"
        File "${APPPATH}\${APPNAME}\Binary\Help\VisualBasic60.htm"
        File "${APPPATH}\${APPNAME}\Binary\Help\VisualBasicNeotext.htm"
        File "${APPPATH}\${APPNAME}\Binary\Help\Warning.htm"
        File "${APPPATH}\${APPNAME}\Binary\Help\WindowsService.htm"
SectionEnd

;!macro ReservePack6
;	ReserveFile "${APPPATH}\${APPNAME}\Deploy\Downloads\Vs6sp6.exe"
;	ReserveFile "${APPPATH}\${APPNAME}\Deploy\Downloads\Vs6sp61.exe"
;	ReserveFile "${APPPATH}\${APPNAME}\Deploy\Downloads\Vs6sp62.exe"
;!macroend

;!macro ReservePack6B
;	ReserveFile "${APPPATH}\${APPNAME}\Deploy\Downloads\Vs6sp6B.exe"
;	ReserveFile "${APPPATH}\${APPNAME}\Deploy\Downloads\Vs6sp6B1.exe"
;	ReserveFile "${APPPATH}\${APPNAME}\Deploy\Downloads\Vs6sp6B2.exe"
;	ReserveFile "${APPPATH}\${APPNAME}\Deploy\Downloads\Vs6sp6B3.exe"
;	ReserveFile "${APPPATH}\${APPNAME}\Deploy\Downloads\Vs6sp6B4.exe"
;!macroend

;!macro ServicePacks
;
;	StrCmp $0 "0" +1 +10
;	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "ServicePackState" "1"
;	CreateDirectory "$%TEMP%\Vs6sp6"
;	!insertmacro ReservePack6
;	ExecWait "$%TEMP%\Vs6sp6\Vs6sp6.exe /Q /C /T:$%TEMP%\Vs6sp6"
;	ExecWait "$%TEMP%\Vs6sp6\Vs6sp61.exe /Q /C /T:$%TEMP%\Vs6sp6"
;	ExecWait "$%TEMP%\Vs6sp6\Vs6sp62.exe /Q /C /T:$%TEMP%\Vs6sp6"
;	ExecWait "$%TEMP%\Vs6sp6\setupsp6.exe /QT"
;	RmDir /r "$%TEMP%\Vs6sp6"
;	StrCmp $0 "2" +1 +14
;	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "ServicePackState" "2"
;	CreateDirectory "$%TEMP%\Vs6sp6B"
;	!insertmacro ReservePack6B
;	ExecWait "$%TEMP%\Vs6sp6B\Vs6sp6B.exe /Q /C /T:$%TEMP%\Vs6sp6B"
;	ExecWait "$%TEMP%\Vs6sp6B\Vs6sp6B1.exe /Q /C /T:$%TEMP%\Vs6sp6B"
;	ExecWait "$%TEMP%\Vs6sp6B\Vs6sp6B2.exe /Q /C /T:$%TEMP%\Vs6sp6B"
;	ExecWait "$%TEMP%\Vs6sp6B\Vs6sp6B3.exe /Q /C /T:$%TEMP%\Vs6sp6B"
;	ExecWait "$%TEMP%\Vs6sp6B\Vs6sp6B4.exe /Q /C /T:$%TEMP%\Vs6sp6B"
;	ExecWait "$%TEMP%\Vs6sp6B\setupsp6.exe /QT"
;	RmDir /r "$%TEMP%\Vs6sp6B"
;	StrCmp $0 "3" +1 +2
;	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "ServicePackState" "4"
;
;!macroend

;!macro unServicePacks
;
;	StrCmp $0 "3" +1 +14
;	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "ServicePackState" "2"
;	CreateDirectory "$%TEMP%\Vs6sp6B"
;	!insertmacro ReservePack6B
;	ExecWait "$%TEMP%\Vs6sp6B\Vs6sp6B.exe /Q /C /T:$%TEMP%\Vs6sp6B"
;	ExecWait "$%TEMP%\Vs6sp6B\Vs6sp6B1.exe /Q /C /T:$%TEMP%\Vs6sp6B"
;	ExecWait "$%TEMP%\Vs6sp6B\Vs6sp6B2.exe /Q /C /T:$%TEMP%\Vs6sp6B"
;	ExecWait "$%TEMP%\Vs6sp6B\Vs6sp6B3.exe /Q /C /T:$%TEMP%\Vs6sp6B"
;	ExecWait "$%TEMP%\Vs6sp6B\Vs6sp6B4.exe /Q /C /T:$%TEMP%\Vs6sp6B"
;	ExecWait "$%TEMP%\Vs6sp6B\setupsp6.exe /U /QT"
;	RmDir /r "$%TEMP%\Vs6sp6B"
;	StrCmp $0 "2" +1 +10
;	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "ServicePackState" "1"
;	CreateDirectory "$%TEMP%\Vs6sp6"
;	!insertmacro ReservePack6
;	ExecWait "$%TEMP%\Vs6sp6\Vs6sp6.exe /Q /C /T:$%TEMP%\Vs6sp6"
;	ExecWait "$%TEMP%\Vs6sp6\Vs6sp61.exe /Q /C /T:$%TEMP%\Vs6sp6"
;	ExecWait "$%TEMP%\Vs6sp6\Vs6sp62.exe /Q /C /T:$%TEMP%\Vs6sp6"
;	ExecWait "$%TEMP%\Vs6sp6\setupsp6.exe /U /QT"
;	RmDir /r "$%TEMP%\Vs6sp6"
;	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "ServicePackState" "0"
;!macroend

;Section /o "Visual Studio SP6"
;	Push $0
;	ReadRegStr $0 HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "ServicePackState"
;	!insertmacro ServicePacks
;	Pop $0
;SectionEnd

Function .onInit		
	!insertmacro PGPPrivateBlock
	!insertmacro SignedUninstaller
;	Push $0
	StrCpy $INSTDIR "C:\Program Files\Microsoft Visual Studio\VB98"
;	ReadRegStr $0 HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "ServicePackState"
;	StrCmp $0 "" outinit
;	!insertmacro ServicePacks
;	Quit
;	outinit:
;	Pop $0
FunctionEnd

Function un.onInit
	!insertmacro PGPPrivateBlock
FunctionEnd

Function .onVerifyInstDir
	IfFileExists "$INSTDIR\VB6.EXE" +2
	Abort
FunctionEnd

Section "Uninstall"

	Delete "$INSTDIR\ScriptInk.dat"
	Delete "$INSTDIR\codesign.dat"
	Delete "$INSTDIR\LnkOutput.txt"

	IfSilent +3
	IfFileExists "$INSTDIR\VBN.exe" +1 +3
	ExecWait "$INSTDIR\VBN.EXE /uninstall"
	ExecWait "$INSTDIR\VBN.EXE /unregserver"

	IfFileExists "$INSTDIR\VBN.dll" +1 +2	
	UnRegDLL "$INSTDIR\VBN.dll"

	Delete "$INSTDIR\VBN.EXE"
	Delete "$INSTDIR\VBN.DLL"
	Delete "$INSTDIR\SINK.EXE"

	Delete "$INSTDIR\VBN.BAS"
	Delete "$INSTDIR\VBN.REG"
	Delete "$INSTDIR\REG.BAK"

	Delete "$INSTDIR\${APPNAME}.txt"
	Delete "$INSTDIR\${APPNAME}-README.txt"
	Delete "$INSTDIR\README.txt"

	Delete "$INSTDIR\ScriptInk.EXE"
	Delete "$INSTDIR\ScriptInk.log"
	Delete "$INSTDIR\ScriptInk.bat"

	IfFileExists "$INSTDIR\LINK.BAK" 0 +3
	IfFileExists "$INSTDIR\LINKLNK.EXE" 0 +2
	Delete "$INSTDIR\LINK.EXE"

	IfFileExists "$INSTDIR\C2.BAK" 0 +3
	IfFileExists "$INSTDIR\C3.EXE" 0 +2
	Delete "$INSTDIR\LINK.EXE"

	Delete "$INSTDIR\Template\Projects\Library1.vbp"
	Delete "$INSTDIR\Template\Projects\Library.vbp"
	Delete "$INSTDIR\Template\Projects\Library1.vbw"
	Delete "$INSTDIR\Template\Projects\Library.vbw"
	Delete "$INSTDIR\Template\Projects\frmMain.frm"

	Delete "$INSTDIR\Template\Projects\API Library.vbw"
	Delete "$INSTDIR\Template\Projects\API Library.vbp"
	Delete "$INSTDIR\Template\Projects\Exports1.def"
	Delete "$INSTDIR\Template\Projects\Exports.def"

	Delete "$INSTDIR\Template\Projects\Module1.bas"
	Delete "$INSTDIR\Template\Projects\Class1.cls"
	Delete "$INSTDIR\Template\Projects\Exports.cls"

	Delete "$INSTDIR\Template\Projects\Module.bas"
	Delete "$INSTDIR\Template\Projects\Class.cls"
	Delete "$INSTDIR\Template\Projects\APIExample.def"
	Delete "$INSTDIR\Template\Projects\Controller.cls"
	Delete "$INSTDIR\Template\Projects\FormHWnd.cls"
	Delete "$INSTDIR\Template\Projects\frmService.frm"
	Delete "$INSTDIR\Template\Projects\LegacyOS.cls"
	Delete "$INSTDIR\Template\Projects\modCommon.bas"
	Delete "$INSTDIR\Template\Projects\modMain.bas"
	Delete "$INSTDIR\Template\Projects\modProcess.bas"
	Delete "$INSTDIR\Template\Projects\modRegistry.bas"
	Delete "$INSTDIR\Template\Projects\modService.bas"
	Delete "$INSTDIR\Template\Projects\modWindow.bas"
	Delete "$INSTDIR\Template\Projects\Service.cls"

	Delete "$INSTDIR\Template\Projects\Service.cls"
	Delete "$INSTDIR\Template\Projects\Service.vbp"
	Delete "$INSTDIR\Template\Projects\Service.vbw"
	Delete "$INSTDIR\Template\Projects\frmService.frx"
	Delete "$INSTDIR\Template\Projects\README.TXT"

	UnRegDLL "$SYSDIR\BasicService.DLL"
	Delete "$SYSDIR\BasicService.DLL"

	UnRegDLL "$INSTDIR\BasicNTAddIn.DLL"
	Delete "$INSTDIR\BasicNTAddIn.DLL"

	Push $0
	ReadRegStr $0 HKCR "batfile\shell\Open\command" ""
	StrCmp $0 "" +1 +2
	WriteRegStr HKCR "batfile\shell\Open\command" "" '"%1" %*'


	ReadRegStr $0 HKCR "batfile\shell\Edit\command" ""
	StrCmp $0 "" +1 +2
	WriteRegStr HKCR "batfile\shell\Edit\command" "" '"$WINDIR\notepad.exe" "%1"'

	ReadRegStr $0 HKCR ".bat" ""
	StrCmp $0 "batfile" +2 +1
	WriteRegStr HKCR ".bat" "" 'batfile'

	ReadRegStr $0 HKCR ".bat\Shell" ""
	StrCmp $0 "Open" +2 +1
	WriteRegStr HKCR ".bat\Shell" "" 'Open'

	ReadRegStr $0 HKCR ".bat\PersistentHandler" ""
	StrCmp $0 "Open" +2 +1
	WriteRegStr HKCR ".bat\PersistentHandler" "" '{5e941d80-bf96-11cd-b579-08002b30bfeb}'

	;ReadRegStr $0 HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "ServicePackState"
	;!insertmacro unServicePacks

	Pop $0

	DeleteRegKey HKCU "Software\Microsoft\Visual Basic\6.0\Addins\BasicNTAddIn.Connect"
	DeleteRegKey HKCU "Software\VB and VBA Program Settings\Notable"

	DeleteRegKey HKCR ".ink"
	DeleteRegKey HKCR "BasicNeotext.Ink"
	DeleteRegKey HKCR ".bat\OpenWithProgids\BasicNeotext.Ink"
	DeleteRegKey HKCR ".bat\Shell\Edit in Notable"
	DeleteRegKey HKCR ".bat\Shell\Exec in Notable"

	IfFileExists "$INSTDIR\LINK.EXE" +1 +2
	Delete "$INSTDIR\LINK.bak"

	IfFileExists "$INSTDIR\C2.EXE" +1 +2
	Delete "$INSTDIR\C2.bak"

	RmDir /r "$INSTDIR\BasicNeotext"


	Delete "$INSTDIR\${APPNAME}-Uninstall.exe"
	Delete "$INSTDIR\Uninstall.exe"

	!insertmacro UninstallLibaries

	DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}"

	DeleteRegValue HKLM "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers" "$INSTDIR\SINK.exe"
	DeleteRegValue HKLM "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers" "$INSTDIR\VBN.exe"
	DeleteRegValue HKLM "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers" "$INSTDIR\LINK.exe"
	DeleteRegValue HKLM "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers" "$INSTDIR\C2.exe"

	DeleteRegValue HKLM "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers" "$INSTDIR\Uninstall.exe"

	DeleteRegKey /ifempty HKLM "SOFTWARE\Neotext"
	DeleteRegKey /ifempty HKCU "SOFTWARE\Neotext"

SectionEnd

