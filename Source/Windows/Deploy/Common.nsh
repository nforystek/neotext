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
LicenseData "${APPPATH}\${APPNAME}\Binary\License.txt"

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

;!include "${NSISDIR}\Include\VB6RunTime.nsh"

;!include "${NSISDIR}\Include\Sections.nsh"

Page license
;Page components
Page directory
Page instfiles

UninstPage uninstConfirm
UninstPage instfiles

Var AlreadyInstalled

!include "${APPPATH}\Windows\Deploy\${APPNAME}.nsi"


;Section "Create Restore Point"
;	IfFileExists "$SYSDIR\wbem\wmic.exe" +1 +2
;	ExecWait "$SYSDIR\wbem\wmic.exe /Namespace:\\root\default Path SystemRestore Call CreateRestorePoint $\"Before Installation of ${APPNAME}$\", 100, 12"
;SectionEnd