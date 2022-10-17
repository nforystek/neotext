!macro BasicInstallLibrary libname libpath tofolder libver libpar
	StrCmp "${libpar}" "0" +1 +2
	SetOverwrite ifnewer
	StrCmp "${libpar}" "1" +1 +2
	SetOverwrite on
	StrCmp "${libpar}" "2" +1 +2
	SetOverwrite ifdiff
	StrCmp "${libpar}" "3" +1 +2
	SetOverwrite try
	File "${libpath}\${libname}"
	StrCmp "${libpar}" "1" +1 +2
	RegDLL "${tofolder}\${libname}"
	StrCmp "${libpar}" "3" +1 +2
	RegDLL "${tofolder}\${libname}"
!macroend
!macro un.BasicInstallLibrary libname tofolder libver libpar
	StrCmp "${libpar}" "2" +4 +1
	StrCmp "${libpar}" "0" +4 +1
	UnRegDLL "${tofolder}\${libname}"
	StrCmp "${libpar}" "1" +1 +2
	Delete "${tofolder}\${libname}"
!macroend
!macro InstallSharedLibrary libname libpath tofolder libver libpar
	SetOverwrite on
	IfSilent +16 +1
	ClearERrors
	SetOverwrite try
	ReadRegStr $0 HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" ${libname}
	StrCmp $0 "" +1 +3
	StrCpy $0 "0"
	IfFileExists "${tofolder}\${libname}" +1 +2
	StrCpy $0 "1"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" ${libname} $0
	ReadRegStr $0 HKLM "SOFTWARE\Neotext\System" ${libname}
	StrCmp $0 "" +1 +2
	StrCpy $0 "0"
	IntOp $0 $0 + 1
	WriteRegStr HKLM "SOFTWARE\Neotext\System" ${libname} $0
	IfFileExists "${tofolder}\${libName}" +1 +2 
	SetOverwrite ifnewer
	!insertmacro BasicInstallLibrary "${libname}" "${libpath}" "${tofolder}" "${libver}" "${libpar}"
!macroend

!macro un.InstallSharedLibrary libname tofolder libver libpar
	ClearERrors
	ReadRegStr $0 HKLM "SOFTWARE\Neotext\System" ${libname}
	StrCmp $0 "" +1 +2
	StrCpy $0 "1"
	IntOp $0 $0 - 1
	WriteRegStr HKLM "SOFTWARE\Neotext\System" ${libname} $0
	StrCmp $0 "0" +1 +2
	DeleteRegValue HKLM "SOFTWARE\Neotext\System" ${libname}
	ReadRegStr $0 HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" ${libname}
	StrCmp $0 "0" +1 +3
	DeleteRegValue HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" ${libname}
	Delete "${tofolder}\${libname}"
!macroend

!macro InstallSystemLibrary libname libpath tofolder libver libpar
	!insertmacro InstallSharedLibrary "${libname}" "${libpath}" "${tofolder}" "${libver}" "${libpar}"
!macroend

!macro un.InstallSystemLibrary libname tofolder libver libpar
	!insertmacro un.InstallSharedLibrary "${libname}" "${tofolder}" "${libver}" "${libpar}"
!macroend

!macro InstallNormalLibrary libname libpath tofolder libver libpar
	!insertmacro InstallSharedLibrary "${libname}" "${libpath}" "${tofolder}" "${libver}" "${libpar}"
!macroend

!macro un.InstallNormalLibrary libname tofolder libver libpar
	!insertmacro un.InstallSharedLibrary "${libname}" "${tofolder}" "${libver}" "${libpar}"
!macroend


!macro VersionIsEqual libpath libver retrn
	StrCpy ${retrn} "0"
	IfFileExists "${libpath}" +1 +9
	GetDllVersion "${libpath}" $R0 $R1
	IntOp $R2 $R0 / 0x00010000
	IntOp $R3 $R0 & 0x0000FFFF
	IntOp $R4 $R1 / 0x00010000
	IntOp $R5 $R1 & 0x0000FFFF
	StrCmp "$R2.$R3.$R4.$R5" "${libver}" +2
	Goto +2
	StrCpy ${retrn} "1"
!macroend
!macro VersionIsGreater libpath libver retrn
	StrCpy ${retrn} "0"
	IfFileExists "${libpath}" +1 +16
	GetDllVersion "${libpath}" $R0 $R1
	IntOp $R2 $R0 / 0x00010000
	IntOp $R3 $R0 & 0x0000FFFF
	IntOp $R4 $R1 / 0x00010000
	IntOp $R5 $R1 & 0x0000FFFF
	!define VERMAJOR
	!define VERMINOR
	!define VERREVIS
	!define VERSERVI
	!searchparse ".${libver}" "." VERMAJOR "." VERMINOR "." VERREVIS "." VERSERVI
	IntCmp $R2 ${VERMAJOR} +1 +2 +5
	IntCmp $R3 ${VERMINOR} +1 +5 +4
	IntCmp $R4 ${VERREVIS} +1 +4 +3
	IntCmp $R5 ${VERSERVI} +3 +3 +2
	Goto +2
	StrCpy ${retrn} "1"
	!undef VERMAJOR
	!undef VERMINOR
	!undef VERREVIS
	!undef VERSERVI
!macroend
!macro VisualBasicScriptLibraries
	SetOverwrite try
	System::Call 'Ole32::CoFreeUnusedLibraries()'
	IfFileExists "$SYSDIR\msvbvm60.dll" +1 +8
	GetDllVersion "$SYSDIR\msvbvm60.dll" $R0 $R1
	IntOp $R2 $R0 / 0x00010000
	IntOp $R3 $R0 & 0x0000FFFF
	IntOp $R4 $R1 / 0x00010000
	IntOp $R5 $R1 & 0x0000FFFF
	StrCpy $0 "$R2.$R3.$R4.$R5"
	StrCmp $0 "6.0.97.97" +2
	File "${PRODUCT}\Windows\ActiveX\msvbvm60.dll"
	System::Call 'Ole32::CoFreeUnusedLibraries()'
	IfFileExists "$SYSDIR\scrrun.dll" +1 +8
	GetDllVersion "$SYSDIR\scrrun.dll" $R0 $R1
	IntOp $R2 $R0 / 0x00010000
	IntOp $R3 $R0 & 0x0000FFFF
	IntOp $R4 $R1 / 0x00010000
	IntOp $R5 $R1 & 0x0000FFFF
	StrCpy $0 "$R2.$R3.$R4.$R5"
	StrCmp $0 "5.7.0.6000" +2
	File "${PRODUCT}\Windows\ActiveX\scrrun.dll"
	System::Call 'Ole32::CoFreeUnusedLibraries()'
	IfFileExists "$SYSDIR\msscript.ocx" +1 +8
	GetDllVersion "$SYSDIR\msscript.ocx" $R0 $R1
	IntOp $R2 $R0 / 0x00010000
	IntOp $R3 $R0 & 0x0000FFFF
	IntOp $R4 $R1 / 0x00010000
	IntOp $R5 $R1 & 0x0000FFFF
	StrCpy $0 "$R2.$R3.$R4.$R5"
	StrCmp $0 "1.0.0.6000" +2
	File "${PRODUCT}\Windows\ActiveX\msscript.ocx"
!macroend
