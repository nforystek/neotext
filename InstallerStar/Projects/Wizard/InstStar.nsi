                     
!include "FileFunc.nsh"
!include "WordFunc.nsh"

Name "%appvalue%"

Icon "%curdir%\InstallStar.ico"

RequestExecutionLevel admin

OutFile "%curdir%\InstStar.exe"

VIProductVersion "1.0.0.0"
VIAddVersionKey "ProductName" "InstStar"
VIAddVersionKey "Comments" "https://www.neotext.org"
VIAddVersionKey "CompanyName" "Neotext"
VIAddVersionKey "LegalTrademarks" "Nicholas Randall Forystek"
VIAddVersionKey "LegalCopyright" "© 1999,2006,2013 by Neotext"
VIAddVersionKey "FileDescription" "Setup program for installation of Neotext.org software titles to run on a majority of Microsoft® Windows(TM) operating systems."
VIAddVersionKey "FileVersion" "1.0.0.1"

var args
Var params
var mode

var shcut
var sclink
var sctrgt
var scargs

Section "-main section"	

SectionEnd

Function .onInit
	${GetParameters} $1
	StrCpy $params $1 2
	StrCmp $params "/C" +1 InstStar
	StrCpy $shcut $1 1024 3
	${WordFind} "$shcut" "|" "+01" $sclink
	${WordFind} "$shcut" "|" "+02" $sctrgt
	${WordFind} "$shcut" "|" "-01" $scargs
	SetShellVarContext all
	CreateShortCut "$sclink" "$sctrgt" "$scargs" "$sctrgt" 0
	Goto InstExit
	InstStar:
	setOutPath "$SYSDIR"
	SetOverwrite ifnewer
	IfFileExists "$SYSDIR\msvbvm60.dll" +1 +10
	GetDllVersion "$SYSDIR\msvbvm60.dll" $R0 $R1
	IntOp $R2 $R0 / 0x00010000
	IntOp $R3 $R0 & 0x0000FFFF
	IntOp $R4 $R1 / 0x00010000
	IntOp $R5 $R1 & 0x0000FFFF
	StrCpy $0 "$R2.$R3.$R4.$R5"
	StrCmp $0 "6.0.97.97" +2
	File "%curdir%\msvbvm60.dll"
	RegDLL "$SYSDIR\msvbvm60.dll"
	SetOverwrite on
	GetTempFileName $0
	Delete $0
	CreateDirectory $0
	SetOutPath "$0"
	StrCpy "$args" "$EXEDIR"
	${GetParameters} $1
	StrCpy $params $1 2
	StrCmp $params "/I" +4 +1
	SetOutPath "$0"
	StrCpy "$args" "$EXEDIR"
	File "%curdir%\Remove.exe"
	File "%curdir%\Wizard.exe"
	File "%curdir%\Manifest.ini"
	IfFileExists "$SYSDIR\msvbvm60.dll" +3 +1
	File "%curdir%\msvbvm60.dll"
	RegDLL "$OUTDIR\msvbvm60.dll"
	StrCpy $mode $1 2 3
	StrCmp $mode "/Q" +1 +3
	ExecWait "$OUTDIR\Wizard.exe /QUIET $args"
	Goto +5
	IfSilent +1 +3
	ExecWait "$OUTDIR\Wizard.exe /SHEEK $args"
	Goto +2
	ExecWait "$OUTDIR\Wizard.exe $args"
	IfFileExists "$OUTDIR\Remove.exe" +1 +2
	Delete "$OUTDIR\Remove.exe"
	Delete "$OUTDIR\Wizard.exe"
	IfFileExists "$OUTDIR\msvbvm60.dll" +1 +3
	UnRegDLL "$OUTDIR\msvbvm60.dll"
	Delete "$OUTDIR\msvbvm60.dll"
	IfFileExists "$OUTDIR\Manifest.ini" +4 +1
	SetRebootFlag false
	RmDir /r "$OUTDIR"
	Goto +6
	SetRebootFlag true
	RmDir /r /REBOOTOK "$OUTDIR"
	Goto +4
	RmDir /r /REBOOTOK "$OUTDIR"
	Goto +8
	IfFileExists "$OUTDIR\*" -2
	StrCmp $params "/I" +6 +1
	IfRebootFlag +1 +5
	IfSilent +1 +2
	SetSilent normal
	MessageBox MB_YESNO|MB_ICONQUESTION "A system restart is required for changes to take effect.$\n$\nDo you want to restart your computer now?" IDYES +1 IDNO +2
	Reboot
	InstExit:
	Quit
FunctionEnd
