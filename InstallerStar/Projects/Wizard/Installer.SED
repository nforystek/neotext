[Version]
Class=IEXPRESS
SEDVersion=3
[Options]
PackagePurpose=InstallApp
ShowInstallProgramWindow=0
HideExtractAnimation=1
UseLongFileName=1
InsideCompressed=0
CAB_FixedSize=0
CAB_ResvCodeSigning=0
RebootMode=N
InstallPrompt=%InstallPrompt%
DisplayLicense=%DisplayLicense%
FinishMessage=%FinishMessage%
TargetName=%TargetName%
FriendlyName=%FriendlyName%
AppLaunched=%AppLaunched%
PostInstallCmd=%PostInstallCmd%
AdminQuietInstCmd=%AdminQuietInstCmd%
UserQuietInstCmd=%UserQuietInstCmd%
SourceFiles=SourceFiles
[Strings]
InstallPrompt=
DisplayLicense=
FinishMessage=
TargetName=%curdir%\Installer.EXE
FriendlyName=%appvalue% Installer
AppLaunched=InstStar.exe /I /S
PostInstallCmd=<None>
AdminQuietInstCmd=InstStar.exe /I /N
UserQuietInstCmd=InstStar.exe /I /Q
FILE0="InstStar.exe"
FILE1=
[SourceFiles]
SourceFiles0=%curdir%\
[SourceFiles0]
%FILE0%=
%FILE1%=
