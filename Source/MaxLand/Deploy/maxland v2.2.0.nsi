
!define APPNAME "MaxLand"
!define APPVER "2.2.0"
!define APPPATH "C:\Development\Neotext"

!include "..\..\Windows\Deploy\Common.nsh"


Icon "..\Media\icon.ico"


Section "MaxLand Game" reqir

	SectionIn RO

	!insertmacro InstallLibaries

	SetOutPath $INSTDIR

	SetOverwrite on

	File "${APPPATH}\${APPNAME}\BInary\MaxLandApp.exe"

	File "${APPPATH}\Windows\Normal\MaxLandLib.dll"
	File "${APPPATH}\${APPNAME}\BInary\Commands.ini"
	File "${APPPATH}\${APPNAME}\BInary\drop.bmp"
	File "${APPPATH}\${APPNAME}\BInary\ctrls.bmp"
	File "${APPPATH}\${APPNAME}\BInary\mouse.cur"
	File "..\Binary\Neotext.org.url"

	IfFileExists "$INSTDIR\MaxLandApp.mdb" +1 +2
	ExecWait "$INSTDIR\MaxLandApp.exe /backupdb"
	File "${APPPATH}\${APPNAME}\BInary\MaxLandApp.mdb"
	IfFileExists "$INSTDIR\MaxLandApp.sql" +1 +2
	ExecWait "$INSTDIR\MaxLandApp.exe /restoredb"

	SetOverwrite ifnewer

	CreateDirectory "$INSTDIR\Sounds"
	SetOutPath "$INSTDIR\Sounds"

	 File "${APPPATH}\${APPNAME}\BInary\Sounds\waterfall.mp3"

	CreateDirectory "$INSTDIR\Levels"
	SetOutPath "$INSTDIR\Levels"

	File "${APPPATH}\${APPNAME}\BInary\Levels\Level1.px"
	File "${APPPATH}\${APPNAME}\BInary\Levels\Tests.px"

	CreateDirectory "$INSTDIR\Models"
	SetOutPath "$INSTDIR\Models"

	 File "${APPPATH}\${APPNAME}\BInary\Models\backfaces.x"
	 File "${APPPATH}\${APPNAME}\BInary\Models\blacklawn.x"
	 File "${APPPATH}\${APPNAME}\BInary\Models\bounds-pawn.x"
	 File "${APPPATH}\${APPNAME}\BInary\Models\bounds-player.x"
	 File "${APPPATH}\${APPNAME}\BInary\Models\bubble.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\bubbles1.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\bubbles2.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\bubbles3.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\bubbles4.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\bubbles5.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\bubbles6.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\bubbles7.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\bubbles8.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\circle.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\debug0.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\debug1.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\debug2.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\debug3.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\debug4.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\decals.x"
	 File "${APPPATH}\${APPNAME}\BInary\Models\diamond.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\diamond.x"
	 File "${APPPATH}\${APPNAME}\BInary\Models\embers.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\giphy01.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\giphy02.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\giphy03.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\giphy04.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\giphy05.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\giphy06.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\giphy07.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\giphy08.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\giphy09.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\giphy10.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\giphy11.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\giphy12.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\giphy13.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\golden.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\granit.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\gravel.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\greekfe.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\ground.x"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hilltops.x"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_earth.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_earth2.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_earth_2.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_earth_b.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_earth_k.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_earth_r.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_fire.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_fire2.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_fire_i1.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_fire_i2.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_fire_i3.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_fire_v1.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_fire_v2.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_fire_v3.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_fire_x1.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_fire_x2.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_fire_x3.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_water.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_water2.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_water_b.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_water_f.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_water_r.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_water_s.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_wind.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_wind2.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_wind_i1.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_wind_i2.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_wind_i3.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_wind_v1.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_wind_v2.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_wind_v3.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_wind_x1.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_wind_x2.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\hud_wind_x3.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\king.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\ladder.x"
	 File "${APPPATH}\${APPNAME}\BInary\Models\liquid.x"
	 File "${APPPATH}\${APPNAME}\BInary\Models\marble.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\marblewall.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\mountain.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\nautical.x"
	 File "${APPPATH}\${APPNAME}\BInary\Models\nn.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\palace.x"
	 File "${APPPATH}\${APPNAME}\BInary\Models\pawn.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\pawnaqua.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\pawnswap.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\player.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\queen.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\restrooms.x"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sky_back.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sky_bottom.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sky_front.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sky_left.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sky_right.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sky_top.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\smoke01.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\smoke02.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\smoke03.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\smoke04.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\smoke05.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\smoke06.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\smoke07.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\smoke08.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\smoke09.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\smoke10.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\smoke11.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\smoke12.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\smoke13.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\smoke14.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\smoke15.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\smoke16.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\smoke17.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\smoke18.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\smoke19.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\smoke20.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\smoke21.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\smoke22.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\smoke23.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\smoke24.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\smoke25.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\spot.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\statues.x"
	 File "${APPPATH}\${APPNAME}\BInary\Models\stonegray.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sundial01.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sundial02.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sundial03.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sundial04.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sundial05.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sundial06.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sundial07.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sundial08.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sundial09.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sundial10.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sundial11.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sundial12.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sundial13.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sundial14.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sundial15.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sundial16.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sundial17.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sundial18.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sundial19.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sundial20.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sundial21.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sundial22.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sundial23.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\sundial24.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\testfield.x"
	 File "${APPPATH}\${APPNAME}\BInary\Models\upfalls.x"
	 File "${APPPATH}\${APPNAME}\BInary\Models\visual-pawn.x"
	 File "${APPPATH}\${APPNAME}\BInary\Models\visual-player.x"
	 File "${APPPATH}\${APPNAME}\BInary\Models\water.x"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterfall0.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterfall1.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterfall2.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterfall3.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterfall4.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterfall5.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterfall6.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterfall7.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterfall8.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterleak01.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterleak02.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterleak03.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterleak04.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterleak05.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterleak06.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterleak07.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterleak08.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterleak09.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterleak10.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterleak11.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterleak12.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterpool.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterpool0.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterpool1.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterpool2.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterpool3.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterpool4.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterpool5.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterpool6.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterpool7.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterpool8.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterpool9.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain01.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain02.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain03.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain04.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain05.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain06.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain07.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain08.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain09.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain10.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain11.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain12.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain13.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain14.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain15.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain16.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain17.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain18.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain19.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain20.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain21.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain22.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain23.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain24.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain25.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain26.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain27.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain28.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain29.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain30.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain31.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain32.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain33.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain34.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain35.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain36.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain37.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain38.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain39.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\waterrain40.bmp"
	 File "${APPPATH}\${APPNAME}\BInary\Models\watertank.bmp"

	CreateDirectory "$SMPROGRAMS\${APPNAME}"
	CreateShortCut "$SMPROGRAMS\${APPNAME}\${APPNAME}.lnk" "$INSTDIR\MaxLandApp.exe" "" "$INSTDIR\MaxLandApp.exe" 0
	CreateShortCut "$SMPROGRAMS\${APPNAME}\Options.lnk" "$INSTDIR\MaxLandApp.exe" "/setup" "$INSTDIR\MaxLandApp.exe" 1

	WriteRegStr HKCU "SOFTWARE\Neotext" "" ""
	WriteRegStr HKLM "SOFTWARE\Neotext" "" ""

	WriteRegStr HKCU "SOFTWARE\Neotext\${APPNAME}" "InstallLic" "BB1C1F0BFACE4FD6D6CFD436DEDFC94CC66D5FDF315233755B44CC705945C6CACBCAC475487A5F7FA3"
	WriteRegStr HKLM "SOFTWARE\Neotext\${APPNAME}" "InstallDir" "$INSTDIR"
	WriteRegStr HKLM "SOFTWARE\Neotext\${APPNAME}" "InstallVer" "${APPNAME} ${APPVER}"
	WriteRegStr HKLM "SOFTWARE\Neotext\${APPNAME}" "InstallFlag" "1"

	WriteRegStr HKLM "SOFTWARE\Neotext\${APPNAME}\Components" "MainProgramFiles" "1"

	WriteRegStr HKLM "SOFTWARE\Neotext\${APPNAME}\Components" "Documentation" "0"

	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayName" "${APPNAME}"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "UninstallString" '"$INSTDIR\Uninstall.exe"'
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "Publisher" "Neotext"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "ProductVersion" '"${APPVER}.0.0"'
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "ModifyPath" '"$INSTDIR\Uninstall.exe"'
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayIcon" '"$INSTDIR\MaxLandApp.exe"'
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "InstallLocation" '"$INSTDIR"'
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "HelpLink" '"http://www.neotext.org/ipub/help/maxland/index.htm"'
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "HelpTelephone" '"+1-952-457-9224"'
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayVersion" '"${APPNAME} v${APPVER}"'
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "VersionMajor" '"2"'
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "VersionMinor" '"2"'

	CreateDirectory "$SMPROGRAMS\${APPNAME}\Support"

	CreateShortCut "$SMPROGRAMS\${APPNAME}\Support\Neotext.org.lnk" "$INSTDIR\Neotext.org.url" "" "$INSTDIR\Neotext.org.url" 0

	WriteRegStr HKLM "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers" "$INSTDIR\Uninstall.exe" "~ RUNASADMIN"
	WriteRegStr HKLM "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers" "$INSTDIR\MaxLandApp.exe" "~ RUNASADMIN"


SectionEnd

Section "Documentation" docu

	CreateShortCut "$SMPROGRAMS\${APPNAME}\Support\Documentation.lnk" "$INSTDIR\Help\Index.htm" "" "$INSTDIR\Help\Index.htm" 0

	SetOverwrite ifnewer

	CreateDirectory "$INSTDIR\Help"
	SetOutPath "$INSTDIR\Help"

	File "${APPPATH}\${APPNAME}\BInary\Help\Console.htm"
	File "${APPPATH}\${APPNAME}\BInary\Help\controls.htm"
	File "${APPPATH}\${APPNAME}\BInary\Help\custom.js"
	File "${APPPATH}\${APPNAME}\BInary\Help\Index.htm"
	File "${APPPATH}\${APPNAME}\BInary\Help\Introduction.htm"
	File "${APPPATH}\${APPNAME}\BInary\Help\screenshot2.png"
	File "${APPPATH}\${APPNAME}\BInary\Help\screenshot3.png"
	File "${APPPATH}\${APPNAME}\BInary\Help\screenshot.png"
	File "${APPPATH}\${APPNAME}\BInary\Help\tree.html"
	File "${APPPATH}\${APPNAME}\BInary\Help\whitepage.htm"

	CreateDirectory "$INSTDIR\Help\Media"
	SetOutPath "$INSTDIR\Help\Media"
	File "${APPPATH}\${APPNAME}\BInary\Help\Media\202.gif"

	CreateDirectory "$INSTDIR\Help\Media\Bullet Icons"
	SetOutPath "$INSTDIR\Help\Media\Bullet Icons"
	File "${APPPATH}\${APPNAME}\BInary\Help\Media\Bullet Icons\bookClosed.gif"
	File "${APPPATH}\${APPNAME}\BInary\Help\Media\Bullet Icons\bookOpen.gif"
	File "${APPPATH}\${APPNAME}\BInary\Help\Media\Bullet Icons\overview.gif"
	File "${APPPATH}\${APPNAME}\BInary\Help\Media\Bullet Icons\topic.gif"
	File "${APPPATH}\${APPNAME}\BInary\Help\Media\Bullet Icons\world.gif"

	CreateDirectory "$INSTDIR\Help\Media\PlusMinus"
	CreateDirectory "$INSTDIR\Help\Media\PlusMinus\Black"
	SetOutPath "$INSTDIR\Help\Media\PlusMinus\Black"
	File "${APPPATH}\${APPNAME}\BInary\Help\Media\PlusMinus\Black\minus.gif"
	File "${APPPATH}\${APPNAME}\BInary\Help\Media\PlusMinus\Black\plus.gif"

	CreateDirectory "$INSTDIR\Help\Media\Treelines"
	CreateDirectory "$INSTDIR\Help\Media\Treelines\Black"
	SetOutPath "$INSTDIR\Help\Media\Treelines\Black"

	File "${APPPATH}\${APPNAME}\BInary\Help\Media\Treelines\Black\btm.gif"
	File "${APPPATH}\${APPNAME}\BInary\Help\Media\Treelines\Black\hline.gif"
	File "${APPPATH}\${APPNAME}\BInary\Help\Media\Treelines\Black\mid.gif"
	File "${APPPATH}\${APPNAME}\BInary\Help\Media\Treelines\Black\top.gif"
	File "${APPPATH}\${APPNAME}\BInary\Help\Media\Treelines\Black\vline.gif"

	WriteRegStr HKLM "SOFTWARE\Neotext\${APPNAME}\Components" "Documentation" "1"

SectionEnd



Function .onInit

	!insertmacro PGPPrivateBlock
	!insertmacro SignedUninstaller

	Push $0
	SectionGetFlags ${docu} $0
	IntOp $0 $0 & !${SF_SELECTED}
	IfFileExists "$INSTDIR\Help\*.*" +1 +2
	IntOp $0 $0 | ${SF_SELECTED}
	SectionSetFlags ${docu} $0

	Pop $0

FunctionEnd

Section "-hidden section"
	Push $0

	SectionGetFlags ${docu} $0
	IntCmp $0 ${SF_SELECTED} +5
	IfFileExists "$INSTDIR\Help\*.*" +1 +2
	RmDir /r "$INSTDIR\Help"
	IfFileExists "$SMPROGRAMS\${APPNAME}\Documentation.lnk" +1 +2
	Delete "$SMPROGRAMS\${APPNAME}\Documentation.lnk"


	Pop $0
SectionEnd


Function .onVerifyInstDir
	IfSilent +5 +1
	ReadRegStr $0 HKLM "SOFTWARE\Neotext\${APPNAME}" "InstallDir"
	StrCmp $0 "" +3
	StrCmp "$0" "$INSTDIR" +2 +1
	Abort
FunctionEnd


Section "Uninstall"



	IfFileExists "$DESKTOP\${APPNAME}.lnk" +1 +2
	Delete "$DESKTOP\${APPNAME}.lnk"

	IfFileExists "$QUICKLAUNCH\${APPNAME}.lnk" +1 +2
	Delete "$QUICKLAUNCH\${APPNAME}.lnk"

	RmDir /r $SMPROGRAMS\${APPNAME}

	RmDir /r "$INSTDIR\Levels"
	RmDir /r "$INSTDIR\Models"
	RmDir /r "$INSTDIR\Sounds"
	RmDir /r "$INSTDIR\Help"

	Delete "$INSTDIR\Commands.ini"
	Delete "$INSTDIR\drop.bmp"
	Delete "$INSTDIR\ctrls.bmp"
	Delete "$INSTDIR\MaxLandApp.exe"
	Delete "$INSTDIR\MaxLandApp.mdb"
	Delete "$INSTDIR\mouse.cur"
	Delete "$INSTDIR\Neotext.org.lnk"
	Delete "$INSTDIR\Neotext.org"
	Delete "$INSTDIR\Neotext.org.url"

	!insertmacro unInstallLibaries

	DeleteRegKey HKLM "SOFTWARE\Neotext\${APPNAME}\Components"

	DeleteRegKey HKCU "SOFTWARE\Neotext\${APPNAME}"
	DeleteRegKey HKLM "SOFTWARE\Neotext\${APPNAME}"
	DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}"

	DeleteRegKey /ifempty HKLM "SOFTWARE\Neotext"

	DeleteRegKey /ifempty HKCU "SOFTWARE\Neotext"

	DeleteRegValue HKLM "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers" "$INSTDIR\Uninstall.exe"

	DeleteRegValue HKLM "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers" "$INSTDIR\MaxLandApp.exe"




	Delete "$INSTDIR\Uninstall.exe"
	RmDir /r "$INSTDIR"
SectionEnd


Function un.onInit

	!insertmacro PGPPrivateBlock

FunctionEnd
