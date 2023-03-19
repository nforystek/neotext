copy /y C:\Development\Neotext\InstallerStar\Binary\backup\Manifest.ini C:\Development\Neotext\InstallerStar\Binary\Manifest.ini
copy /y C:\Development\Neotext\InstallerStar\Binary\backup\MSVBVM60.DLL C:\Development\Neotext\InstallerStar\Binary\MSVBVM60.DLL
copy /y C:\Development\Neotext\InstallerStar\Binary\backup\Remove.exe C:\Development\Neotext\InstallerStar\Binary\Remove.exe
copy /y C:\Development\Neotext\InstallerStar\Binary\backup\Wizard.exe C:\Development\Neotext\InstallerStar\Binary\Wizard.exe
"C:\Program Files\NSiS\makensis.exe" C:\Development\Neotext\InstallerStar\Projects\Wizard\InstStar.nsi
C:\Development\Neotext\InstallerStar\Binary\InstStar.exe /I
