Attribute VB_Name = "Module1"
Option Explicit

Public Sub Main()
    Dim list As String
 list = list & "C:\Development\Neotext\BasicNeotext\Binary\SINK.exe" & vbCrLf & _
"C:\Development\Neotext\BasicNeotext\Binary\VBN.EXE" & vbCrLf & _
"C:\Development\Neotext\BasicNeotext\Deploy\BasicNeotext v3.0.0.exe" & vbCrLf & _
"C:\Development\Neotext\Blacklawn\Binary\Blacklawn.exe" & vbCrLf & _
"C:\Development\Neotext\Blacklawn\Binary\BlkLServer.exe" & vbCrLf & _
"C:\Development\Neotext\Blacklawn\Deploy\Blacklawn v1.1.0.exe" & vbCrLf & _
"C:\Development\Neotext\Common\Binary\Packages\WebImaging\WebImaging.CAB" & vbCrLf & _
"C:\Development\Neotext\Common\Projects\WebControls\Package\WebControls.CAB" & vbCrLf & _
"C:\Development\Neotext\CrayonStill\Binary\CrayonStall.exe" & vbCrLf & _
"C:\Development\Neotext\CrayonStill\Binary\CrayonStiff.exe" & vbCrLf & _
"C:\Development\Neotext\CrayonStill\Binary\CrayonStill.exe" & vbCrLf & _
"C:\Development\Neotext\CrayonStill\Deploy\CrayonStill v0.0.0.exe" & vbCrLf & _
"C:\Development\Neotext\Creata-Tree\Binary\CreataTree.exe" & vbCrLf & _
"C:\Development\Neotext\Creata-Tree\Deploy\Creata-Tree v3.1.0.exe" & vbCrLf & _
"C:\Development\Neotext\HouseOfGlass\Binary\HouseOfGlass.exe" & vbCrLf & _
"C:\Development\Neotext\HouseOfGlass\Deploy\HouseOfGlass v1.0.0.exe" & vbCrLf & _
"C:\Development\Neotext\IdentAuth\Binary\Ident.exe" & vbCrLf
 list = list & "C:\Development\Neotext\IdentAuth\Binary\Reload.exe" & vbCrLf & _
"C:\Development\Neotext\IdentAuth\Deploy\IdentAuth v7.1.0.exe" & vbCrLf & _
"C:\Development\Neotext\InstallerStar\Binary\Remove.exe" & vbCrLf & _
"C:\Development\Neotext\InstallerStar\Binary\Wizard.exe" & vbCrLf & _
"C:\Development\Neotext\InstallerStar\Binary\InstStar.exe" & vbCrLf & _
"C:\Development\Neotext\InstallerStar\Binary\Installer.exe" & vbCrLf & _
"C:\Development\Neotext\InstallerStar\Binary\Inst0001.cab" & vbCrLf & _
"C:\Development\Neotext\InstallerStar\Binary\Inst0002.cab" & vbCrLf & _
"C:\Development\Neotext\InstallerStar\Binary\Inst0003.cab" & vbCrLf & _
"C:\Development\Neotext\InstallerStar\Binary\Inst0004.cab" & vbCrLf & _
"C:\Development\Neotext\InstallerStar\Binary\Inst0005.cab" & vbCrLf & _
"C:\Development\Neotext\InstallerStar\Binary\Inst0006.cab" & vbCrLf & _
"C:\Development\Neotext\InstallerStar\Binary\Inst0007.cab" & vbCrLf & _
"C:\Development\Neotext\InstallerStar\Binary\Inst0008.cab" & vbCrLf & _
"C:\Development\Neotext\InstallerStar\Binary\Inst0010.cab" & vbCrLf & _
"C:\Development\Neotext\InstallerStar\Binary\Inst0011.cab" & vbCrLf & _
"C:\Development\Neotext\InstallerStar\Binary\Inst0012.cab" & vbCrLf & _
"C:\Development\Neotext\InstallerStar\Binary\Inst0013.cab" & vbCrLf
 list = list & "C:\Development\Neotext\InstallerStar\Binary\Inst0014.cab" & vbCrLf & _
"C:\Development\Neotext\InstallerStar\Binary\Inst0015.cab" & vbCrLf & _
"C:\Development\Neotext\InstallerStar\Binary\Inst0016.cab" & vbCrLf & _
"C:\Development\Neotext\InstallerStar\Binary\Inst0017.cab" & vbCrLf & _
"C:\Development\Neotext\InstallerStar\Binary\Inst0018.cab" & vbCrLf & _
"C:\Development\Neotext\InstallerStar\Binary\Inst0019.cab" & vbCrLf & _
"C:\Development\Neotext\InstallerStar\Binary\Inst0020.cab" & vbCrLf & _
"C:\Development\Neotext\KadPatch\Binary\KadPatch.exe" & vbCrLf & _
"C:\Development\Neotext\Max-FTP\Binary\MaxFTP.exe" & vbCrLf & _
"C:\Development\Neotext\Max-FTP\Binary\MaxIDE.exe" & vbCrLf & _
"C:\Development\Neotext\Max-FTP\Binary\MaxService.exe" & vbCrLf & _
"C:\Development\Neotext\Max-FTP\Binary\MaxUtility.exe" & vbCrLf & _
"C:\Development\Neotext\Max-FTP\Deploy\Max-FTP v6.1.0.exe" & vbCrLf & _
"C:\Development\Neotext\MaxLand\Binary\MaxLandApp.exe" & vbCrLf & _
"C:\Development\Neotext\MaxLand\Deploy\MaxLand v2.2.0.exe" & vbCrLf & _
"C:\Development\Neotext\RemindMe\Binary\RemindMe.exe" & vbCrLf
 list = list & "C:\Development\Neotext\RemindMe\Binary\RmdMeSrv.exe" & vbCrLf & _
"C:\Development\Neotext\RemindMe\Binary\Utility.exe" & vbCrLf & _
"C:\Development\Neotext\RemindMe\Deploy\RemindMe v2.1.0.exe" & vbCrLf & _
"C:\Development\Neotext\Sequencer\Binary\SoSouiXSeq.exe" & vbCrLf & _
"C:\Development\Neotext\Sequencer\Deploy\Sequencer v5.0.5.exe" & vbCrLf & _
"C:\Development\Neotext\Source\Deploy\Source v7.0.0.exe" & vbCrLf & _
"C:\Development\Neotext\To-Doster\Binary\ToDoster.exe" & vbCrLf & _
"C:\Development\Neotext\To-Doster\Deploy\To-Doster v1.2.0.exe" & vbCrLf & _
"C:\Development\Neotext\BasicNeotext\Binary\BasicService.dll" & vbCrLf & _
"C:\Development\Neotext\BasicNeotext\Binary\Template\Projects\Exports.dll" & vbCrLf & _
"C:\Development\Neotext\BasicNeotext\Binary\VBN.DLL" & vbCrLf & _
"C:\Development\Neotext\Common\Binary\NTAdvFTP61.dll" & vbCrLf
 list = list & "C:\Development\Neotext\Common\Binary\NTCipher10.dll" & vbCrLf & _
"C:\Development\Neotext\Common\Binary\NTNodes10.dll" & vbCrLf & _
"C:\Development\Neotext\Common\Binary\NTPopup21.dll" & vbCrLf & _
"C:\Development\Neotext\Common\Binary\NTSchedule20.dll" & vbCrLf & _
"C:\Development\Neotext\Common\Binary\NTService20.dll" & vbCrLf & _
"C:\Development\Neotext\Common\Binary\NTShell22.dll" & vbCrLf & _
"C:\Development\Neotext\Common\Binary\NTSmpFTP30.dll" & vbCrLf & _
"C:\Development\Neotext\Common\Binary\NTSMTP23.dll" & vbCrLf & _
"C:\Development\Neotext\Common\Binary\NTSoSweet.dll" & vbCrLf & _
"C:\Development\Neotext\Common\Binary\NTSound20.dll" & vbCrLf

Dim reg As New Registry



reg.SetValue HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\BasicNeotext\Options", "RestrictList", list





'    Dim txt As String
'    txt = SearchPath("*.dll", , "C:\Development\Neotext", FindAll)
'    Debug.Print txt
    
End Sub
