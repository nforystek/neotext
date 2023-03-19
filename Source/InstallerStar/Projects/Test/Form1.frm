VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1788
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   6876
   LinkTopic       =   "Form1"
   ScaleHeight     =   1788
   ScaleWidth      =   6876
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Place"
      Height          =   228
      Left            =   2592
      TabIndex        =   6
      Top             =   1416
      Width           =   696
   End
   Begin VB.OptionButton Option2 
      Caption         =   "rev. 1"
      Height          =   204
      Left            =   3336
      TabIndex        =   5
      Top             =   1344
      Width           =   1092
   End
   Begin VB.OptionButton Option1 
      Caption         =   "rev. 0"
      Height          =   192
      Left            =   1764
      TabIndex        =   4
      Top             =   1368
      Value           =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Restore"
      Height          =   492
      Left            =   4056
      TabIndex        =   1
      Top             =   588
      Width           =   2520
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Install"
      Height          =   468
      Left            =   444
      TabIndex        =   0
      Top             =   576
      Width           =   1956
   End
   Begin VB.Label Label2 
      Height          =   240
      Left            =   2832
      TabIndex        =   3
      Top             =   720
      Width           =   888
   End
   Begin VB.Label Label1 
      Caption         =   "Need Reboot"
      Height          =   264
      Left            =   2700
      TabIndex        =   2
      Top             =   396
      Width           =   1176
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN


Private Sub Command1_Click()

    
    Label2.Caption = MoveLibrary("C:\Development\Neotext\InstallerStar\Projects\DLLTest\System32\Project2.dll", "C:\Development\Neotext\InstallerStar\Projects\DLLTest\CommonFiles\Project2.bak", True, "C:\Development\Neotext\InstallerStar\Projects\DLLTest\CommonFiles\Project2.dll") Or _
    MoveLibrary("C:\Development\Neotext\InstallerStar\Projects\DLLTest\CommonFiles\Project2.dll", "C:\Development\Neotext\InstallerStar\Projects\DLLTest\System32\Project2.dll", True, "C:\Development\Neotext\InstallerStar\Projects\DLLTest\CommonFiles\Project2.bak")
    
    FormEnabled
End Sub

Private Sub Command2_Click()
    
    Label2.Caption = MoveLibrary("C:\Development\Neotext\InstallerStar\Projects\DLLTest\System32\Project2.dll", "C:\Development\Neotext\InstallerStar\Projects\DLLTest\CommonFiles\Project2.dll", True, "C:\Development\Neotext\InstallerStar\Projects\DLLTest\CommonFiles\Project2.bak") Or _
    MoveLibrary("C:\Development\Neotext\InstallerStar\Projects\DLLTest\CommonFiles\Project2.bak", "C:\Development\Neotext\InstallerStar\Projects\DLLTest\System32\Project2.dll", True, "C:\Development\Neotext\InstallerStar\Projects\DLLTest\CommonFiles\Project2.dll")
    FormEnabled
End Sub

Private Sub FormEnabled()
    Command1.Enabled = Not PathExists("C:\Development\Neotext\InstallerStar\Projects\DLLTest\CommonFiles\Project2.bak", True)
    Command2.Enabled = PathExists("C:\Development\Neotext\InstallerStar\Projects\DLLTest\CommonFiles\Project2.bak", True)
End Sub

Private Sub Command3_Click()
    If Option1.Value Then
        FileCopy "C:\Development\Neotext\InstallerStar\Projects\DLLTest\PretendInstall\Project0.dll", "C:\Development\Neotext\InstallerStar\Projects\DLLTest\CommonFiles\Project2.dll"
    Else
        FileCopy "C:\Development\Neotext\InstallerStar\Projects\DLLTest\PretendInstall\Project1.dll", "C:\Development\Neotext\InstallerStar\Projects\DLLTest\CommonFiles\Project2.dll"
    End If
    FormEnabled
End Sub

Private Sub Form_Load()
    FormEnabled
End Sub
