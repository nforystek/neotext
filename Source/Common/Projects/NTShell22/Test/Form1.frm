VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12510
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   12510
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5700
      Left            =   840
      TabIndex        =   6
      Top             =   3195
      Width           =   10605
      ExtentX         =   18706
      ExtentY         =   10054
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Reset"
      Height          =   360
      Left            =   405
      TabIndex        =   3
      Top             =   2520
      Width           =   1425
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Final"
      Height          =   405
      Left            =   450
      TabIndex        =   2
      Top             =   1800
      Width           =   1395
   End
   Begin VB.CommandButton Command2 
      Caption         =   "WanIP"
      Height          =   405
      Left            =   465
      TabIndex        =   1
      Top             =   1140
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   390
      Left            =   405
      TabIndex        =   0
      Top             =   465
      Width           =   1410
   End
   Begin VB.Label Label4 
      Height          =   240
      Left            =   1890
      TabIndex        =   8
      Top             =   1875
      Width           =   5595
   End
   Begin VB.Label Label3 
      Height          =   225
      Left            =   1965
      TabIndex        =   7
      Top             =   1200
      Width           =   5130
   End
   Begin VB.Label Label2 
      Height          =   270
      Left            =   1890
      TabIndex        =   5
      Top             =   555
      Width           =   5565
   End
   Begin VB.Label Label1 
      Height          =   180
      Left            =   1935
      TabIndex        =   4
      Top             =   120
      Width           =   4650
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Private Sub Form_Load()
    Dim p As New NTShell22.internet
    
    p.OpenWebsite "http://www.neotext.org"
    
    'p.Run "C:\WINDOWS\SYSTEM32\calc.EXE", "C -f", 0, False
    
    
End Sub

