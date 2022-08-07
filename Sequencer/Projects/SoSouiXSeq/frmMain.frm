VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SoSouiX.net Sequencer"
   ClientHeight    =   8175
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   12375
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   12375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Record 
      Caption         =   "Recorder"
      Height          =   255
      Left            =   1410
      TabIndex        =   1
      Top             =   75
      Width           =   1065
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000007&
      Caption         =   "S"
      CausesValidation=   0   'False
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   615
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   570
      Width           =   240
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000007&
      Caption         =   "D"
      CausesValidation=   0   'False
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   3
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   570
      Width           =   240
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000007&
      Caption         =   "G"
      CausesValidation=   0   'False
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   6
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   570
      Width           =   240
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000007&
      Caption         =   "H"
      CausesValidation=   0   'False
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   8
      Left            =   2550
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   570
      Width           =   240
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000007&
      Caption         =   "J"
      CausesValidation=   0   'False
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   10
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   570
      Width           =   240
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000007&
      Caption         =   "L"
      CausesValidation=   0   'False
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   13
      Left            =   4035
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   570
      Width           =   240
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000007&
      Caption         =   ";"
      CausesValidation=   0   'False
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   15
      Left            =   4530
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   570
      Width           =   240
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000007&
      Caption         =   "2"
      CausesValidation=   0   'False
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   18
      Left            =   5025
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   570
      Width           =   240
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000007&
      Caption         =   "3"
      CausesValidation=   0   'False
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   20
      Left            =   6015
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   570
      Width           =   240
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000007&
      Caption         =   "4"
      CausesValidation=   0   'False
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   22
      Left            =   6525
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   570
      Width           =   240
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000007&
      Caption         =   "6"
      CausesValidation=   0   'False
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   25
      Left            =   7515
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   570
      Width           =   240
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000007&
      Caption         =   "7"
      CausesValidation=   0   'False
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   27
      Left            =   8010
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   570
      Width           =   240
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000007&
      Caption         =   "9"
      CausesValidation=   0   'False
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   30
      Left            =   8505
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   570
      Width           =   240
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000007&
      Caption         =   "0"
      CausesValidation=   0   'False
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   32
      Left            =   9495
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   570
      Width           =   240
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000007&
      Caption         =   "-"
      CausesValidation=   0   'False
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   34
      Left            =   10005
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   570
      Width           =   240
   End
   Begin MSComDlg.CommonDialog PatternFile 
      Left            =   11655
      Top             =   735
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SoSouiXSeq.Pattern Pattern1 
      Height          =   672
      Index           =   0
      Left            =   36
      TabIndex        =   51
      Top             =   1548
      Width           =   12312
      _ExtentX        =   21722
      _ExtentY        =   1191
   End
   Begin VB.Timer PlayTimer 
      Enabled         =   0   'False
      Left            =   10860
      Top             =   135
   End
   Begin MSComctlLib.Slider Tempo 
      Height          =   270
      Left            =   7695
      TabIndex        =   13
      Top             =   105
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   476
      _Version        =   393216
      TickStyle       =   3
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   0
      Left            =   270
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   92
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   1
      Left            =   570
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   91
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   2
      Left            =   870
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   90
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   3
      Left            =   1170
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   89
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   4
      Left            =   1575
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   88
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   5
      Left            =   1875
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   87
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   6
      Left            =   2175
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   86
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   7
      Left            =   2475
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   85
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   8
      Left            =   2880
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   84
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   9
      Left            =   3180
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   83
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   10
      Left            =   3480
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   82
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   11
      Left            =   3780
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   81
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   12
      Left            =   4185
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   80
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   13
      Left            =   4485
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   79
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   14
      Left            =   4785
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   78
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   15
      Left            =   5085
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   77
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   16
      Left            =   5490
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   76
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   17
      Left            =   5790
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   75
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   18
      Left            =   6090
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   74
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   19
      Left            =   6390
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   73
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   20
      Left            =   6795
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   72
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   21
      Left            =   7095
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   71
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   22
      Left            =   7395
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   70
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   23
      Left            =   7695
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   69
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   24
      Left            =   8100
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   68
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   25
      Left            =   8400
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   67
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   26
      Left            =   8700
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   66
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   27
      Left            =   9000
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   65
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   28
      Left            =   9405
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   64
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   29
      Left            =   9705
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   63
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   30
      Left            =   10005
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   62
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox Position 
      Height          =   135
      Index           =   31
      Left            =   10305
      ScaleHeight     =   75
      ScaleWidth      =   255
      TabIndex        =   61
      Top             =   390
      Width           =   315
   End
   Begin VB.CheckBox Play 
      Caption         =   "Playback"
      Height          =   255
      Left            =   255
      TabIndex        =   0
      Top             =   75
      Width           =   1005
   End
   Begin VB.CheckBox SongMode 
      Caption         =   "Auto Switch"
      Height          =   225
      Left            =   5715
      TabIndex        =   12
      Top             =   90
      Width           =   1260
   End
   Begin VB.OptionButton Pattern 
      Caption         =   "optPattern"
      Height          =   225
      Index           =   9
      Left            =   5460
      TabIndex        =   11
      Top             =   90
      Width           =   195
   End
   Begin VB.OptionButton Pattern 
      Caption         =   "optPattern"
      Height          =   225
      Index           =   8
      Left            =   5220
      TabIndex        =   10
      Top             =   90
      Width           =   195
   End
   Begin VB.OptionButton Pattern 
      Caption         =   "optPattern"
      Height          =   225
      Index           =   7
      Left            =   4980
      TabIndex        =   9
      Top             =   90
      Width           =   195
   End
   Begin VB.OptionButton Pattern 
      Caption         =   "optPattern"
      Height          =   225
      Index           =   6
      Left            =   4740
      TabIndex        =   8
      Top             =   90
      Width           =   195
   End
   Begin VB.OptionButton Pattern 
      Caption         =   "optPattern"
      Height          =   225
      Index           =   5
      Left            =   4500
      TabIndex        =   7
      Top             =   90
      Width           =   195
   End
   Begin VB.OptionButton Pattern 
      Caption         =   "optPattern"
      Height          =   225
      Index           =   4
      Left            =   4260
      TabIndex        =   6
      Top             =   90
      Width           =   195
   End
   Begin VB.OptionButton Pattern 
      Caption         =   "optPattern"
      Height          =   225
      Index           =   3
      Left            =   4020
      TabIndex        =   5
      Top             =   90
      Width           =   195
   End
   Begin VB.OptionButton Pattern 
      Caption         =   "optPattern"
      Height          =   225
      Index           =   2
      Left            =   3780
      TabIndex        =   4
      Top             =   90
      Width           =   195
   End
   Begin VB.OptionButton Pattern 
      Caption         =   "optPattern"
      Height          =   225
      Index           =   1
      Left            =   3540
      TabIndex        =   3
      Top             =   90
      Width           =   195
   End
   Begin VB.OptionButton Pattern 
      Caption         =   "optPattern"
      Height          =   225
      Index           =   0
      Left            =   3300
      TabIndex        =   2
      Top             =   90
      Value           =   -1  'True
      Width           =   195
   End
   Begin SoSouiXSeq.Pattern Pattern1 
      Height          =   672
      Index           =   1
      Left            =   96
      TabIndex        =   52
      Top             =   2028
      Width           =   12312
      _ExtentX        =   21722
      _ExtentY        =   1191
   End
   Begin SoSouiXSeq.Pattern Pattern1 
      Height          =   672
      Index           =   2
      Left            =   276
      TabIndex        =   53
      Top             =   2568
      Width           =   12312
      _ExtentX        =   21722
      _ExtentY        =   1191
   End
   Begin SoSouiXSeq.Pattern Pattern1 
      Height          =   672
      Index           =   3
      Left            =   336
      TabIndex        =   54
      Top             =   3168
      Width           =   12312
      _ExtentX        =   21722
      _ExtentY        =   1191
   End
   Begin SoSouiXSeq.Pattern Pattern1 
      Height          =   672
      Index           =   4
      Left            =   276
      TabIndex        =   55
      Top             =   3828
      Width           =   12312
      _ExtentX        =   21722
      _ExtentY        =   1191
   End
   Begin SoSouiXSeq.Pattern Pattern1 
      Height          =   672
      Index           =   5
      Left            =   336
      TabIndex        =   56
      Top             =   4488
      Width           =   12312
      _ExtentX        =   21722
      _ExtentY        =   1191
   End
   Begin SoSouiXSeq.Pattern Pattern1 
      Height          =   672
      Index           =   6
      Left            =   96
      TabIndex        =   57
      Top             =   5208
      Width           =   12312
      _ExtentX        =   21722
      _ExtentY        =   1191
   End
   Begin SoSouiXSeq.Pattern Pattern1 
      Height          =   672
      Index           =   7
      Left            =   216
      TabIndex        =   58
      Top             =   5748
      Width           =   12312
      _ExtentX        =   21722
      _ExtentY        =   1191
   End
   Begin SoSouiXSeq.Pattern Pattern1 
      Height          =   672
      Index           =   8
      Left            =   396
      TabIndex        =   59
      Top             =   6048
      Width           =   12312
      _ExtentX        =   21722
      _ExtentY        =   1191
   End
   Begin SoSouiXSeq.Pattern Pattern1 
      Height          =   672
      Index           =   9
      Left            =   516
      TabIndex        =   60
      Top             =   6552
      Width           =   12312
      _ExtentX        =   21722
      _ExtentY        =   1191
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "["
      CausesValidation=   0   'False
      Height          =   630
      Index           =   35
      Left            =   10140
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   570
      Width           =   480
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "P"
      CausesValidation=   0   'False
      Height          =   630
      Index           =   33
      Left            =   9645
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   570
      Width           =   480
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "O"
      CausesValidation=   0   'False
      Height          =   630
      Index           =   31
      Left            =   9150
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   570
      Width           =   480
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "I"
      CausesValidation=   0   'False
      Height          =   630
      Index           =   29
      Left            =   8655
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   570
      Width           =   480
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "U"
      CausesValidation=   0   'False
      Height          =   630
      Index           =   28
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   570
      Width           =   480
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "Y"
      CausesValidation=   0   'False
      Height          =   630
      Index           =   26
      Left            =   7665
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   570
      Width           =   480
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "T"
      CausesValidation=   0   'False
      Height          =   630
      Index           =   24
      Left            =   7170
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   570
      Width           =   480
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "R"
      CausesValidation=   0   'False
      Height          =   630
      Index           =   23
      Left            =   6675
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   570
      Width           =   480
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "E"
      CausesValidation=   0   'False
      Height          =   630
      Index           =   21
      Left            =   6180
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   570
      Width           =   480
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "W"
      CausesValidation=   0   'False
      Height          =   630
      Index           =   19
      Left            =   5685
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   570
      Width           =   480
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "Q"
      CausesValidation=   0   'False
      Height          =   630
      Index           =   17
      Left            =   5190
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   570
      Width           =   480
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "/"
      CausesValidation=   0   'False
      Height          =   630
      Index           =   16
      Left            =   4695
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   570
      Width           =   480
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "."
      CausesValidation=   0   'False
      Height          =   630
      Index           =   14
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   570
      Width           =   480
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   ","
      CausesValidation=   0   'False
      Height          =   630
      Index           =   12
      Left            =   3705
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   570
      Width           =   480
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "M"
      CausesValidation=   0   'False
      Height          =   630
      Index           =   11
      Left            =   3210
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   570
      Width           =   480
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "N"
      CausesValidation=   0   'False
      Height          =   630
      Index           =   9
      Left            =   2715
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   570
      Width           =   480
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "B"
      CausesValidation=   0   'False
      Height          =   630
      Index           =   7
      Left            =   2220
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   570
      Width           =   480
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "V"
      CausesValidation=   0   'False
      Height          =   630
      Index           =   5
      Left            =   1725
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   570
      Width           =   480
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "C"
      CausesValidation=   0   'False
      Height          =   630
      Index           =   4
      Left            =   1230
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   570
      Width           =   480
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "X"
      CausesValidation=   0   'False
      Height          =   630
      Index           =   2
      Left            =   750
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   570
      Width           =   480
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "Z"
      CausesValidation=   0   'False
      Height          =   630
      Index           =   0
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   570
      Width           =   480
   End
   Begin MSComctlLib.Slider keyVol 
      Height          =   264
      Left            =   372
      TabIndex        =   50
      Top             =   1272
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   476
      _Version        =   393216
      Max             =   127
      SelStart        =   127
      TickStyle       =   3
      Value           =   127
   End
   Begin VB.Label Label5 
      Caption         =   "-"
      Height          =   180
      Left            =   312
      TabIndex        =   97
      Top             =   1260
      Width           =   180
   End
   Begin VB.Label Label4 
      Caption         =   "+"
      Height          =   180
      Left            =   10512
      TabIndex        =   96
      Top             =   1272
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   10995
      Picture         =   "frmMain.frx":038A
      Top             =   180
      Width           =   1020
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   195
      Left            =   10110
      TabIndex        =   95
      Top             =   105
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Tempo"
      Height          =   195
      Left            =   7170
      TabIndex        =   94
      Top             =   90
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Pattern"
      Height          =   195
      Left            =   2640
      TabIndex        =   93
      Top             =   75
      Width           =   555
   End
   Begin VB.Menu mnSequence 
      Caption         =   "&Sequence"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuDash389 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As.."
      End
      Begin VB.Menu mnuDash32873 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuPiano 
      Caption         =   "Se&tup"
      Begin VB.Menu mnuSetup 
         Caption         =   "&Device"
         Begin VB.Menu mnuDevice 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu mnuDevice 
            Caption         =   ""
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDevice 
            Caption         =   ""
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDevice 
            Caption         =   ""
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDevice 
            Caption         =   ""
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDevice 
            Caption         =   ""
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDevice 
            Caption         =   ""
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDevice 
            Caption         =   ""
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDevice 
            Caption         =   ""
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDevice 
            Caption         =   ""
            Index           =   9
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDevice 
            Caption         =   ""
            Index           =   10
            Visible         =   0   'False
         End
      End
      Begin VB.Menu ChannelOption 
         Caption         =   "&Channel"
         Begin VB.Menu Chan 
            Caption         =   "1"
            Index           =   0
         End
         Begin VB.Menu Chan 
            Caption         =   "2"
            Index           =   1
         End
         Begin VB.Menu Chan 
            Caption         =   "3"
            Index           =   2
         End
         Begin VB.Menu Chan 
            Caption         =   "4"
            Index           =   3
         End
         Begin VB.Menu Chan 
            Caption         =   "5"
            Index           =   4
         End
         Begin VB.Menu Chan 
            Caption         =   "6"
            Index           =   5
         End
         Begin VB.Menu Chan 
            Caption         =   "7"
            Index           =   6
         End
         Begin VB.Menu Chan 
            Caption         =   "8"
            Index           =   7
         End
         Begin VB.Menu Chan 
            Caption         =   "9"
            Index           =   8
         End
         Begin VB.Menu Chan 
            Caption         =   "10"
            Index           =   9
         End
         Begin VB.Menu Chan 
            Caption         =   "11"
            Index           =   10
         End
         Begin VB.Menu Chan 
            Caption         =   "12"
            Index           =   11
         End
         Begin VB.Menu Chan 
            Caption         =   "13"
            Index           =   12
         End
         Begin VB.Menu Chan 
            Caption         =   "14"
            Index           =   13
         End
         Begin VB.Menu Chan 
            Caption         =   "15"
            Index           =   14
         End
         Begin VB.Menu Chan 
            Caption         =   "16"
            Index           =   15
         End
      End
      Begin VB.Menu base 
         Caption         =   "&Base note"
      End
      Begin VB.Menu nuDhadaeoh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMetronome 
         Caption         =   "&Metronome"
      End
   End
   Begin VB.Menu mnuMixer 
      Caption         =   "Mi&xer"
      Begin VB.Menu mnuAddWave 
         Caption         =   "Add &Wave Track"
      End
      Begin VB.Menu mnuAddMidi 
         Caption         =   "Add &MIDI Track"
      End
      Begin VB.Menu mnuPianoAdd 
         Caption         =   "Add a &Recorder"
      End
      Begin VB.Menu mnuAddSynth 
         Caption         =   "Add a Synthesizer"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDash45 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDrumKit 
         Caption         =   "Template &Drum Kit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private Const INVALID_NOTE = -1     ' Code for keyboard keys that we don't handle

Private pSelPattern As Integer

'*************************************************************
Private channel As Integer       ' midi output channel
Private BaseNote As Integer      ' the first note on our "piano"
'*************************************************************

Private mHertz As Long
Private mLength As Long

Public ElapsedTiming As Single
Private CancelNotes As Boolean

'*************************************************************
Public AtBeat As Integer
'*************************************************************
Public Property Get SelPattern() As Integer
    SelPattern = pSelPattern
End Property
Public Property Let SelPattern(ByVal newval As Integer)
    pSelPattern = newval
    Dim cnt As Integer
    For cnt = 1 To 9

        Pattern1(cnt).Visible = (cnt = newval)
        If Pattern1(cnt).Visible Then
        
        End If
    Next
    
End Property
' Set the value for the starting note of the piano
Private Sub base_Click()
   Dim s As String
   Dim i As Integer
   s = InputBox("Enter the new base note for the keyboard (0 - 92)", "Base note", CStr(BaseNote))
   If IsNumeric(s) Then
      i = CInt(s)
      If (i >= 0 And i < 93) Then
         BaseNote = i
      End If
   End If
End Sub



' Select the midi output channel
Private Sub chan_Click(Index As Integer)
   Chan(channel).Checked = False
   channel = Index
   Chan(channel).Checked = True
End Sub


' If user presses a keyboard key, start the corresponding midi note
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not CancelNotes Then
        StartNoteEx NoteFromKey(KeyCode), keyVol.Value, BaseNote, channel
        Debug.Print "Form_KeyDown"
    End If
End Sub

' If user lifts a keyboard key, stop the corresponding midi note
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not CancelNotes Then
        StopNoteEx NoteFromKey(KeyCode), BaseNote, channel
        Debug.Print "Form_KeyUp"
    End If
End Sub

Private Sub key_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    If Not CancelNotes Then
        StopNoteEx Index, BaseNote, channel
        Debug.Print "key_DragDrop"
    End If
End Sub

Private Sub key_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    If Not CancelNotes Then
        StartNoteEx Index, keyVol.Value, BaseNote, channel
        Debug.Print "key_DragOver"
    End If
End Sub

Private Sub key_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not CancelNotes Then
        StartNoteEx NoteFromKey(KeyCode), keyVol.Value, BaseNote, channel
        Debug.Print "Form_KeyDown"
    End If
End Sub

Private Sub key_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not CancelNotes Then
        StopNoteEx NoteFromKey(KeyCode), BaseNote, channel
        Debug.Print "Form_KeyUp"
    End If
End Sub

' Start a note when user click on it
Private Sub key_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not CancelNotes Then
        StartNoteEx Index, keyVol.Value, BaseNote, channel
        Debug.Print "key_MouseDown"
    End If
End Sub

Private Sub key_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not CancelNotes Then
        StopNoteEx Index, BaseNote, channel
        Debug.Print "key_MouseUp"
    End If
End Sub


' Press the button and send midi start event
Public Sub StartNoteEx(ByVal Index As Integer, ByVal volume As Integer, Optional ByVal BaseNote1 As Integer, Optional ByVal channel1 As Integer, Optional Recorder As Boolean = True)
    
    If Record.Value And Recorder Then
        Dim X
        For Each X In Pattern1(SelPattern).Mixers
            If Not X.Mute Then X.MarkPianoNote True, Index, volume, BaseNote1, channel1
        Next
    End If

    If (Index = INVALID_NOTE) Then
        Exit Sub
    End If

    If (key(Index).Value = 1) Then
        Exit Sub
    End If

        key(Index).Value = 1
        midimsg = &H90 + ((BaseNote1 + Index) * &H100) + (volume * &H10000) + channel1
        midiOutShortMsg hmidi, midimsg

    DoEvents
End Sub

' Raise the button and send midi stop event
Public Sub StopNoteEx(ByVal Index As Integer, Optional ByVal BaseNote1 As Integer, Optional ByVal channel1 As Integer, Optional Recorder As Boolean = True)
    
    If Record.Value And Recorder Then
        Dim X
        For Each X In Pattern1(SelPattern).Mixers
            If Not X.Mute Then X.MarkPianoNote False, Index, 0, BaseNote1, channel1
        Next
    End If
    If (Index = INVALID_NOTE) Then
        Exit Sub
    End If

        key(Index).Value = 0
        midimsg = &H80 + ((BaseNote1 + Index) * &H100) + channel1
        midiOutShortMsg hmidi, midimsg

    DoEvents
End Sub

' Get the note corresponding to a keyboard key
Private Function NoteFromKey(key As Integer)

    NoteFromKey = INVALID_NOTE
    Select Case key
        Case vbKeyZ
            NoteFromKey = 0
        Case vbKeyS
            NoteFromKey = 1
        Case vbKeyX
            NoteFromKey = 2
        Case vbKeyD
            NoteFromKey = 3
        Case vbKeyC
            NoteFromKey = 4
        Case vbKeyV
            NoteFromKey = 5
        Case vbKeyG
            NoteFromKey = 6
        Case vbKeyB
            NoteFromKey = 7
        Case vbKeyH
            NoteFromKey = 8
        Case vbKeyN
            NoteFromKey = 9
        Case vbKeyJ
            NoteFromKey = 10
        Case vbKeyM
            NoteFromKey = 11
        Case 188 ' comma
            NoteFromKey = 12
        Case vbKeyL
            NoteFromKey = 13
        Case 190 ' period
            NoteFromKey = 14
        Case 186 ' semicolon
            NoteFromKey = 15
        Case 191 ' forward slash
            NoteFromKey = 16
        Case vbKeyQ
            NoteFromKey = 17
        Case vbKey2
            NoteFromKey = 18
        Case vbKeyW
            NoteFromKey = 19
        Case vbKey3
            NoteFromKey = 20
        Case vbKeyE
            NoteFromKey = 21
        Case vbKey4
            NoteFromKey = 22
        Case vbKeyR
            NoteFromKey = 23
        Case vbKeyT
            NoteFromKey = 24
        Case vbKey6
            NoteFromKey = 25
        Case vbKeyY
            NoteFromKey = 26
        Case vbKey7
            NoteFromKey = 27
        Case vbKeyU
            NoteFromKey = 28
        Case vbKeyI
            NoteFromKey = 29
        Case vbKey9
            NoteFromKey = 30
        Case vbKeyO
            NoteFromKey = 31
        Case vbKey0
            NoteFromKey = 32
        Case vbKeyP
            NoteFromKey = 33
        Case 189 'dash
            NoteFromKey = 34
        Case 219
            NoteFromKey = 35
   End Select
   'If NoteFromKey = INVALID_NOTE And IsNumeric(Me.Tag) Then NoteFromKey = Me.Tag
End Function
Private Function LabelKeys()
    
    key(0).Caption = "Z" & vbCrLf & "1"
    key(1).Caption = "S" & vbCrLf & "2"
    key(2).Caption = "X" & vbCrLf & "3"
    key(3).Caption = "D" & vbCrLf & "4"
    key(4).Caption = "C" & vbCrLf & "5"
    key(5).Caption = "V" & vbCrLf & "6"
    key(6).Caption = "G" & vbCrLf & "7"
    key(7).Caption = "B" & vbCrLf & "8"
    key(8).Caption = "H" & vbCrLf & "9"
    key(9).Caption = "N" & vbCrLf & "10"
    key(10).Caption = "J" & vbCrLf & "11"
    key(11).Caption = "M" & vbCrLf & "12"
    key(12).Caption = "," & vbCrLf & "13"
    key(13).Caption = "L" & vbCrLf & "14"
    key(14).Caption = "." & vbCrLf & "15"
    key(15).Caption = ";" & vbCrLf & "16"
    key(16).Caption = "/" & vbCrLf & "17"
    key(17).Caption = "Q" & vbCrLf & "18"
    key(18).Caption = "2" & vbCrLf & "19"
    key(19).Caption = "W" & vbCrLf & "20"
    key(20).Caption = "3" & vbCrLf & "21"
    key(21).Caption = "E" & vbCrLf & "22"
    key(22).Caption = "4" & vbCrLf & "23"
    key(23).Caption = "R" & vbCrLf & "24"
    key(24).Caption = "T" & vbCrLf & "25"
    key(25).Caption = "6" & vbCrLf & "26"
    key(26).Caption = "Y" & vbCrLf & "27"
    key(27).Caption = "7" & vbCrLf & "28"
    key(28).Caption = "U" & vbCrLf & "29"
    key(29).Caption = "I" & vbCrLf & "30"
    key(30).Caption = "9" & vbCrLf & "31"
    key(31).Caption = "O" & vbCrLf & "32"
    key(32).Caption = "0" & vbCrLf & "33"
    key(33).Caption = "P" & vbCrLf & "34"
    key(34).Caption = "-" & vbCrLf & "35"
    key(35).Caption = "[" & vbCrLf & "36"
    
End Function

Public Sub ClearControls(ByRef ControlArray, Optional ByVal Saftey As Integer = 0)
    Dim cnt
    For Each cnt In ControlArray
        If cnt.Index > Saftey Then Unload cnt
    Next
    DoEvents
End Sub
Public Sub ClearMixers()
    CancelNotes = True
    
    Dim cnt As Integer
    Dim X
    
    For cnt = 0 To 9
        Do Until Pattern1(cnt).Mixers.Count = 1
            Pattern1(cnt).RemoveMixer 1
        Loop
    Next
    
    For cnt = 1 To 9
        Pattern1(cnt).top = Pattern1(0).top
        Pattern1(cnt).Left = Pattern1(0).Left
        Pattern1(cnt).Visible = False
    Next
    Pattern(0).Value = True
    SelPattern = 0
    Tempo.Max = 500
    Tempo.Value = 92
    Record.Value = False
    Play.Value = False
    SongMode.Value = False
    
    
   ' Set the default channel
   channel = 0
   Chan(channel).Checked = True
   
   ' Set the base note
   BaseNote = 60

    Me.Caption = "SoSouiX.net Sequencer"
    
    CancelNotes = False
End Sub

Private Function InitializeMIDI() As Long
   Dim i As Long
   Dim caps As MIDIOUTCAPS
   
   ' Set the first device as midi mapper
   mnuDevice(0).Caption = "MIDI Mapper"
   mnuDevice(0).Visible = True
   mnuDevice(0).Enabled = True
   
   ' Get the rest of the midi devices
   numDevices = midiOutGetNumDevs()
   For i = 0 To (numDevices - 1)
      midiOutGetDevCaps i, caps, Len(caps)
      mnuDevice(i + 1).Caption = caps.szPname
      mnuDevice(i + 1).Visible = True
      mnuDevice(i + 1).Enabled = True
   Next
   
   ' Select the MIDI Mapper as the default device
   mnuDevice_Click (0)
   
   InitializeMIDI = (rc <> 0)
End Function

Private Sub Form_Load()

    On Error GoTo failedsound
    
    Set directSound = directX.DirectSoundCreate("")  'create a DSound object
    directSound.SetCooperativeLevel Me.hwnd, DSSCL_NORMAL  'DSSCL_PRIORITY
            
    mnuNew_Click

    Me.Show
    DoEvents
    
    'LabelKeys
    Me.SetFocus
    
failedsound:
    If Err Or InitializeMIDI <> 0 Then
        MsgBox "Unable to initialize DirectX. Please check that your computer is setup properly for sound and that DirectX is installed.", vbCritical, "Sound Failure"
        Err.Clear
        End
    End If
    
    Dim X
    For Each X In Pattern1
        X.AddMixer "P"
    Next
    
    On Error GoTo 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.Height = Pattern1(SelPattern).top + Pattern1(SelPattern).Height + (Screen.TwipsPerPixelY * 3)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Proc = GetWindowLong(frmMain.HWnd, GWL_WNDPROC)
    Set directSound = Nothing
    Set directX = Nothing
    
   rc = midiOutClose(hmidi)
       
       End
End Sub


Private Sub keyVol_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not CancelNotes Then
        StartNoteEx NoteFromKey(KeyCode), keyVol.Value, BaseNote, channel
        Debug.Print "Form_KeyDown"
    End If
End Sub

Private Sub keyVol_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not CancelNotes Then
        StopNoteEx NoteFromKey(KeyCode), BaseNote, channel
        Debug.Print "Form_KeyUp"
    End If
End Sub

'Private Sub keyvol_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    If Not CancelNotes Then
'        StartNoteEx NoteFromKey(KeyCode), keyVol.Value, BaseNote, channel
'        Debug.Print "Form_KeyDown"
'    End If
'End Sub
'
'Private Sub keyvol_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'    If Not CancelNotes Then
'        StopNoteEx NoteFromKey(KeyCode), BaseNote, channel
'        Debug.Print "Form_KeyUp"
'    End If
'End Sub


Private Sub mnSequence_Click()
    
    mnuOpen.Enabled = Not PlayTimer.Enabled
    mnuSave.Enabled = Not PlayTimer.Enabled
    mnuNew.Enabled = Not PlayTimer.Enabled
    mnuSaveAs.Enabled = Not PlayTimer.Enabled
    
End Sub

Private Sub mnuAddMidi_Click()
    Dim X
    For Each X In Pattern1
        X.AddMixer "M"
    Next
End Sub

Private Sub mnuAddSynth_Click()
        Dim X
    For Each X In Pattern1
        X.AddMixer "S"
    Next
End Sub

Private Sub mnuAddWave_Click()
    Dim X
    For Each X In Pattern1
        X.AddMixer "W"
    Next
End Sub


Private Sub mnuMetronome_Click()
    mnuMetronome.Checked = Not mnuMetronome.Checked
    
End Sub

Private Sub mnuPianoAdd_Click()
    Dim X
    For Each X In Pattern1
        X.AddMixer "P"
    Next

End Sub

Private Sub mnuDevice_Click(Index As Integer)
   mnuDevice(curDevice + 1).Checked = False
   mnuDevice(Index).Checked = True
   curDevice = Index - 1
   rc = midiOutClose(hmidi)
   rc = midiOutOpen(hmidi, curDevice, 0, 0, 0)
   
End Sub

Private Sub mnuDrumKit_Click()
    Dim X
    For Each X In Pattern1
        X.AddMixer "W"
        X.Mixers(X.Mixers.Count - 1).LoadWaveFileName AppPath & "WAV\Drumkit\Kick.wav"
        X.AddMixer "W"
        X.Mixers(X.Mixers.Count - 1).LoadWaveFileName AppPath & "WAV\Drumkit\Double.wav"
        X.AddMixer "W"
        X.Mixers(X.Mixers.Count - 1).LoadWaveFileName AppPath & "WAV\Drumkit\Snare.wav"

        X.AddMixer "W"
        X.Mixers(X.Mixers.Count - 1).LoadWaveFileName AppPath & "WAV\Drumkit\Step.wav"
        X.AddMixer "W"
        X.Mixers(X.Mixers.Count - 1).LoadWaveFileName AppPath & "WAV\Drumkit\RimShot.wav"
        
        X.AddMixer "W"
        X.Mixers(X.Mixers.Count - 1).LoadWaveFileName AppPath & "WAV\Drumkit\Crash.wav"
        X.AddMixer "W"
        X.Mixers(X.Mixers.Count - 1).LoadWaveFileName AppPath & "WAV\Drumkit\Ride.wav"
        
        X.AddMixer "W"
        X.Mixers(X.Mixers.Count - 1).LoadWaveFileName AppPath & "WAV\Drumkit\HighTom.wav"
        X.AddMixer "W"
        X.Mixers(X.Mixers.Count - 1).LoadWaveFileName AppPath & "WAV\Drumkit\MidTom.wav"
        X.AddMixer "W"
        X.Mixers(X.Mixers.Count - 1).LoadWaveFileName AppPath & "WAV\Drumkit\LowTom.wav"




    Next
End Sub

Private Sub mnuExit_Click()
    Unload Me
    End
End Sub

Private Sub mnuNew_Click()
    ClearMixers
    Me.Caption = "SoSouiX.net Sequencer - [New Pattern]"
End Sub

Private Sub mnuOpen_Click()
    On Error Resume Next
    CancelNotes = True
    If PatternFile.FileName <> "" Then
        PatternFile.InitDir = Left(PatternFile.FileName, InStrRev(PatternFile.FileName, "\") - 1)
    End If
    PatternFile.CancelError = True
    PatternFile.Filter = "Pattern File (*.pattern)|*.pattern|All Files (*.*)|*.*"
    PatternFile.FilterIndex = 1
    PatternFile.DialogTitle = "Load Pattern"
    PatternFile.ShowOpen
    If Err = 0 Then
        LoadPattern PatternFile.FileName
    End If
    CancelNotes = False
End Sub

Private Sub mnuSave_Click()
    CancelNotes = True
    If PathExists(PatternFile.FileName, True) Then
        SavePattern PatternFile.FileName
    Else
        mnuSaveAs_Click
    End If
    CancelNotes = False
End Sub

Private Sub mnuSaveAs_Click()
    On Error Resume Next
    CancelNotes = True
    If PatternFile.FileName <> "" Then
        PatternFile.InitDir = Left(PatternFile.FileName, InStrRev(PatternFile.FileName, "\") - 1)
    End If
    PatternFile.CancelError = True
    PatternFile.Filter = "Pattern File (*.pattern)|*.pattern|All Files (*.*)|*.*"
    PatternFile.FilterIndex = 1
    PatternFile.DialogTitle = "Save Pattern"
    If PatternFile.FileName = "" Then PatternFile.FileName = "New Pattern"
    PatternFile.ShowSave
    If Err = 0 Then
        SavePattern PatternFile.FileName
    End If
    CancelNotes = False
End Sub


Private Sub Pattern_Click(Index As Integer)
    Dim cnt As Integer
    For cnt = 0 To 9
        Pattern1(cnt).Visible = False
    Next
    Pattern1(Index).Visible = True
    SelPattern = Index
    Form_Resize
End Sub


Private Sub Pattern1_Resize(Index As Integer)
    Form_Resize
End Sub

Private Sub Play_Click()
   
    If Play.Value = 0 Then
        PlayTimer.Enabled = False
        Dim cnt As Integer
        For cnt = 0 To 31
            Position(cnt).BackColor = vbButtonFace
        Next
        Dim X
'        If SongMode.Value = 0 Then
            For Each X In Pattern1(SelPattern).Mixers
                If X.MixerType = "P" Then
                    X.StopPiano
                End If
            Next
'        End If
        ElapsedTiming = Timer
    Else
        AtBeat = 0
        PlayTimer.Enabled = True
        keyVol.SetFocus
    End If
End Sub

Private Sub PlayTimer_Timer()
    ElapsedTiming = Timer

    
    Dim X
    
    If AtBeat = 0 Then

        Position(31).BackColor = vbButtonFace

        For Each X In Pattern1(SelPattern).Mixers
            If X.MixerType = "P" Then
                X.StopPiano
                If Not Record.Value Then
                    X.PlayPiano
                ElseIf Record.Value Then
                    X.RecordPiano
                End If
            End If
        Next

    Else
        Position(AtBeat - 1).BackColor = vbButtonFace
    End If
    
    Position(AtBeat).BackColor = vbYellow

    For Each X In Pattern1(SelPattern).Mixers
        If X.Mute.Value = 1 Then
            '0-31
            If X.Beat(AtBeat).Value Then
                Select Case X.Beat(AtBeat).Caption
                    Case "S"
                        StopNote X.channel, X.BaseNote
                        StartNote X.channel, X.BaseNote, X.volume.Value
                        StopNote X.channel, X.BaseNote
                    Case "L"
                        StopNote X.channel, X.BaseNote
                        StartNote X.channel, X.BaseNote, X.volume.Value
                    Case "W"
                        X.PlaySound 0
                        'MsgBox beatVol(cnt).Value / beatVol(cnt).Max
                End Select
            End If
        End If
    Next
    
    
    If mnuMetronome.Checked Then
        If AtBeat Mod 4 = 0 Then Beep 6000, 1
    
    End If
    AtBeat = AtBeat + 1
    If AtBeat >= 32 Then
        AtBeat = 0
        If SongMode.Value Then
            If SelPattern = 9 Then
                Pattern(0).Value = True
            Else
                Pattern(SelPattern + 1).Value = True
            End If
            
            'stop/start
            
        End If
    End If
    
End Sub

Private Sub Record_Click()
    keyVol.SetFocus
'    If Record.Value Then
'        StartRecording
'    Else
'        StopRecording
'    End If
End Sub

Private Sub Tempo_Change()
    PlayTimer.Interval = Tempo.Value
    Label3.Caption = Tempo.Value
End Sub

Public Sub SavePattern(ByVal FileName As String)
    
    Dim X, cnt As Integer
    Dim FileNum As Integer
    Dim fname As String * 1024
    Dim cnt2 As Integer
    Dim cnt3 As Integer
    
    FileNum = FreeFile
    
    If PathExists(FileName, True) Then Kill FileName
    Open FileName For Output As #FileNum
'    Close #FileNum
'    Kill FileName
'
'    Open FileName For Binary Lock Write As #FileNum
   
'*********************************************
'Save Tempo
    Print #FileNum, CByte(frmMain.Tempo.Value)
    
'*********************************************
'Save Main Mute
    Print #FileNum, CByte(frmMain.Play.Value)

'*********************************************
'Save MixerCount
    For cnt = 0 To 9
    
        Print #FileNum, CByte(Pattern1(cnt).Mixers.Count - 1)
    
'*********************************************
'Save Beat Types
        For cnt2 = 1 To Pattern1(cnt).Mixers.Count - 1
            Print #FileNum, CByte(Asc(Pattern1(cnt).Mixers(cnt2).MixerType))
        Next
    
'*********************************************
'Save Beat Mutes
        For cnt2 = 1 To Pattern1(cnt).Mixers.Count - 1
            Print #FileNum, CByte(Pattern1(cnt).Mixers(cnt2).Mute.Value)
    
'*********************************************
'Save Beat Volumes
            Print #FileNum, CByte(Pattern1(cnt).Mixers(cnt2).volume.Value)
    
'*********************************************
'Save Beat File Names
            If Pattern1(cnt).Mixers(cnt2).MixerType = "W" Then
                If InStr(Pattern1(cnt).Mixers(cnt2).Resource.Caption, "[") > 0 Then
                    fname = Mid(Pattern1(cnt).Mixers(cnt2).Resource.Caption, InStr(Pattern1(cnt).Mixers(cnt2).Resource.Caption, "[") + 1)
                    fname = Left(fname, InStrRev(fname, "]") - 1)
                Else
                    fname = ""
                End If
                Print #FileNum, fname
            ElseIf Pattern1(cnt).Mixers(cnt2).MixerType = "M" Then
                fname = Pattern1(cnt).Mixers(cnt2).channel & "," & Pattern1(cnt).Mixers(cnt2).BaseNote
                Print #FileNum, fname
            ElseIf Pattern1(cnt).Mixers(cnt2).MixerType = "P" Then
                If Pattern1(cnt).Mixers(cnt2).Beginning.Count > 0 Then
                    For cnt3 = 1 To Pattern1(cnt).Mixers(cnt2).Beginning.Count
                        fname = Trim(Pattern1(cnt).Mixers(cnt2).Beginning(cnt3))
                        If fname <> "" Then
                            Print #FileNum, fname
                        End If
                    Next
                End If
                Print #FileNum, ""
            End If
    
'*********************************************
'Save Patterns
            For cnt3 = 0 To 31
                If Pattern1(cnt).Mixers(cnt2).Beat(cnt3).Caption = " " Or Pattern1(cnt).Mixers(cnt2).Beat(cnt3).Caption = "" Then
                    Print #FileNum, CByte(Asc(" "))
                Else
                    Print #FileNum, CByte(Asc(Pattern1(cnt).Mixers(cnt2).Beat(cnt3).Caption))
                End If
            Next
            
        Next
    
    Next
'*********************************************
    
    Close #FileNum

    Me.Caption = "SoSouiX.net Sequencer - [" & GetFileTitle(FileName) & "]"

End Sub

Public Sub LoadPattern(ByVal FileName As String)

    Dim X, cnt As Integer, cnt2 As Integer
    Dim FileNum As Integer
    Dim fname As String * 1024
    Dim inVal As String
    Dim MixerCnt As Byte
    Dim cnt3 As Integer
    
    FileNum = FreeFile
    Open FileName For Input As #FileNum
   
'*********************************************
'Load Tempo
    Line Input #FileNum, inVal
    frmMain.Tempo.Value = CByte(inVal)
    
'*********************************************
'Load Main Mute
    Line Input #FileNum, inVal
    frmMain.Play.Value = CByte(inVal)
    
'*********************************************
'Load MixerCount
    For cnt = 0 To 9
    
        Line Input #FileNum, inVal
        MixerCnt = CByte(inVal)
        
'*********************************************
'Load Beat Types
        For cnt2 = 1 To MixerCnt
            Line Input #FileNum, inVal
            Pattern1(cnt).AddMixer Chr(inVal)
        
        Next

'*********************************************
'Load Beat Mutes
        For cnt2 = 1 To Pattern1(cnt).Mixers.Count - 1
            
            Line Input #FileNum, inVal
            Pattern1(cnt).Mixers(cnt2).Mute.Value = CByte(inVal)
        
'*********************************************
'Load Beat Volumes
            Line Input #FileNum, inVal
            Pattern1(cnt).Mixers(cnt2).volume.Value = CByte(inVal)
    
'*********************************************
'Load Beat File Names

            ClearCollection Pattern1(cnt).Mixers(cnt2).Beginning
            ClearCollection Pattern1(cnt).Mixers(cnt2).Recordings
            If Pattern1(cnt).Mixers(cnt2).MixerType = "W" Then
                Line Input #FileNum, fname
                If Not Trim(fname) = "" Then
                    Pattern1(cnt).Mixers(cnt2).LoadWaveFileName AppPath & fname
                End If
            ElseIf Pattern1(cnt).Mixers(cnt2).MixerType = "M" Then
                Line Input #FileNum, fname
                Pattern1(cnt).Mixers(cnt2).channel = CLng(Left(fname, InStr(fname, ",") - 1))
                Pattern1(cnt).Mixers(cnt2).BaseNote = CLng(Mid(fname, InStr(fname, ",") + 1))
            ElseIf Pattern1(cnt).Mixers(cnt2).MixerType = "P" Then
                Do
                    Line Input #FileNum, fname
                    If (Trim(fname) <> "") Then
                        Pattern1(cnt).Mixers(cnt2).Recordings.Add fname
                        Pattern1(cnt).Mixers(cnt2).Beginning.Add fname
                    End If
                Loop Until Trim(fname) = ""
            End If
   
'*********************************************
'Load Patterns
            For cnt3 = 0 To 31
                Line Input #FileNum, inVal
                If Chr(inVal) = " " Then
                    Pattern1(cnt).Mixers(cnt2).Beat(cnt3).Caption = " "
                Else
                    Pattern1(cnt).Mixers(cnt2).Beat(cnt3).Caption = Chr(inVal)
                End If
                Select Case Pattern1(cnt).Mixers(cnt2).Beat(cnt3).Caption
                    Case "S"
                        Pattern1(cnt).Mixers(cnt2).Beat(cnt3).Value = 1
                    Case "L"
                        Pattern1(cnt).Mixers(cnt2).Beat(cnt3).Value = 1
                    Case "W"
                        Pattern1(cnt).Mixers(cnt2).Beat(cnt3).Value = 1
                    Case " ", ""
                        Pattern1(cnt).Mixers(cnt2).Beat(cnt3).Value = 0
                End Select
            Next
        Next
    Next
    
'*********************************************
    
    Close #FileNum

    Me.Caption = "SoSouiX.net Sequencer - [" & GetFileTitle(FileName) & "]"

End Sub


