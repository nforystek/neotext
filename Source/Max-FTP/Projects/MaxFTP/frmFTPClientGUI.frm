VERSION 5.00
Object = "{C98B112F-745F-4542-B5B3-DDFADF1F6E2F}#1180.0#0"; "NTControls22.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmFTPClientGUI 
   AutoRedraw      =   -1  'True
   Caption         =   "Max-FTP"
   ClientHeight    =   6465
   ClientLeft      =   4050
   ClientTop       =   2625
   ClientWidth     =   11490
   HelpContextID   =   2
   Icon            =   "frmFTPClientGUI.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "frmSiteFTP"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6465
   ScaleWidth      =   11490
   Tag             =   "ftpundefined"
   Visible         =   0   'False
   Begin VB.PictureBox pFileIcon 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   4905
      ScaleHeight     =   240
      ScaleMode       =   0  'User
      ScaleWidth      =   240
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1035
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox userGUI 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4890
      Index           =   1
      Left            =   7095
      ScaleHeight     =   4890
      ScaleWidth      =   4140
      TabIndex        =   18
      Top             =   60
      Width           =   4140
      Begin VB.ListBox pHistory 
         Height          =   255
         Index           =   1
         Left            =   3045
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   3345
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox dContainer 
         BorderStyle     =   0  'None
         Height          =   735
         Index           =   5
         Left            =   105
         ScaleHeight     =   735
         ScaleWidth      =   3855
         TabIndex        =   21
         Top             =   1875
         Width           =   3855
         Begin VB.PictureBox pAddressBar 
            BorderStyle     =   0  'None
            Height          =   330
            Index           =   1
            Left            =   270
            ScaleHeight     =   330
            ScaleWidth      =   2850
            TabIndex        =   22
            Top             =   270
            Width           =   2850
            Begin MSComctlLib.Toolbar UserGo 
               Height          =   330
               Index           =   1
               Left            =   1395
               TabIndex        =   23
               Top             =   45
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   582
               ButtonWidth     =   767
               ButtonHeight    =   582
               AllowCustomize  =   0   'False
               Wrappable       =   0   'False
               HelpContextID   =   2
               Style           =   1
               _Version        =   393216
            End
            Begin NTControls22.AutoType setLocation 
               Height          =   315
               Index           =   1
               Left            =   540
               TabIndex        =   6
               Top             =   15
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   556
               ReadOnly        =   0   'False
            End
         End
         Begin VB.Line Line8 
            BorderColor     =   &H80000014&
            Index           =   5
            Tag             =   "highlight"
            X1              =   540
            X2              =   1680
            Y1              =   705
            Y2              =   705
         End
         Begin VB.Line Line7 
            BorderColor     =   &H80000010&
            Index           =   5
            Tag             =   "shadow"
            X1              =   480
            X2              =   1635
            Y1              =   570
            Y2              =   570
         End
         Begin VB.Line Line6 
            BorderColor     =   &H80000014&
            Index           =   5
            Tag             =   "highlight"
            X1              =   465
            X2              =   1605
            Y1              =   180
            Y2              =   180
         End
         Begin VB.Line Line5 
            BorderColor     =   &H80000010&
            Index           =   5
            Tag             =   "shadow"
            X1              =   420
            X2              =   1710
            Y1              =   45
            Y2              =   45
         End
         Begin VB.Line Line4 
            BorderColor     =   &H80000010&
            Index           =   5
            Tag             =   "shadow"
            X1              =   1815
            X2              =   1815
            Y1              =   135
            Y2              =   630
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000014&
            Index           =   5
            Tag             =   "highlight"
            X1              =   2025
            X2              =   2025
            Y1              =   90
            Y2              =   660
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000014&
            Index           =   5
            Tag             =   "highlight"
            X1              =   255
            X2              =   255
            Y1              =   60
            Y2              =   675
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   5
            Tag             =   "shadow"
            X1              =   120
            X2              =   120
            Y1              =   60
            Y2              =   720
         End
      End
      Begin VB.PictureBox dContainer 
         BorderStyle     =   0  'None
         Height          =   735
         Index           =   4
         Left            =   420
         ScaleHeight     =   735
         ScaleWidth      =   2685
         TabIndex        =   20
         Top             =   825
         Width           =   2685
         Begin VB.DriveListBox pViewDrives 
            Height          =   315
            HelpContextID   =   2
            Index           =   1
            Left            =   585
            TabIndex        =   5
            Top             =   225
            Width           =   4320
         End
         Begin VB.Line Line8 
            BorderColor     =   &H80000014&
            Index           =   4
            Tag             =   "highlight"
            X1              =   540
            X2              =   1680
            Y1              =   705
            Y2              =   705
         End
         Begin VB.Line Line7 
            BorderColor     =   &H80000010&
            Index           =   4
            Tag             =   "shadow"
            X1              =   480
            X2              =   1635
            Y1              =   570
            Y2              =   570
         End
         Begin VB.Line Line6 
            BorderColor     =   &H80000014&
            Index           =   4
            Tag             =   "highlight"
            X1              =   465
            X2              =   1605
            Y1              =   180
            Y2              =   180
         End
         Begin VB.Line Line5 
            BorderColor     =   &H80000010&
            Index           =   4
            Tag             =   "shadow"
            X1              =   420
            X2              =   1710
            Y1              =   45
            Y2              =   45
         End
         Begin VB.Line Line4 
            BorderColor     =   &H80000010&
            Index           =   4
            Tag             =   "shadow"
            X1              =   1815
            X2              =   1815
            Y1              =   135
            Y2              =   630
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000014&
            Index           =   4
            Tag             =   "highlight"
            X1              =   2025
            X2              =   2025
            Y1              =   90
            Y2              =   660
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000014&
            Index           =   4
            Tag             =   "highlight"
            X1              =   255
            X2              =   255
            Y1              =   60
            Y2              =   675
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   4
            Tag             =   "shadow"
            X1              =   120
            X2              =   120
            Y1              =   60
            Y2              =   720
         End
      End
      Begin VB.PictureBox dContainer 
         BorderStyle     =   0  'None
         Height          =   735
         Index           =   3
         Left            =   150
         ScaleHeight     =   735
         ScaleWidth      =   3120
         TabIndex        =   19
         Top             =   45
         Width           =   3120
         Begin MSComctlLib.Toolbar UserControls 
            Height          =   450
            Index           =   1
            Left            =   1275
            TabIndex        =   4
            Top             =   240
            Width           =   4380
            _ExtentX        =   7726
            _ExtentY        =   794
            ButtonWidth     =   820
            ButtonHeight    =   794
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            HelpContextID   =   2
            Style           =   1
            _Version        =   393216
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   3
            Tag             =   "shadow"
            X1              =   120
            X2              =   120
            Y1              =   60
            Y2              =   720
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000014&
            Index           =   3
            Tag             =   "highlight"
            X1              =   255
            X2              =   255
            Y1              =   60
            Y2              =   675
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000014&
            Index           =   3
            Tag             =   "highlight"
            X1              =   2025
            X2              =   2025
            Y1              =   90
            Y2              =   660
         End
         Begin VB.Line Line4 
            BorderColor     =   &H80000010&
            Index           =   3
            Tag             =   "shadow"
            X1              =   1815
            X2              =   1815
            Y1              =   135
            Y2              =   630
         End
         Begin VB.Line Line5 
            BorderColor     =   &H80000010&
            Index           =   3
            Tag             =   "shadow"
            X1              =   420
            X2              =   1710
            Y1              =   45
            Y2              =   45
         End
         Begin VB.Line Line6 
            BorderColor     =   &H80000014&
            Index           =   3
            Tag             =   "highlight"
            X1              =   465
            X2              =   1605
            Y1              =   180
            Y2              =   180
         End
         Begin VB.Line Line7 
            BorderColor     =   &H80000010&
            Index           =   3
            Tag             =   "shadow"
            X1              =   480
            X2              =   1635
            Y1              =   570
            Y2              =   570
         End
         Begin VB.Line Line8 
            BorderColor     =   &H80000014&
            Index           =   3
            Tag             =   "highlight"
            X1              =   540
            X2              =   1680
            Y1              =   705
            Y2              =   705
         End
      End
      Begin MSComctlLib.ListView pView 
         Height          =   1245
         HelpContextID   =   2
         Index           =   1
         Left            =   -15
         TabIndex        =   7
         Top             =   2880
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   2196
         View            =   3
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         OLEDropMode     =   1
         AllowReorder    =   -1  'True
         PictureAlignment=   4
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Size"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Access"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ProgressBar pProgress 
         Height          =   300
         Index           =   1
         Left            =   105
         TabIndex        =   24
         Top             =   4500
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   1
         Max             =   1
         Scrolling       =   1
      End
      Begin MSComctlLib.StatusBar pStatus 
         Height          =   315
         Index           =   1
         Left            =   15
         TabIndex        =   25
         Top             =   4275
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   1
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   1
               Object.Width           =   4868
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000005&
         Height          =   1125
         Index           =   1
         Left            =   645
         ScaleHeight     =   1065
         ScaleWidth      =   1335
         TabIndex        =   33
         Top             =   2955
         Visible         =   0   'False
         Width           =   1395
         Begin RichTextLib.RichTextBox pDummyView 
            Height          =   810
            Index           =   1
            Left            =   330
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   405
            Visible         =   0   'False
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   1429
            _Version        =   393217
            BorderStyle     =   0
            ReadOnly        =   -1  'True
            ScrollBars      =   3
            Appearance      =   0
            TextRTF         =   $"frmFTPClientGUI.frx":08CA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Image Image1 
            Height          =   600
            Index           =   1
            Left            =   165
            Top             =   240
            Visible         =   0   'False
            Width           =   435
         End
      End
   End
   Begin VB.PictureBox pFileIcon 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   5055
      ScaleHeight     =   240
      ScaleMode       =   0  'User
      ScaleWidth      =   240
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   570
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox userGUI 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4785
      Index           =   0
      Left            =   150
      ScaleHeight     =   4785
      ScaleWidth      =   3930
      TabIndex        =   9
      Top             =   150
      Width           =   3930
      Begin VB.ListBox pHistory 
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   3540
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox dContainer 
         BorderStyle     =   0  'None
         Height          =   735
         Index           =   0
         Left            =   240
         ScaleHeight     =   735
         ScaleWidth      =   3120
         TabIndex        =   14
         Top             =   60
         Width           =   3120
         Begin MSComctlLib.Toolbar UserControls 
            Height          =   330
            Index           =   0
            Left            =   1020
            TabIndex        =   0
            Top             =   195
            Width           =   4380
            _ExtentX        =   7726
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            HelpContextID   =   2
            Style           =   1
            _Version        =   393216
         End
         Begin VB.Line Line8 
            BorderColor     =   &H80000014&
            Index           =   0
            Tag             =   "highlight"
            X1              =   540
            X2              =   1680
            Y1              =   705
            Y2              =   705
         End
         Begin VB.Line Line7 
            BorderColor     =   &H80000010&
            Index           =   0
            Tag             =   "shadow"
            X1              =   480
            X2              =   1635
            Y1              =   570
            Y2              =   570
         End
         Begin VB.Line Line6 
            BorderColor     =   &H80000014&
            Index           =   0
            Tag             =   "highlight"
            X1              =   465
            X2              =   1605
            Y1              =   180
            Y2              =   180
         End
         Begin VB.Line Line5 
            BorderColor     =   &H80000010&
            Index           =   0
            Tag             =   "shadow"
            X1              =   420
            X2              =   1710
            Y1              =   45
            Y2              =   45
         End
         Begin VB.Line Line4 
            BorderColor     =   &H80000010&
            Index           =   0
            Tag             =   "shadow"
            X1              =   1815
            X2              =   1815
            Y1              =   135
            Y2              =   630
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000014&
            Index           =   0
            Tag             =   "highlight"
            X1              =   2025
            X2              =   2025
            Y1              =   90
            Y2              =   660
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000014&
            Index           =   0
            Tag             =   "highlight"
            X1              =   255
            X2              =   255
            Y1              =   60
            Y2              =   675
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   0
            Tag             =   "shadow"
            X1              =   120
            X2              =   120
            Y1              =   60
            Y2              =   720
         End
      End
      Begin VB.PictureBox dContainer 
         BorderStyle     =   0  'None
         Height          =   735
         Index           =   1
         Left            =   420
         ScaleHeight     =   735
         ScaleWidth      =   2685
         TabIndex        =   13
         Top             =   825
         Width           =   2685
         Begin VB.DriveListBox pViewDrives 
            Height          =   315
            HelpContextID   =   2
            Index           =   0
            Left            =   690
            TabIndex        =   1
            Top             =   225
            Width           =   4320
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   1
            Tag             =   "shadow"
            X1              =   120
            X2              =   120
            Y1              =   60
            Y2              =   720
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000014&
            Index           =   1
            Tag             =   "highlight"
            X1              =   255
            X2              =   255
            Y1              =   60
            Y2              =   675
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000014&
            Index           =   1
            Tag             =   "highlight"
            X1              =   2025
            X2              =   2025
            Y1              =   90
            Y2              =   660
         End
         Begin VB.Line Line4 
            BorderColor     =   &H80000010&
            Index           =   1
            Tag             =   "shadow"
            X1              =   1815
            X2              =   1815
            Y1              =   135
            Y2              =   630
         End
         Begin VB.Line Line5 
            BorderColor     =   &H80000010&
            Index           =   1
            Tag             =   "shadow"
            X1              =   420
            X2              =   1710
            Y1              =   45
            Y2              =   45
         End
         Begin VB.Line Line6 
            BorderColor     =   &H80000014&
            Index           =   1
            Tag             =   "highlight"
            X1              =   465
            X2              =   1605
            Y1              =   180
            Y2              =   180
         End
         Begin VB.Line Line7 
            BorderColor     =   &H80000010&
            Index           =   1
            Tag             =   "shadow"
            X1              =   480
            X2              =   1635
            Y1              =   570
            Y2              =   570
         End
         Begin VB.Line Line8 
            BorderColor     =   &H80000014&
            Index           =   1
            Tag             =   "highlight"
            X1              =   540
            X2              =   1680
            Y1              =   705
            Y2              =   705
         End
      End
      Begin VB.PictureBox dContainer 
         BorderStyle     =   0  'None
         Height          =   735
         Index           =   2
         Left            =   30
         ScaleHeight     =   735
         ScaleWidth      =   3105
         TabIndex        =   10
         Top             =   1710
         Width           =   3105
         Begin VB.PictureBox pAddressBar 
            BorderStyle     =   0  'None
            Height          =   330
            Index           =   0
            Left            =   525
            ScaleHeight     =   330
            ScaleWidth      =   2130
            TabIndex        =   11
            Top             =   255
            Width           =   2130
            Begin NTControls22.AutoType setLocation 
               Height          =   315
               Index           =   0
               Left            =   240
               TabIndex        =   2
               Top             =   0
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   556
               ReadOnly        =   0   'False
            End
            Begin MSComctlLib.Toolbar UserGo 
               Height          =   324
               Index           =   0
               Left            =   912
               TabIndex        =   12
               Top             =   48
               Width           =   708
               _ExtentX        =   1244
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               AllowCustomize  =   0   'False
               Wrappable       =   0   'False
               HelpContextID   =   2
               Style           =   1
               _Version        =   393216
            End
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   2
            Tag             =   "shadow"
            X1              =   120
            X2              =   120
            Y1              =   60
            Y2              =   720
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000014&
            Index           =   2
            Tag             =   "highlight"
            X1              =   255
            X2              =   255
            Y1              =   60
            Y2              =   675
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000014&
            Index           =   2
            Tag             =   "highlight"
            X1              =   2025
            X2              =   2025
            Y1              =   90
            Y2              =   660
         End
         Begin VB.Line Line4 
            BorderColor     =   &H80000010&
            Index           =   2
            Tag             =   "shadow"
            X1              =   1815
            X2              =   1815
            Y1              =   135
            Y2              =   630
         End
         Begin VB.Line Line5 
            BorderColor     =   &H80000010&
            Index           =   2
            Tag             =   "shadow"
            X1              =   420
            X2              =   1710
            Y1              =   45
            Y2              =   45
         End
         Begin VB.Line Line6 
            BorderColor     =   &H80000014&
            Index           =   2
            Tag             =   "highlight"
            X1              =   465
            X2              =   1605
            Y1              =   180
            Y2              =   180
         End
         Begin VB.Line Line7 
            BorderColor     =   &H80000010&
            Index           =   2
            Tag             =   "shadow"
            X1              =   480
            X2              =   1635
            Y1              =   570
            Y2              =   570
         End
         Begin VB.Line Line8 
            BorderColor     =   &H80000014&
            Index           =   2
            Tag             =   "highlight"
            X1              =   540
            X2              =   1680
            Y1              =   705
            Y2              =   705
         End
      End
      Begin MSComctlLib.ListView pView 
         Height          =   1245
         HelpContextID   =   2
         Index           =   0
         Left            =   15
         TabIndex        =   3
         Top             =   2895
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   2196
         View            =   3
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         OLEDropMode     =   1
         PictureAlignment=   4
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Access"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Access"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ProgressBar pProgress 
         Height          =   300
         Index           =   0
         Left            =   225
         TabIndex        =   15
         Top             =   4485
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   1
         Max             =   1
         Scrolling       =   1
      End
      Begin MSComctlLib.StatusBar pStatus 
         Height          =   315
         Index           =   0
         Left            =   105
         TabIndex        =   16
         Top             =   4305
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   1
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   1
               Object.Width           =   4868
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000005&
         Height          =   1065
         Index           =   0
         Left            =   870
         ScaleHeight     =   1005
         ScaleWidth      =   1095
         TabIndex        =   32
         Top             =   3000
         Visible         =   0   'False
         Width           =   1155
         Begin RichTextLib.RichTextBox pDummyView 
            Height          =   540
            HelpContextID   =   2
            Index           =   0
            Left            =   315
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   953
            _Version        =   393217
            BorderStyle     =   0
            ReadOnly        =   -1  'True
            ScrollBars      =   3
            Appearance      =   0
            TextRTF         =   $"frmFTPClientGUI.frx":094D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Image Image1 
            Height          =   585
            Index           =   0
            Left            =   135
            Top             =   60
            Visible         =   0   'False
            Width           =   435
         End
      End
   End
   Begin VB.PictureBox hSizer 
      BorderStyle     =   0  'None
      Height          =   90
      Left            =   4755
      MousePointer    =   7  'Size N S
      ScaleHeight     =   90
      ScaleWidth      =   1740
      TabIndex        =   30
      Top             =   4125
      Width           =   1740
   End
   Begin VB.PictureBox vSizer 
      BorderStyle     =   0  'None
      Height          =   3285
      Left            =   5520
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3285
      ScaleWidth      =   75
      TabIndex        =   31
      Top             =   705
      Width           =   75
   End
   Begin VB.PictureBox userInfo 
      BorderStyle     =   0  'None
      Height          =   1080
      Left            =   960
      ScaleHeight     =   1080
      ScaleWidth      =   8925
      TabIndex        =   29
      Top             =   5205
      Width           =   8925
      Begin MSComctlLib.ListView ListView1 
         Height          =   750
         HelpContextID   =   10
         Left            =   75
         TabIndex        =   8
         Top             =   105
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   1323
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Action"
            Object.Width           =   2364
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Progress"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Bytes"
            Object.Width           =   4586
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Rate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "File Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Source Folder"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Destination Folder"
            Object.Width           =   5292
         EndProperty
      End
   End
   Begin VB.Menu mnuSite 
      Caption         =   "&Max"
      WindowList      =   -1  'True
      Begin VB.Menu mnuNewWindow 
         Caption         =   "New &Client Window"
      End
      Begin VB.Menu mnuNewScheduleWin 
         Caption         =   "Schedule &Manager"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuShowActiveCache 
         Caption         =   "&Active App Cache"
      End
      Begin VB.Menu mnuDash65 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuNewScriptWin 
         Caption         =   "&Scripting IDE"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRunScript 
         Caption         =   "&Run Script"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuDoubleWin 
         Caption         =   "Double &Window"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuMultiThread 
         Caption         =   "&Multi-Thread"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuDash20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInformation 
         Caption         =   "&Information"
         Begin VB.Menu mnuInfoSystem 
            Caption         =   "&System"
         End
         Begin VB.Menu mnudash232435 
            Caption         =   "-"
         End
         Begin VB.Menu mnuInfoBanner 
            Caption         =   "&Banner"
         End
         Begin VB.Menu mnuInfoMOTD 
            Caption         =   "&MOTD"
         End
         Begin VB.Menu mnudash23782 
            Caption         =   "-"
         End
         Begin VB.Menu mnuInfoStat 
            Caption         =   "S&tat"
         End
         Begin VB.Menu mnuInfoHelp 
            Caption         =   "&Help"
         End
      End
      Begin VB.Menu mnuShowLog 
         Caption         =   "Show &Logging"
      End
      Begin VB.Menu mnudash2378 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowToolBar 
         Caption         =   "&Tool Bar"
      End
      Begin VB.Menu mnuShowDriveList 
         Caption         =   "&Drive List"
      End
      Begin VB.Menu mnuShowAddressBar 
         Caption         =   "&Address Bar"
      End
      Begin VB.Menu mnuShowStatusBar 
         Caption         =   "&Status Bar"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuCache 
         Caption         =   "Cac&he"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDash22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "C&ut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuDash6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewFolder 
         Caption         =   "&New Folder"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "De&lete"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "&Rename"
      End
      Begin VB.Menu mnuDash7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelect 
         Caption         =   "S&elect"
         Begin VB.Menu mnuSelAll 
            Caption         =   "&All"
         End
         Begin VB.Menu mnuSelFiles 
            Caption         =   "&Files"
         End
         Begin VB.Menu mnuSelFolders 
            Caption         =   "F&olders"
         End
         Begin VB.Menu mnuWildCard 
            Caption         =   "&Pattern..."
         End
      End
      Begin VB.Menu mnudash4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Re&fresh"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "&Stop"
      End
      Begin VB.Menu mnuDash231 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConnect 
         Caption         =   "&Connect"
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "&Disconnect"
      End
   End
   Begin VB.Menu mnuFavorites 
      Caption         =   "Fav&orites"
      Begin VB.Menu mnuSSetup 
         Caption         =   "&Manage Favorites"
      End
      Begin VB.Menu mnuDash2345 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFavorite 
         Caption         =   "(No Favorites Found)"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Setup"
      Begin VB.Menu mnuPrefrences 
         Caption         =   "Setup &Options"
      End
      Begin VB.Menu mnuFileAssoc 
         Caption         =   "&File Associations"
      End
      Begin VB.Menu mnuNetDrives 
         Caption         =   "&Network Drives"
      End
      Begin VB.Menu mnuDash2198 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowToolTips 
         Caption         =   "Show Toolti&ps"
      End
      Begin VB.Menu mnuTipOfDay 
         Caption         =   "&Tip of the Day"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp2 
         Caption         =   "&Documentation..."
         HelpContextID   =   1
      End
      Begin VB.Menu mnuWebSite 
         Caption         =   "&Neotext.org..."
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu mnuThread 
      Caption         =   "&Threads"
      Visible         =   0   'False
      Begin VB.Menu mnuClearFinished 
         Caption         =   "Clear &Finished"
      End
      Begin VB.Menu mnuClearStopped 
         Caption         =   "Clear &Stopped"
      End
      Begin VB.Menu mnuDash287 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancelTransfer 
         Caption         =   "&Cancel Transfer"
      End
      Begin VB.Menu mnuRetryTransfer 
         Caption         =   "&Retry Transfer"
      End
   End
End
Attribute VB_Name = "frmFTPClientGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Option Compare Binary

Private Const Border = 4

Private ftpRates(1, 2) As Double
Private ftpTimer As Long

Public PopUp As NTPopup21.Window

Public FocusIndex As Integer
Public HasFocus As Boolean

Public MenuAction As MenuActions

Public Enum LogMsgTypes
    log_Outgoing = 1
    log_Incomming = 2
    log_Error = 3
End Enum

Public Enum FTPStateTypes
    st_StartUp = 0
    st_Processing = 1
    st_ProcessFailed = 2
    st_ProcessSuccess = 3
    st_ViewingLog = 4
End Enum
Private LastState(1) As Integer

Public MyDescription As String
Public ClientIndex As Long

Public WithEvents myClient0 As NTAdvFTP61.Client
Attribute myClient0.VB_VarHelpID = -1
Public WithEvents myClient1 As NTAdvFTP61.Client
Attribute myClient1.VB_VarHelpID = -1

Private URLType(1) As NTAdvFTP61.URLTypes
Private tmpListFile(1) As String

Private FTPCommand(1) As String
Private FTPState(1) As Integer
Private FTPUnloading(1) As Boolean

Private vBarIsSizing As Boolean
Private hBarIsSizing As Boolean
Public ReCenterSizers As Boolean

Private HistoryPointer(1) As Integer
Private AllowAddToHistory As Boolean

Public NoResize As Boolean
Public MultiThread As Boolean
Private pClientGUILoaded(1) As Boolean
Private CancelStatusGUI(1) As Boolean

Private PreviousFileRate(1) As String
Private PreviousFileSize(1) As Double

Private InRecursiveAction(1) As Boolean
Private RecursiveListItems0() As String
Private RecursiveListItems1() As String
Private RecursiveFileName(1) As String
Private RecursiveFileSize(1) As Double

Private pCopyIsSource(1) As Boolean
Private pCopyToClientIndex(1) As Integer
Private pCopyToClientForm(1) As frmFTPClientGUI

Private myTransfers() As clsTransfer
Private myTransferCount As Long

Private pPrevWndProc As Long

Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)

Public Property Get PrevWndProc() As Long
    PrevWndProc = pPrevWndProc
End Property
Public Property Let PrevWndProc(ByVal newVal As Long)
    pPrevWndProc = newVal
End Property

Public Property Let InRecursion(ByVal Index As Long, ByVal newVal As Boolean)
    InRecursiveAction(Index) = newVal
End Property

Public Property Get InRecursion(ByVal Index As Long) As Boolean
    InRecursion = InRecursiveAction(Index)
End Property

Public Sub CreateTransferThread(ByRef myForm As Form, ByVal OwnerID As String, ByVal Action As String, ByRef copyClient As NTAdvFTP61.Client, ByRef copyToClient As NTAdvFTP61.Client, ByVal FileName As String, ByVal FileSize As Double, Optional ByVal ResumeByte As Double = 0, Optional ByVal DestFileSize As Double = 0)
   
    Dim newTransfer As clsTransfer
    Set newTransfer = New clsTransfer
    
    newTransfer.LoadTransfer "ClientID: "
    
    myTransferCount = myTransferCount + 1
    ReDim Preserve myTransfers(1 To myTransferCount) As clsTransfer
    Set myTransfers(myTransferCount) = newTransfer
    
    Set newTransfer.MyParent = myForm
    newTransfer.sOwnerID = OwnerID
    newTransfer.SetSource copyClient.URL, copyClient.Port, copyClient.Username, copyClient.Password, copyClient.ImplicitSSL
    newTransfer.SetDestination copyToClient.URL, copyToClient.Port, copyToClient.Username, copyToClient.Password, copyToClient.ImplicitSSL
    newTransfer.SetTransfer Action, copyClient.Folder, copyToClient.Folder, FileName, FileSize, ResumeByte, DestFileSize
    newTransfer.ThreadOpen
    
    newTransfer.StartTransfer
    
    Set newTransfer = Nothing
    
    SetCaption
    
End Sub

Public Sub DestroyTransferThread(ByVal TransferIndex As Long)

    If (myTransferCount > 1) Then

        Dim cnt As Long
        Dim tmp() As clsTransfer
        ReDim tmp(1 To myTransferCount) As clsTransfer
             
         If ((Not ((TransferIndex = 2) And (myTransferCount = 2))) Or (myTransferCount >= 2)) Then
             
             Set tmp(myTransferCount) = myTransfers(TransferIndex)
             Set myTransfers(TransferIndex) = myTransfers(myTransferCount)
             Set myTransfers(myTransferCount) = tmp(myTransferCount)
 
         End If
   
             For cnt = 1 To myTransferCount - 1
                 Set tmp(cnt) = myTransfers(IIf((cnt >= TransferIndex) And (myTransferCount >= 2), cnt + 1, cnt))
             Next
         
         If ((Not (myTransferCount = 2)) And (Not (TransferIndex = myTransferCount))) Or ((myTransferCount >= 2) And (Not (TransferIndex = myTransferCount))) Then
             Set tmp(myTransferCount - 1) = myTransfers(TransferIndex)
         End If
         
          For cnt = 1 To myTransferCount - 1
              Set myTransfers(cnt) = tmp(cnt)
          Next

        myTransferCount = myTransferCount - 1
        ReDim Preserve myTransfers(1 To myTransferCount) As clsTransfer
        
    ElseIf (myTransferCount = 1) Then
        myTransferCount = myTransferCount - 1
        Set myTransfers(1) = Nothing
    End If
    
    SetCaption

End Sub

Public Property Get GetTransferThread(ByVal TransferIndex As Long) As clsTransfer
    If TransferIndex <= myTransferCount Then
        If Not (myTransfers(TransferIndex) Is Nothing) Then
            On Error GoTo 0
            Set GetTransferThread = myTransfers(TransferIndex)
        End If
    End If
End Property
Public Property Set GetTransferThread(ByVal TransferIndex As Long, ByVal newVal As clsTransfer)
    If TransferIndex <= myTransferCount Then
        If Not (myTransfers(TransferIndex) Is Nothing) Then
            On Error GoTo 0
            myTransfers(TransferIndex) = newVal
        End If
    End If
End Property

Public Property Get copyToClientForm(ByVal Index As Integer) As frmFTPClientGUI
    Set copyToClientForm = pCopyToClientForm(Index)
End Property
Public Property Set copyToClientForm(ByVal Index As Integer, ByRef newVal As frmFTPClientGUI)
    Set pCopyToClientForm(Index) = newVal
End Property

Public Property Get copyToClientIndex(ByVal Index As Integer) As Integer
    copyToClientIndex = pCopyToClientIndex(Index)
End Property
Public Property Let copyToClientIndex(ByVal Index As Integer, ByVal newVal As Integer)
    pCopyToClientIndex(Index) = newVal
End Property

Public Property Get copyIsSource(ByVal Index As Integer) As Boolean
    copyIsSource = pCopyIsSource(Index)
End Property
Public Property Let copyIsSource(ByVal Index As Integer, ByVal newVal As Boolean)
    pCopyIsSource(Index) = newVal
End Property

Public Property Get ClientGUILoaded(ByVal Index As Integer) As Boolean
    ClientGUILoaded = pClientGUILoaded(Index)
End Property
Public Property Let ClientGUILoaded(ByVal Index As Integer, ByVal newVal As Boolean)
    pClientGUILoaded(Index) = newVal
End Property

Public Function GetState(ByVal Index As Integer) As Integer
    GetState = FTPState(Index)
End Function

Public Sub SetMyFocus()
    Dim frms
    For Each frms In Forms
        If TypeName(frms) = "frmFTPClientGUI" Then
            frms.HasFocus = (frms.hwnd = Me.hwnd)
        End If
    Next
End Sub

Public Sub SetCancelStatusGUI(ByVal Index As Integer, ByVal newVal As Boolean)
    CancelStatusGUI(Index) = newVal
End Sub
Public Sub SetInRecursiveAction(ByVal Index As Integer, ByVal newVal As Boolean)
    InRecursiveAction(Index) = newVal
End Sub

Public Function GetCancelStatusGUI(ByVal Index As Integer) As Boolean
    GetCancelStatusGUI = CancelStatusGUI(Index)
End Function
Public Function GetInRecursiveAction(ByVal Index As Integer) As Boolean
    GetInRecursiveAction = InRecursiveAction(Index)
End Function

Public Function GetFTPCommand(ByVal Index As Integer) As Boolean
    GetFTPCommand = FTPCommand(Index)
End Function
Public Function SetFTPCommand(ByVal Index As Integer, ByVal newVal As String)
    FTPCommand(Index) = newVal
End Function

Private Sub dContainer_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    Source.DragIcon = frmMain.imgDragDrop.ListImages("abort").Picture
End Sub

Private Sub dContainer_Resize(Index As Integer)

    On Error Resume Next
    dContainerResize Me, Index
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Private Sub Form_Activate()

    SetMyFocus

End Sub

Private Sub Form_DragOver(Source As Control, X As Single, Y As Single, State As Integer)

    Source.DragIcon = frmMain.imgDragDrop.ListImages("abort").Picture

End Sub

Public Sub LoadClient(Optional ByVal ClientDescription As String = "")
            
    mnuNewScheduleWin.Visible = IsSchedulerInstalled
    mnuNewScriptWin.Visible = IsScriptIDEInstalled
    mnuRunScript.Visible = IsScriptIDEInstalled
    mnuDash65.Visible = mnuNewScriptWin.Visible

    mnuShowToolTips.Checked = dbSettings.GetProfileSetting("ViewToolTips")

    mnuMultiThread.Checked = dbSettings.GetClientSetting("MultiThread")
    
    frmMain.RefreshFavorite Me
    
    If ClientDescription = "" Then
        ClientDescription = "ClientID: "
    End If
    
    ClientIndex = ThreadManager.AddClients()
    MyDescription = ClientDescription & ClientIndex
    
    Set myClient0 = ThreadManager.GetClients(ClientIndex).FTPClient1
    myClient0.timeout = dbSettings.GetProfileSetting("TimeOut")
    myClient0.PauseOnStandBy = dbSettings.GetPublicSetting("ServiceStandBy")
    myClient0.LogBytes = dbSettings.GetClientSetting("LogFileSize")
    myClient0.TransferRates(0) = dbSettings.GetProfileSetting("ftpLocalSize")
    myClient0.TransferRates(1) = dbSettings.GetProfileSetting("ftpBufferSize")
    myClient0.TransferRates(2) = dbSettings.GetProfileSetting("ftpPacketSize")
    myClient0.LargeFileMode = dbSettings.GetProfileSetting("LargeFileMode")
    myClient0.ImplicitSSL = (dbSettings.GetProfileSetting("SSL") = 1)
    
    Set myClient1 = ThreadManager.GetClients(ClientIndex).FTPClient2
    myClient1.timeout = dbSettings.GetProfileSetting("TimeOut")
    myClient1.PauseOnStandBy = dbSettings.GetPublicSetting("ServiceStandBy")
    myClient1.LogBytes = dbSettings.GetClientSetting("LogFileSize")
    myClient1.TransferRates(0) = dbSettings.GetProfileSetting("ftpLocalSize")
    myClient1.TransferRates(1) = dbSettings.GetProfileSetting("ftpBufferSize")
    myClient1.TransferRates(2) = dbSettings.GetProfileSetting("ftpPacketSize")
    myClient1.LargeFileMode = dbSettings.GetProfileSetting("LargeFileMode")
    myClient1.ImplicitSSL = (dbSettings.GetProfileSetting("SSL") = 1)
    
    If dbSettings.GetProfileSetting("ConnectionMode") = 1 Then
        myClient0.ConnectionMode = "PORT"
        myClient1.ConnectionMode = "PORT"
    Else
        myClient0.ConnectionMode = "PASV"
        myClient1.ConnectionMode = "PASV"
    End If

    pDummyView(0).RightMargin = 12000
    pDummyView(1).RightMargin = 12000
    
    tmpListFile(0) = GetTemporaryFolder & "\maxFTP_" & Me.hwnd & "_0.lst"
    tmpListFile(1) = GetTemporaryFolder & "\maxFTP_" & Me.hwnd & "_1.lst"

    ClientGUILoaded(0) = False
    ClientGUILoaded(1) = False

    CancelStatusGUI(0) = False
    CancelStatusGUI(1) = False
    InRecursiveAction(0) = False
    InRecursiveAction(1) = False

    NoResize = True

    HistoryPointer(0) = -1
    HistoryPointer(1) = -1
    AllowAddToHistory = True

    FocusIndex = 0
    
End Sub
 
Public Sub SetCaption(Optional ByVal FileName As String = "", Optional ByVal percent As Integer = -1)
    Dim newCap As String
    If MultiThread Then
        newCap = AppName & " [" & Trim(myTransferCount) & IIf(myTransferCount = 1, " Transfer", " Transfer(s)") & "]"
    ElseIf (FileName = "") Then
        newCap = AppName
    Else
        If (percent >= 0) And (percent <= 100) Then
            newCap = AppName & " [" & Trim(CStr(percent)) & "% - " & FileName & "]"
        Else
            newCap = AppName & " [" & FileName & "]"
        End If
    End If
    If Not (Me.Caption = newCap) Then Me.Caption = newCap
End Sub

Public Sub ShowClient()
    
    Me.Caption = AppName

    If dbSettings.GetClientSetting("wState") = -1 Then
        dbSettings.SetClientSetting "wState", 0
        dbSettings.SetClientSetting "wLeft", ((Screen.Width / 2) - (Me.Width / 2))
        dbSettings.SetClientSetting "wTop", ((Screen.Height / 2) - (Me.Height / 2))
        dbSettings.SetClientSetting "wWidth", Me.Width
        dbSettings.SetClientSetting "wHeight", Me.Height
    Else
       
        Dim frm As Form
        For Each frm In Forms
            If (TypeName(frm) = TypeName(Me)) Then
                If (Not (frm.hwnd = Me.hwnd)) And Me.Visible Then
                    dbSettings.SetClientSetting "wLeft", frm.Left + (Screen.TwipsPerPixelX * 32)
                    dbSettings.SetClientSetting "wTop", frm.Top + (Screen.TwipsPerPixelY * 32)
                    dbSettings.SetClientSetting "wWidth", frm.Width
                    dbSettings.SetClientSetting "wHeight", frm.Height
                End If
            End If
        Next
    End If
    
    If ((dbSettings.GetClientSetting("wLeft") + dbSettings.GetClientSetting("wWidth")) > Screen.Width) Or _
        ((dbSettings.GetClientSetting("wTop") + dbSettings.GetClientSetting("wHeight")) > Screen.Height) Then
        dbSettings.SetClientSetting "wLeft", (32 * Screen.TwipsPerPixelX)
        dbSettings.SetClientSetting "wTop", (32 * Screen.TwipsPerPixelY)
    End If

    Me.Move dbSettings.GetClientSetting("wLeft"), dbSettings.GetClientSetting("wTop"), dbSettings.GetClientSetting("wWidth"), dbSettings.GetClientSetting("wHeight")
    
    If Me.Left + Me.Width > Screen.Width Then
        If Me.Left > Me.Width Then
            Me.Left = Me.Left - (Screen.Width - (Me.Left + Me.Width))
        Else
            Me.Width = Me.Width - (Screen.Width - (Me.Left + Me.Width))
        End If
        
        If Me.Top > Me.Height Then
            Me.Top = Me.Top - (Screen.Height - (Me.Top + Me.Height))
        Else
            Me.Height = Me.Height - (Screen.Height - (Me.Top + Me.Height))
        End If
    End If
    
    Me.WindowState = IIf(dbSettings.GetProfileSetting("SystemTray"), 0, dbSettings.GetClientSetting("wState"))
    
    vSizer.Move dbSettings.GetClientSetting("wVSizer")
    hSizer.Move dbSettings.GetClientSetting("wHSizer")
    pView(0).ColumnHeaders(1).Width = dbSettings.GetClientSetting("wColumn0_1")
    pView(0).ColumnHeaders(3).Width = dbSettings.GetClientSetting("wColumn0_3")
    pView(0).ColumnHeaders(4).Width = dbSettings.GetClientSetting("wColumn0_4")
    pView(1).ColumnHeaders(1).Width = dbSettings.GetClientSetting("wColumn1_1")
    pView(1).ColumnHeaders(2).Width = dbSettings.GetClientSetting("wColumn1_2")
    pView(1).ColumnHeaders(3).Width = dbSettings.GetClientSetting("wColumn1_3")
    pView(1).ColumnHeaders(4).Width = dbSettings.GetClientSetting("wColumn1_4")

    LoadGUI 0

    ViewDoubleWindow Me, dbSettings.GetClientSetting("ViewDoubleWindow")
    ViewTransferInfo Me, dbSettings.GetClientSetting("MultiThread")
    ViewDriveList Me, dbSettings.GetClientSetting("ViewDriveList")
    ViewStatusBar Me, dbSettings.GetClientSetting("ViewStatusBar")
    ViewToolBar Me, dbSettings.GetClientSetting("ViewToolBar")
    ViewAddressBar Me, dbSettings.GetClientSetting("ViewAddressBar")

    SetAutoTypeList Me, setLocation(0)
    SetAutoTypeList Me, setLocation(1)

    SetTooltip
    
    SetClientGUIState 0, st_StartUp
    SetClientGUIState 1, st_StartUp

    NoResize = False
    MultiThread = dbSettings.GetClientSetting("MultiThread")

    frmMain.RefreshFavorites

    SendMessageLngPtr pDummyView(0).hwnd, EM_SETTEXTMODE, TM_RICHTEXT, ByVal 0&
    SendMessageLngPtr pDummyView(1).hwnd, EM_SETTEXTMODE, TM_RICHTEXT, ByVal 0&
   
    Me.Show
    
End Sub

Private Sub Form_LostFocus()
    FocusIndex = -1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If IsActiveForm(Me) Or UnloadMode = 1 Then
        Cancel = 0
    Else
        If PromptAbortClose("Are you sure you want to close this window?" & vbCrLf & "(any connections or transfers will be canceled)" & vbCrLf) Then
    
            Cancel = 0
            
        Else
            Cancel = 1
            frmMain.timGlobal.enabled = True
        End If
    End If
    
    If Cancel = 0 Then
            
        If Not (PopUp Is Nothing) Then
            If PopUp.Visible Then PopUp.Hide
            PopUp.ParentHWnd = 0
            Set PopUp = Nothing
        End If
    
        If (FTPState(0) = st_Processing) Then FTPStop 0 ' = st_Processing Then FTPStop 0
        If (FTPState(1) = st_Processing) Then FTPStop 1 ' = st_Processing
    
        Dim frms
        For Each frms In Forms
            If (TypeName(frms) = "frmPassword") Or (TypeName(frms) = "frmOverwrite") Then
                If frms.ParentHWnd = Me.hwnd Then
                    frms.ParentHWnd = 0
                    Unload frms
                End If
            End If
        Next
        
        FTPDisconnect 0
        FTPDisconnect 1
        
        CancelAllTransfers
            
        ThreadManager.RemoveClients ClientIndex
        
        Set myClient0 = Nothing
        Set myClient1 = Nothing
        
    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If IsVisibleClient(Me) And Me.Visible And Me.WindowState = 0 Then
    
        Me.ReCenterSizers = Not (vBarIsSizing Or hBarIsSizing)
        
    End If
    
    If Me.WindowState <> 1 Then
        FormResize Me
    End If

    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Private Sub FormUnload()
    
    dbSettings.SetClientSetting "wState", Me.WindowState
    
    If IsVisibleClient(Me) And Me.Visible And Me.WindowState = 0 Then
    
        If Me.Top > 0 And Me.Left > 0 Then
            dbSettings.SetClientSetting "wTop", Me.Top
            dbSettings.SetClientSetting "wLeft", Me.Left
            dbSettings.SetClientSetting "wHeight", Me.Height
            dbSettings.SetClientSetting "wWidth", Me.Width
        End If
    
    End If
    
    If IsVisibleClient(Me) Then
    
        dbSettings.SetClientSetting "wColumn0_1", pView(0).ColumnHeaders(1).Width
        dbSettings.SetClientSetting "wColumn0_2", pView(0).ColumnHeaders(2).Width
        dbSettings.SetClientSetting "wColumn0_3", pView(0).ColumnHeaders(3).Width
        dbSettings.SetClientSetting "wColumn0_4", pView(0).ColumnHeaders(4).Width
    
        If mnuMultiThread.Checked Then
            dbSettings.SetClientSetting "wHSizer", hSizer.Top
        End If
    
        UnloadGUI 0
    
        If mnuDoubleWin.Checked Then
            UnloadGUI 1
            dbSettings.SetClientSetting "wVSizer", vSizer.Left
            dbSettings.SetClientSetting "wColumn1_1", pView(1).ColumnHeaders(1).Width
            dbSettings.SetClientSetting "wColumn1_2", pView(1).ColumnHeaders(2).Width
            dbSettings.SetClientSetting "wColumn1_3", pView(1).ColumnHeaders(3).Width
            dbSettings.SetClientSetting "wColumn1_4", pView(1).ColumnHeaders(4).Width
        End If
        
    End If
    
    MyDescription = ""
    
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormUnload
    
End Sub

Private Sub hSizer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.ReCenterSizers = True
    If Button = 1 And Not hBarIsSizing Then
        hBarIsSizing = True
        hSizer.BackColor = &H808080
        vSizer.BackColor = &H808080
    Else
        If Button = 1 And hBarIsSizing Then
            
            hSizer.Top = hSizer.Top + Y
            hBarIsSizing = True
            Form_Resize
        Else
            If hBarIsSizing Then
                Form_Resize
                hBarIsSizing = False
                hSizer.BackColor = &H8000000F
                vSizer.BackColor = &H8000000F
            End If
        End If
    End If

End Sub

Private Sub SetToolBarGUIState(ByVal Index As Integer, ByVal cState As Integer)

    Dim AllowBack As Boolean
    Dim AllowForward As Boolean
    
    If pHistory(Index).ListCount > 0 And HistoryPointer(Index) > -1 Then
        If HistoryPointer(Index) <= 0 Then
            AllowBack = False
        Else
            AllowBack = True
        End If
    Else
        AllowBack = False
    End If
    If pHistory(Index).ListCount > 0 And HistoryPointer(Index) > -1 Then
        If HistoryPointer(Index) >= pHistory(Index).ListCount - 1 Then
            AllowForward = False
        Else
            AllowForward = True
        End If
    Else
        AllowForward = False
    End If

    If ClientGUILoaded(Index) Then
        With UserControls(Index)
            .Buttons(1).enabled = (cState = st_ProcessSuccess)
            .Buttons(2).enabled = ((cState = st_ProcessFailed) Or (cState = st_ProcessSuccess)) And AllowBack
            .Buttons(3).enabled = ((cState = st_ProcessFailed) Or (cState = st_ProcessSuccess)) And AllowForward
            .Buttons(5).enabled = (cState = st_Processing) Or (cState = st_ProcessSuccess) Or (cState = st_ProcessFailed)
            .Buttons(6).enabled = (cState = st_ProcessFailed) Or (cState = st_ProcessSuccess)
            .Buttons(8).enabled = (cState = st_ProcessSuccess)
            .Buttons(9).enabled = (cState = st_ProcessSuccess)
            .Buttons(11).enabled = (cState = st_ProcessSuccess)
            .Buttons(12).enabled = (cState = st_ProcessSuccess)
            .Buttons(13).enabled = (cState = st_ProcessSuccess)
        
        End With

    End If
    
    SetConnectGUIState Index, cState
    
End Sub

Public Sub SetConnectGUIState(ByVal Index As Integer, ByVal cState As Integer)
    If ClientGUILoaded(Index) Then
        With UserGo(Index)
            .Buttons(1).enabled = (cState = st_ProcessFailed) Or (cState = st_ProcessSuccess) Or (cState = st_StartUp)
            .Buttons(1).Visible = SetMyClient(Index).ConnectedState() = False
            .Buttons(2).Visible = Not .Buttons(1).Visible
            .Buttons(3).enabled = (cState = st_ProcessFailed) Or (cState = st_ProcessSuccess) Or (cState = st_StartUp)
        End With
    End If
End Sub

Private Sub SetContentGUIState(ByVal Index As Integer, ByVal cState As Integer)

    If Not (pDummyView(Index).Visible = (URLType(Index) = URLTypes.ftp)) Then pDummyView(Index).Visible = (URLType(Index) = URLTypes.ftp)
    If Not (Image1(Index).Visible = Not (URLType(Index) = URLTypes.ftp)) Then Image1(Index).Visible = Not (URLType(Index) = URLTypes.ftp)
    
    Select Case cState
        Case st_StartUp
            pView(Index).Visible = False
            If (Not Picture1(Index).Visible) Then Picture1(Index).Visible = True
        Case st_Processing
            Select Case URLType(Index)
                Case URLTypes.File, URLTypes.Remote, URLTypes.ftp
                    If pDummyView(Index).Text = "" Then
                        pView(Index).Visible = False
                        If (Not Picture1(Index).Visible) Then Picture1(Index).Visible = True
                    Else
                        pView(Index).Visible = False
                        If (Not Picture1(Index).Visible) Then Picture1(Index).Visible = True
                    End If
            End Select
        Case st_ProcessFailed
            Select Case URLType(Index)
                Case URLTypes.File, URLTypes.Remote, URLTypes.ftp
                    If pDummyView(Index).Text = "" Then
                        pView(Index).Visible = False
                        If (Not Picture1(Index).Visible) Then Picture1(Index).Visible = True
                    Else
                        pView(Index).Visible = False
                        If (Not Picture1(Index).Visible) Then Picture1(Index).Visible = True
                    End If
            End Select
        Case st_ViewingLog
            Select Case URLType(Index)
                Case URLTypes.File, URLTypes.Remote, URLTypes.ftp
                    pView(Index).Visible = False
                    If Not (Picture1(Index).Visible) Then Picture1(Index).Visible = True
            End Select
        Case st_ProcessSuccess
            Select Case URLType(Index)
                Case URLTypes.File, URLTypes.Remote, URLTypes.ftp
                    If (Picture1(Index).Visible) Then Picture1(Index).Visible = False
                    pView(Index).Visible = True
            End Select
    End Select

End Sub

Private Sub SetInputGUIState(ByVal Index As Integer, ByVal cState As Integer)
    
    Select Case cState
        Case st_StartUp
            pViewDrives(Index).enabled = True
            setLocation(Index).enabled = True
        
        Case st_Processing
            pViewDrives(Index).enabled = False
            setLocation(Index).enabled = False
        
        Case st_ProcessFailed
            pViewDrives(Index).enabled = True
            setLocation(Index).enabled = True
        
        Case st_ProcessSuccess
            pViewDrives(Index).enabled = True
            setLocation(Index).enabled = True

    End Select

End Sub

Public Sub SetClientGUIState(ByVal Index As Integer, ByVal cState As Integer)
    
    If cState = st_Processing And (FTPState(Index) = st_Processing) Then Exit Sub
    
    FTPState(Index) = cState
    
    If Not CancelStatusGUI(Index) Then
        SetToolBarGUIState Index, cState
        SetContentGUIState Index, cState
        SetInputGUIState Index, cState
    End If
    
    If cState = st_ProcessSuccess Then
        SetStatusTotals Index
    Else
        SetStatus Index, ""
    End If
    
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        RefreshThreadMenu
    End If
End Sub

'Private Sub mnuCache_Click()
'    FTPOpen FocusIndex, True
'End Sub

Private Sub mnuFavorite_Click(Index As Integer)
    If PathExists(mnuFavorite(Index).Tag) Then
    
        Dim ftpSite1 As New frmFavoriteSite
        ftpSite1.LoadSite mnuFavorite(Index).Tag
    
        Me.FTPOpenSite ftpSite1
        
        Unload ftpSite1
        
    End If
End Sub

Private Sub mnuInfoBanner_Click()
    Dim myClient As NTAdvFTP61.Client
    Set myClient = SetMyClient(FocusIndex)

    Set PopUp = New NTPopup21.Window
    PopUp.ParentHWnd = Me.hwnd
    
    PopUp.AlwaysOnTop = True
    PopUp.Message = Replace(Replace(Replace(Replace(myClient.BannerMessage, vbCrLf, vbLf), vbCr, "\n"), vbLf, "\n"), "\n", vbCrLf)
    PopUp.Title = "Server Information"
    PopUp.Icon = vbInformation
    PopUp.LinkText = "Welcome Banner Message"
    PopUp.Visible = True
    
    Set myClient = Nothing
End Sub

Private Sub mnuInfoHelp_Click()
    Dim myClient As NTAdvFTP61.Client
    Set myClient = SetMyClient(FocusIndex)

    Set PopUp = New NTPopup21.Window
    PopUp.ParentHWnd = Me.hwnd
    
    PopUp.AlwaysOnTop = True
    PopUp.Message = Replace(Replace(Replace(Replace(myClient.SystemHelp, vbCrLf, vbLf), vbCr, "\n"), vbLf, "\n"), "\n", vbCrLf)
    PopUp.Title = "Server Information"
    PopUp.Icon = vbInformation
    PopUp.LinkText = "Remote Command Help"
    PopUp.Visible = True
    
    Set myClient = Nothing
End Sub

Private Sub mnuInfoMOTD_Click()
    Dim myClient As NTAdvFTP61.Client
    Set myClient = SetMyClient(FocusIndex)

    Set PopUp = New NTPopup21.Window
    PopUp.ParentHWnd = Me.hwnd
    
    PopUp.AlwaysOnTop = True
    PopUp.Message = Replace(Replace(Replace(Replace(myClient.MessageOfTheDay, vbCrLf, vbLf), vbCr, "\n"), vbLf, "\n"), "\n", vbCrLf)
    PopUp.Title = "Server Information"
    PopUp.Icon = vbInformation
    PopUp.LinkText = "Login Message of the Day"
    PopUp.Visible = True
    
    Set myClient = Nothing
End Sub

Private Sub mnuInfoStat_Click()
    Dim myClient As NTAdvFTP61.Client
    Set myClient = SetMyClient(FocusIndex)

    Set PopUp = New NTPopup21.Window
    PopUp.ParentHWnd = Me.hwnd
    
    PopUp.AlwaysOnTop = True
    PopUp.Message = Replace(Replace(Replace(Replace(myClient.StatInformation, vbCrLf, vbLf), vbCr, "\n"), vbLf, "\n"), "\n", vbCrLf)
    PopUp.Title = "Server Information"
    PopUp.Icon = vbInformation
    PopUp.LinkText = "Stat Information"
    PopUp.Visible = True
    
    Set myClient = Nothing
End Sub

Private Sub mnuInfoSystem_Click()
    Dim myClient As NTAdvFTP61.Client
    Set myClient = SetMyClient(FocusIndex)

    Set PopUp = New NTPopup21.Window
    PopUp.ParentHWnd = Me.hwnd
    
    PopUp.AlwaysOnTop = True
    PopUp.Message = Replace(Replace(Replace(Replace(myClient.RemoteSystem, vbCrLf, vbLf), vbCr, "\n"), vbLf, "\n"), "\n", vbCrLf)
    PopUp.Title = "Server Information"
    PopUp.Icon = vbInformation
    PopUp.LinkText = "Remote System Type"
    PopUp.Visible = True
    
    Set myClient = Nothing
End Sub

Private Sub mnuNewScriptWin_Click()
    RunProcess AppPath & MaxIDEFileName, "", vbNormalFocus, False
End Sub

Private Sub mnuConnect_Click()
    FTPConnect FocusIndex
End Sub

Private Sub mnuCopy_Click()
    CopyFiles Me, FocusIndex, "Copy"
End Sub

Private Sub mnuCut_Click()
    CopyFiles Me, FocusIndex, "Move"
End Sub

Private Sub mnuDelete_Click()
    FTPDelete FocusIndex
End Sub

Private Sub mnuDisconnect_Click()
    FTPDisconnect FocusIndex
End Sub

Private Sub mnuDoubleWin_Click()
    ViewDoubleWindow Me, Not mnuDoubleWin.Checked
End Sub

Private Sub mnuExit_Click()
    If PromptAbortClose("Are you sure you want to close this window?") Then
        Unload Me
    End If
End Sub

Private Sub mnuFile_Click()
    EnableFileMenu FocusIndex
End Sub

Private Sub mnuFileAssoc_Click()
    frmFileAssoc.Show
End Sub

Private Sub mnuHelp2_Click()
    GotoHelp
End Sub

Private Sub mnuHelpAbout_Click()
    
    frmAbout.Show

End Sub

Private Sub mnuNetDrives_Click()
    frmNetDrives.Show
End Sub

Private Sub mnuNewFolder_Click()
    FTPNewFolder FocusIndex, GetNewFolderName(FocusIndex)
End Sub

Private Sub mnuNewScheduleWin_Click()
    frmSchManager.Show
End Sub

Private Sub mnuNewWindow_Click()
    Dim newWin As New frmFTPClientGUI
    newWin.LoadClient
    newWin.ShowClient

End Sub

Private Sub mnuOpen_Click()
    FTPOpen FocusIndex
End Sub

Private Sub mnuPaste_Click()
    PasteFromClipboard Me, FocusIndex
End Sub

Private Sub mnuPrefrences_Click()
    frmSetup.Show
End Sub

Private Sub mnuRefresh_Click()
    FTPRefresh FocusIndex
End Sub

Private Sub mnuRename_Click()
    pView(FocusIndex).StartLabelEdit
End Sub

Private Sub mnuRunScript_Click()
    
    frmMain.RunScript
   
End Sub

Private Sub mnuSelAll_Click()
    SelectWildCard pView(FocusIndex), "all", "*"
    SetStatusTotals FocusIndex
End Sub

Private Sub mnuSelFiles_Click()
    SelectWildCard pView(FocusIndex), "files", "*"
    SetStatusTotals FocusIndex

End Sub

Private Sub mnuSelFolders_Click()
    SelectWildCard pView(FocusIndex), "folders", "*"
    SetStatusTotals FocusIndex
End Sub

Private Sub mnuShowActiveCache_Click()
    frmActiveCache.ShowForm
End Sub

Private Sub mnuShowAddressBar_Click()
    ViewAddressBar Me, Not mnuShowAddressBar.Checked
End Sub

Private Sub mnuShowDriveList_Click()
    ViewDriveList Me, Not mnuShowDriveList.Checked
End Sub

Private Sub mnuShowLog_Click()
    If mnuShowLog.Checked Then
        mnuShowLog.Checked = False
        
        SetClientGUIState 0, LastState(0)
        SetClientGUIState 1, LastState(1)
    Else
        mnuShowLog.Checked = True
        
        LastState(0) = FTPState(0)
        LastState(1) = FTPState(1)
        
        SetClientGUIState 0, st_ViewingLog
        SetClientGUIState 1, st_ViewingLog
    End If
    
End Sub

Private Sub mnuShowStatusBar_Click()
    ViewStatusBar Me, Not mnuShowStatusBar.Checked
End Sub

Private Sub mnuShowToolBar_Click()
    ViewToolBar Me, Not mnuShowToolBar.Checked
End Sub

Private Sub mnuShowToolTips_Click()
    mnuShowToolTips.Checked = Not mnuShowToolTips.Checked
    dbSettings.SetProfileSetting "ViewToolTips", CBool(mnuShowToolTips.Checked)
    frmMain.ResetToolTips CBool(mnuShowToolTips.Checked)
End Sub

Private Sub mnuSSetup_Click()
    frmFavorites.Show

End Sub

Private Sub mnuStop_Click()
    If PromptAbortClose("Do you really want to cancel the current action?") Then
        FTPStop FocusIndex
    End If
End Sub

Private Sub mnuMultiThread_Click()

    ViewTransferInfo Me, Not mnuMultiThread.Checked
    MultiThread = mnuMultiThread.Checked

End Sub

Private Sub mnuTipOfDay_Click()
    frmTipOfDay.Show
End Sub

Private Sub mnuView_Click()
    EnableFileMenu FocusIndex
End Sub

Private Sub mnuWebSite_Click()
    RunFile AppPath & "Neotext.org.url"
End Sub

Private Sub mnuWildCard_Click()
    Dim newCard As frmWildCards
    Set newCard = New frmWildCards
    newCard.Show 1
    If newCard.IsOk Then
        
        Select Case newCard.WildOption
            Case 0
                SelectWildCard pView(FocusIndex), "files", newCard.WildCard
            Case 1
                SelectWildCard pView(FocusIndex), "folders", newCard.WildCard
            Case 2
                SelectWildCard pView(FocusIndex), "all", newCard.WildCard
        End Select
    End If
    SetStatusTotals FocusIndex
    Unload newCard
    Set newCard = Nothing
End Sub

Private Sub myClient0_DataComplete(ByVal ProgressType As NTAdvFTP61.ProgressTypes)

    If ProgressType = FileListing Then
        If InRecursiveAction(0) Then
            RecursiveClientListComplete 0
        Else
            ClientListComplete 0
        End If
    Else
        SetCaption
        RecursiveClientFileComplete 0
    End If
    
End Sub

Private Sub myClient0_DataProgress(ByVal ProgressType As NTAdvFTP61.ProgressTypes, ByVal ReceivedBytes As Double)
    
    If ProgressType = FileListing Then
        ClientListProgress 0, ReceivedBytes
    Else
        ClientFileProgress 0, ProgressType, ReceivedBytes
    End If
    
End Sub

Public Sub myClient0_Error(ByVal Number As Long, ByVal Source As String, ByVal Description As String)
    SetCaption
    ClientError 0, Number, Source, Description
End Sub


Private Sub myClient0_LogMessage(ByVal MessageType As NTAdvFTP61.MessageTypes, ByVal AddedText As String)

    ClientAddToLog 0, MessageType, AddedText
    
End Sub

Private Sub myClient1_DataComplete(ByVal ProgressType As NTAdvFTP61.ProgressTypes)

    If ProgressType = FileListing Then
        If InRecursiveAction(1) Then
            RecursiveClientListComplete 1
        Else
            ClientListComplete 1
        End If
    Else
        SetCaption
        RecursiveClientFileComplete 1
    
    End If
End Sub

Private Sub myClient1_DataProgress(ByVal ProgressType As NTAdvFTP61.ProgressTypes, ByVal ReceivedBytes As Double)

    If ProgressType = FileListing Then
        ClientListProgress 1, ReceivedBytes
    Else
        ClientFileProgress 1, ProgressType, ReceivedBytes
    End If
    
End Sub

Public Sub myClient1_Error(ByVal Number As Long, ByVal Source As String, ByVal Description As String)
    SetCaption
    
    ClientError 1, Number, Source, Description

End Sub


Private Sub myClient1_ItemListing(ByVal ItemName As String, ByVal ItemSize As String, ByVal ItemDate As String, ByVal ItemAccess As String)
'
End Sub

Private Sub myClient1_LogMessage(ByVal MessageType As NTAdvFTP61.MessageTypes, ByVal AddedText As String)
    
    ClientAddToLog 1, MessageType, AddedText
    
End Sub

Private Sub pAddressBar_Resize(Index As Integer)
    On Error Resume Next
    If NoResize Then Exit Sub
    setLocation(Index).Left = 15
    setLocation(Index).Width = pAddressBar(Index).Width - setLocation(Index).Left - UserGo(Index).Width
    UserGo(Index).Left = pAddressBar(Index).Width - UserGo(Index).Width
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Private Sub pDummyView_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    Source.DragIcon = frmMain.imgDragDrop.ListImages("abort").Picture

End Sub

Private Sub pDummyView_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    FocusIndex = Index

End Sub

Private Sub pProgress_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    Source.DragIcon = frmMain.imgDragDrop.ListImages("abort").Picture

End Sub

Private Sub pStatus_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    Source.DragIcon = frmMain.imgDragDrop.ListImages("abort").Picture

End Sub

Private Sub pView_AfterLabelEdit(Index As Integer, Cancel As Integer, NewString As String)

    Dim OldString As String
    
    OldString = pView(Index).SelectedItem.Text
    
    If Left(NewString, 1) = "/" Then NewString = Mid(NewString, 2)
    
    If Not IsFileNameValid(NewString) Then
        If Me.Visible Then MsgBox "File and Folder names can not contain the characters \ / : * ? "" < > |", vbInformation, AppName
        Cancel = True
    Else
        If Not pView(Index).SelectedItem Is Nothing Then
            FTPRename Index, pView(Index).SelectedItem.Text, NewString
        End If
    
        If Left(OldString, 1) = "/" Then
            If Not Left(NewString, 1) = "/" Then NewString = "/" & NewString
        End If
    End If

End Sub

Private Sub pView_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ColumnSortClick pView(Index), ColumnHeader
    
    dbSettings.SetClientSetting "wColumnKey" & Trim(Index), pView(Index).SortKey
    dbSettings.SetClientSetting "wColumnSort" & Trim(Index), pView(Index).SortOrder

End Sub

Private Sub pView_DblClick(Index As Integer)
    FTPOpen Index
End Sub

Private Sub pView_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)

    If Not (Index = Source.Index And Source.Parent.hwnd = Me.hwnd) Then
        If dbSettings.GetClientSetting("DragOption") = 0 Then


            frmDragOption.PromptDragOption

            If frmDragOption.IsOk = True Then
                Select Case frmDragOption.DragOption
                    Case 1
                        CopyFiles Source.Parent, Source.Index, "Copy"
                        PasteFromClipboard Me, Index
                    Case 2
                        CopyFiles Source.Parent, Source.Index, "Move"
                        PasteFromClipboard Me, Index
                End Select
            End If

            Unload frmDragOption
        Else
            Select Case dbSettings.GetClientSetting("DragOption")
                Case 1
                    CopyFiles Source.Parent, Source.Index, "Copy"
                    PasteFromClipboard Me, Index
                Case 2
                    CopyFiles Source.Parent, Source.Index, "Move"
                    PasteFromClipboard Me, Index
            End Select
        End If
    End If

End Sub

Private Sub pView_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    If Index = Source.Index And Source.Parent.hwnd = Me.hwnd Then
        If dbSettings.GetClientSetting("DragOption") > 1 Then
            pView(Index).DragIcon = frmMain.imgDragDrop.ListImages("cut").Picture
        Else
            pView(Index).DragIcon = frmMain.imgDragDrop.ListImages("copy").Picture
        End If
    Else
        Source.DragIcon = frmMain.imgDragDrop.ListImages("paste").Picture
    End If

End Sub

Private Sub pView_GotFocus(Index As Integer)
    FocusIndex = Index

End Sub

Private Sub pView_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        FTPDelete Index
    End If
End Sub

Private Sub pView_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 2 Then
        
        FocusIndex = Index
        
        EnableFileMenu Index

        modMenuUp.PopUp Me.hwnd, 2
        
    End If
End Sub
Public Property Get FileMenuObject() As Menu
    Set FileMenuObject = mnuFile
End Property

Private Sub SetStatusTotals(ByVal Index As Integer)
    Dim tot As Long
    tot = pView(Index).ListItems.count
    If tot > 0 Then
        Dim cnt As Long
        Dim tmp As Long
        Dim fld1 As Long
        Dim fld2 As Long
                
        For cnt = 1 To tot
            If pView(Index).ListItems(cnt).Selected Then
                tmp = tmp + 1
                If Left(pView(Index).ListItems(cnt).Text, 1) = "/" Then
                    fld2 = fld2 + 1
                End If
            End If
            If Left(pView(Index).ListItems(cnt).Text, 1) = "/" Then
                fld1 = fld1 + 1
            End If
        Next
        
        Dim fle1 As Long
        Dim fle2 As Long
        fle1 = tot - fld1
        fle2 = tmp - fld2
            
        If (tmp > 1) Then
            If ((fld2 = 0) Or (fle2 = 0)) And ((fld1 = 0) Or (fle1 = 0)) Then
                SetStatus Index, Trim(tmp) & " Selected"
            Else
                SetStatus Index, Trim(tmp) & " Selected (" & Trim(fld2) & " folders, " & Trim(fle2) & " files)"
            End If
        Else

            If ((fld1 = 0) Or (fle1 = 0)) Then
                SetStatus Index, Trim(tot) & " Total"
            Else
                SetStatus Index, Trim(tot) & " Total (" & Trim(fld1) & " folders, " & Trim(fle1) & " files)"
            End If
        End If
    Else
        SetStatus Index, "0 Total"
    End If
End Sub

Private Sub pView_OLEDragDrop(Index As Integer, Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    PasteFromDragDrop Me, Index, Data
    
End Sub

Private Sub pViewDrives_Change(Index As Integer)
    Dim myClient As NTAdvFTP61.Client
    Set myClient = SetMyClient(Index)
    
    If Not LCase(Left(pViewDrives(Index).Drive, 2)) = LCase(Left(myClient.Folder, 2)) Then
        
        setLocation(Index).Text = LCase(Left(pViewDrives(Index).Drive, 2)) & "\"
        
        FTPConnect Index
    
    End If
    
    Set myClient = Nothing
End Sub

Private Sub pViewDrives_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
        Source.DragIcon = frmMain.imgDragDrop.ListImages("abort").Picture
End Sub

Private Sub pViewDrives_GotFocus(Index As Integer)
    FocusIndex = Index
End Sub

Private Sub setLocation_Change(Index As Integer)
   SetConnectGUIState Index, LastState(Index)
End Sub

Private Sub setLocation_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    Source.DragIcon = frmMain.imgDragDrop.ListImages("abort").Picture
End Sub

Private Sub setLocation_GotFocus(Index As Integer)
    FocusIndex = Index
End Sub

Private Sub setLocation_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        FTPConnect Index
    End If
End Sub

Private Sub UserControls_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Key
        Case "uplevel"
            FTPChangeFolder Index, ".."
            
        Case "back"
            GoHistory Index, "back"
        Case "forward"
            GoHistory Index, "forward"
    
        Case "stop"
            If (FTPState(Index) = st_ProcessSuccess Or FTPState(Index) = st_ProcessFailed) And (Not (FTPCommand(Index) = "CopyMove")) Then
                FTPDisconnect Index
            Else
                If PromptAbortClose("Do you really want to cancel the current action?") Then
                    FTPStop Index
                End If
            End If
                    
        Case "refresh"
            FTPConnect Index

        Case "newfolder"
            FTPNewFolder Index, GetNewFolderName(Index)
        
        Case "delete"
            FTPDelete Index
        Case "copy"
            CopyFiles Me, Index, "Copy"
        Case "cut"
            CopyFiles Me, Index, "Move"
        Case "paste"
            PasteFromClipboard Me, Index
    End Select
End Sub

Private Sub UserControls_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    Source.DragIcon = frmMain.imgDragDrop.ListImages("abort").Picture
End Sub

Private Sub UserGo_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "close"
            If Not (FTPState(Index) = st_ProcessSuccess Or FTPState(Index) = st_ProcessFailed) Then
                If PromptAbortClose("Do you really want to close the current connection?") Then

                    FTPStop Index

                    FTPDisconnect Index
                    
                End If
            Else
                FTPDisconnect Index
            End If
        Case "go"
           FTPConnect Index
        Case "browse"
            Dim nBrowse As String
            nBrowse = BrowseAction(bws_Desktop, Me.hwnd)
            If nBrowse <> "" Then
                If Len(nBrowse) = 2 And Right(nBrowse, 1) = ":" Then
                    setLocation(Index).Text = nBrowse & "\"
                Else
                    setLocation(Index).Text = nBrowse
                End If
                FTPConnect Index
            End If
    End Select
End Sub

Private Sub userGUI_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    Source.DragIcon = frmMain.imgDragDrop.ListImages("abort").Picture
End Sub

Private Sub vSizer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.ReCenterSizers = True
    If Button = 1 And Not vBarIsSizing Then
        vBarIsSizing = True
        vSizer.BackColor = &H808080
        hSizer.BackColor = &H808080
    Else
        If Button = 1 And vBarIsSizing Then

            vSizer.Left = vSizer.Left + X
            vBarIsSizing = True
            Form_Resize
        Else
            If vBarIsSizing Then
                Form_Resize
                vBarIsSizing = False
                vSizer.BackColor = &H8000000F
                hSizer.BackColor = &H8000000F
            End If
        End If
    End If

End Sub

Public Sub SetStatus(ByVal Index As Integer, ByVal StatusText As String)

    If mnuShowStatusBar.Checked Then
        If Not pStatus(Index).Visible = True Then pStatus(Index).Visible = True
        If Not pProgress(Index).Visible = False Then pProgress(Index).Visible = False
        If Not pStatus(Index).Panels(1).Text = StatusText Then pStatus(Index).Panels(1).Text = StatusText
    Else
        If Not pStatus(Index).Visible = False Then pStatus(Index).Visible = False
        If Not pProgress(Index).Visible = False Then pProgress(Index).Visible = False
    End If

End Sub
Public Sub SetProgress(ByVal Index As Integer, ByVal ReceivedBytes As Double, Optional ByVal TotalBytes As Double = 0)

    If mnuShowStatusBar.Checked Then

        If Not pProgress(Index).Min = 0 Then pProgress(Index).Min = 0
        If TotalBytes <> 0 Then
            If Not pProgress(Index).Max = TotalBytes Then pProgress(Index).Max = TotalBytes
        End If

        If ReceivedBytes > pProgress(Index).Max Then
            pProgress(Index).Value = pProgress(Index).Max
        Else
            If Not pProgress(Index).Value = ReceivedBytes Then pProgress(Index).Value = ReceivedBytes
        End If

        If Not pProgress(Index).Visible = True Then pProgress(Index).Visible = True
        If Not pStatus(Index).Visible = False Then pStatus(Index).Visible = False
        If Not pStatus(Index).Panels(1).Text = "" Then pStatus(Index).Panels(1).Text = ""
    Else
        If Not pProgress(Index).Visible = False Then pProgress(Index).Visible = False
        If Not pStatus(Index).Visible = False Then pStatus(Index).Visible = False
    End If

End Sub

Public Function SetMyClient(ByVal Index As Integer) As NTAdvFTP61.Client

    If Index = 0 Then
        Set SetMyClient = myClient0
    Else
        Set SetMyClient = myClient1
    End If

End Function

Public Sub ClearFileView(ByVal Index As Integer)

    Dim lstItem

    Set pView(Index).SmallIcons = Nothing

    For Each lstItem In pView(Index).ListItems

        If lstItem.SmallIcon <> "folder" Then RemoveAssociation lstItem.SmallIcon

    Next

    pView(Index).ListItems.Clear

    pView(Index).SmallIcons = frmMain.imgFiles

End Sub
Private Sub RefreshFileView(ByVal Index As Integer, ByRef ListItems() As String)
    On Error GoTo catcH
    
    Dim lstItem

    ClearFileView Index
        
    If ListItems(0) <> "" Then
            
        Dim FileName As String
        Dim fileType As String
        Dim cnt As Integer
        Dim attr As Long
            
        If Not Trim(ListItems(0)) = "" Then
            For cnt = 0 To UBound(ListItems)
                    
                FileName = RemoveNextArg(ListItems(cnt), "|")
                If Not (URLType(Index) = URLTypes.ftp) Then
                    On Error Resume Next
                    attr = 0
                    Select Case Index
                        Case 0
                            If Not FileName = "pagefile.sys" Then
                                attr = GetAttr(MapFolder(myClient0.Folder, FileName))
                            Else
                                attr = vbSystem + vbReadOnly
                            End If
                        Case 1
                            If Not FileName = "pagefile.sys" Then
                                attr = GetAttr(MapFolder(myClient0.Folder, FileName))
                            Else
                                attr = vbSystem + vbReadOnly
                            End If
                            
                    End Select
                    If Err Then Err.Clear
                    On Error GoTo catcH
                End If

                If (((attr And vbSystem) = vbSystem) And dbSettings.GetPublicSetting("ServiceSystem")) Or (Not ((attr And vbSystem) = vbSystem)) Then
                
                    If (Left(FileName, 1) = "/") Or ((attr And vbDirectory) = vbDirectory) Then
                        fileType = "folder"
                        
                        Set lstItem = pView(Index).ListItems.Add(, MapFolder(setLocation(Index).Text, FileName), FileName, , "folder")
                    Else

                        fileType = Trim(LCase(GetFileExt(FileName)))
                    
                        If Right(setLocation(Index).Text, 1) = "\" Then
                            GetAssociation fileType, setLocation(Index).Text & FileName, pFileIcon(Index)
                        ElseIf Right(setLocation(Index).Text, 1) = "/" Then
                            GetAssociation fileType, setLocation(Index).Text & FileName, pFileIcon(Index)
                        Else
                            GetAssociation fileType, setLocation(Index).Text & "\" & FileName, pFileIcon(Index)
                        End If
                    
                        LoadAssociation pFileIcon(Index)
                    
                        Set lstItem = pView(Index).ListItems.Add(, MapFolder(setLocation(Index).Text, FileName), FileName, , pFileIcon(Index).Tag)
                        
                    End If
                    
                    lstItem.SubItems(1) = RemoveNextArg(ListItems(cnt), "|")
                    lstItem.SubItems(2) = RemoveNextArg(ListItems(cnt), "|")
                    lstItem.SubItems(3) = RemoveNextArg(ListItems(cnt), "|")
                
                End If
                
                If FTPCommand(Index) = "Stop" Then Exit For
            Next
            
        End If
        
    End If

    Exit Sub
catcH:
   ' Debug.Print Err.Description
    
    Err.Clear
End Sub

Public Sub GoHistory(ByVal Index As Integer, ByVal Action As String)
    AllowAddToHistory = False
    Select Case Action
        Case "back"
            HistoryPointer(Index) = HistoryPointer(Index) - 1
            setLocation(Index).Text = pHistory(Index).List(HistoryPointer(Index))
            
            FTPConnect Index
            
        Case "forward"
            HistoryPointer(Index) = HistoryPointer(Index) + 1
            setLocation(Index).Text = pHistory(Index).List(HistoryPointer(Index))
            
            FTPConnect Index
            
    End Select
End Sub

Public Sub AddToHistory(ByVal Index As Integer, ByVal hURL As String)
    If Not AllowAddToHistory Then Exit Sub
    
    If Not IsActiveAppFolder(hURL) Then
        If HistoryPointer(Index) = pHistory(Index).ListCount - 1 Then
            If Not (hURL = pHistory(Index).List(pHistory(Index).ListCount - 1)) Then
                pHistory(Index).AddItem hURL
                HistoryPointer(Index) = pHistory(Index).ListCount - 1
            End If
        Else
            If Not (hURL = pHistory(Index).List(HistoryPointer(Index))) Then
                pHistory(Index).AddItem hURL, HistoryPointer(Index) + 1
                HistoryPointer(Index) = HistoryPointer(Index) + 1
                Dim cnt As Integer
                For cnt = HistoryPointer(Index) + 1 To pHistory(Index).ListCount - 1
                    pHistory(Index).RemoveItem HistoryPointer(Index) + 1
                Next
            End If
        
        End If
    End If

End Sub

Public Sub RecursiveResetArray(ByVal Index As Integer)
    If Index = 0 Then
        ReDim RecursiveListItems0(0)
        RecursiveListItems0(0) = ""
    Else
        ReDim RecursiveListItems1(1)
        RecursiveListItems1(1) = ""
    End If
End Sub

Public Sub RecursiveLoopArrayToCollection(ByVal Index As Integer, ByRef listCol As NTNodes10.Collection)

    Dim cnt As Integer

    listCol.Clear

    If Index = 0 Then
        If RecursiveListItems0(0) <> "" Then
            For cnt = 0 To UBound(RecursiveListItems0)
                listCol.Add RecursiveListItems0(cnt)
            Next
        End If
    Else
        If RecursiveListItems1(0) <> "" Then
            For cnt = 0 To UBound(RecursiveListItems1)
                listCol.Add RecursiveListItems1(cnt)
            Next
        End If
    End If

End Sub

Public Sub HTTPOpenSite(ByVal URL As String)
    
    ViewDoubleWindow Me, False
    ViewDriveList Me, False
    
    setLocation(0).Text = URL
    
    FTPConnect 0
    
End Sub

Public Sub FTPOpenSite(ByVal frmSiteInfo As frmFavoriteSite)
    
    If frmSiteInfo.SiteInformation1.sHostURL.Text <> "" Then
        FTPConnect 0, frmSiteInfo.SiteInformation1
        If frmSiteInfo.SiteInformation2.sHostURL.Text <> "" Then
            If Not mnuDoubleWin.Checked Then
                mnuDoubleWin_Click
            End If
            FTPConnect 1, frmSiteInfo.SiteInformation2
        End If
        
    ElseIf frmSiteInfo.SiteInformation2.sHostURL.Text <> "" Then
    
        If Not mnuDoubleWin.Checked Then
            FTPConnect 0, frmSiteInfo.SiteInformation2
        Else
            FTPConnect 1, frmSiteInfo.SiteInformation2
        End If
        
    End If

End Sub

Public Sub FTPConnect(ByVal Index As Integer, Optional ByVal SiteInfo As Control, Optional ByVal AutoLogin As Boolean = False)
On Error GoTo catcH
        
    Dim myClient As NTAdvFTP61.Client
    Set myClient = SetMyClient(Index)

    FTPUnloading(Index) = False
    FTPCommand(Index) = "Connect"

    SetClientGUIState Index, st_Processing

    Dim nUrl As New NTAdvFTP61.URL

    If TypeName(SiteInfo) = "SiteInformation" Then
        setLocation(Index).Text = MapFolderVariables(SiteInfo.sHostURL.Text)
    End If

    If Not (nUrl.GetType(setLocation(Index).Text) = URLTypes.ftp) Then
        If (Not PathExists(setLocation(Index).Text)) And setLocation(Index).Text <> "" And ((Not (Left(LCase(Trim(setLocation(Index).Text)), 6) = "ftp://") And (Not (Left(LCase(Trim(setLocation(Index).Text)), 7) = "ftps://")))) Then
            setLocation(Index).Text = IIf(dbSettings.GetProfileSetting("SSL") = 1, "ftps://", "ftp://") & setLocation(Index).Text
        End If
    End If

    URLType(Index) = nUrl.GetType(setLocation(Index).Text)

    If (Not nUrl.GetType(setLocation(Index).Text) = myClient.URLType) Or (Not LCase(nUrl.GetServer(setLocation(Index).Text)) = LCase(myClient.Server)) Then
        myClient.Disconnect
    End If

    If (URLType(Index) = URLTypes.ftp And nUrl.GetServer(setLocation(Index).Text) = myClient.Server And myClient.Server <> "" And (Not myClient.ConnectedState() = False)) Then

        If Not (Trim(nUrl.GetFolder(setLocation(Index).Text)) = "") Then
            myClient.ChangeFolderAbsolute nUrl.GetFolder(setLocation(Index).Text)
        End If

    Else
        Dim addUrl As Boolean

        If (URLType(Index) = URLTypes.File Or URLType(Index) = URLTypes.Remote) Then

            If Not myClient.ConnectedState() = False Then
                myClient.ChangeFolderAbsolute nUrl.GetFolder(setLocation(Index).Text)
            Else
                myClient.URL = setLocation(Index).Text
                addUrl = True
            End If

        Else

            If (Not (nUrl.GetServer(setLocation(Index).Text) = nUrl.GetServer(myClient.URL))) Then

                myClient.URL = ""
                myClient.Server = ""
                myClient.Username = ""
                myClient.Password = ""

                myClient.Port = 21
                myClient.ConnectionMode = IIf((dbSettings.GetProfileSetting("ConnectionMode") = 0), "PASV", "PORT")
                myClient.DataPortRange = dbSettings.GetProfileSetting("DefaultPortRange")
                myClient.NetAdapter = dbSettings.GetProfileSetting("AdapterIndex")
                myClient.ImplicitSSL = (dbSettings.GetProfileSetting("SSL") = 1)
                addUrl = True

            ElseIf (nUrl.GetServer(setLocation(Index).Text) = nUrl.GetServer(myClient.URL)) Then
                If myClient.ConnectedState() = False Then
                    myClient.Username = ""
                    myClient.Password = ""
                    addUrl = True
                Else
                    addUrl = False
                End If
            Else
                addUrl = False
            End If

            If Not myClient.ConnectedState() = False Then
                FTPDisconnect Index
            End If

            If ((Not Left(Trim(LCase(setLocation(Index).Text)), 6) = "ftp://") And setLocation(Index).Text <> "" And (Not Left(Trim(LCase(setLocation(Index).Text)), 7) = "ftps://")) And (Not setLocation(Index).Text = "") Then

                setLocation(Index).Text = IIf(dbSettings.GetProfileSetting("SSL") = 1, "ftps://", "ftp://") & setLocation(Index).Text

                URLType(Index) = nUrl.GetType(setLocation(Index).Text)

            End If

            If (Not setLocation(Index).Text = "") Then
                myClient.URL = setLocation(Index).Text
            End If

            If TypeName(SiteInfo) = "SiteInformation" Then
                If Not (SiteInfo.sUserName.Text = "") Then myClient.Username = SiteInfo.sUserName.Text
                If Not (SiteInfo.sPassword.Text = "") Then myClient.Password = SiteInfo.sPassword.Text
                If IsNumeric(SiteInfo.sPort.Text) Then myClient.Port = CLng(SiteInfo.sPort.Text)
                myClient.ConnectionMode = IIf((SiteInfo.sPassive.Value = 1), "PASV", "PORT")
                myClient.DataPortRange = SiteInfo.sPortRange.Text
                myClient.NetAdapter = dbSettings.GetProfileSetting("AdapterIndex")
                myClient.ImplicitSSL = (SiteInfo.sSSL.Value = 1)
            End If

            If myClient.Username = "" Or myClient.Password = "" Then
                Dim openForm As Long
                openForm = Me.hwnd

                Dim Paswd As frmPassword
                Set Paswd = New frmPassword

                If IsActiveForm(Me) Then
 
                    Paswd.sInfo.Reset
                        
                    Paswd.sInfo.sHostURL.Text = setLocation(Index).Text
                    Paswd.sInfo.sPort.Text = CStr(myClient.Port)
                    Paswd.sInfo.sPassive.Value = -CInt(CBool(myClient.ConnectionMode = "PASV"))
                    Paswd.sInfo.sPortRange.Text = myClient.DataPortRange
                    Paswd.sInfo.sAdapter.ListIndex = (myClient.NetAdapter - 1)
                    Paswd.sInfo.sSSL.Value = -CInt(myClient.ImplicitSSL)
                    
                    LoadCache Paswd.sInfo

                    setLocation(Index).Text = Paswd.sInfo.sHostURL.Text
                    URLType(Index) = nUrl.GetType(Paswd.sInfo.sHostURL.Text)
                    myClient.URL = setLocation(Index).Text
                    myClient.Username = Paswd.sInfo.sUserName.Text
                    myClient.Password = Paswd.sInfo.sPassword.Text
                    myClient.Port = CLng(Paswd.sInfo.sPort.Text)
                    myClient.ConnectionMode = IIf(CBool(Paswd.sInfo.sPassive.Value = 1), "PASV", "PORT")
                    myClient.DataPortRange = Paswd.sInfo.sPortRange.Text
                    myClient.NetAdapter = Paswd.sInfo.sAdapter.ListIndex + 1
                    myClient.ImplicitSSL = (Paswd.sInfo.sSSL.Value = 1)
                Else
                    Dim passInfo As PasswordInfo
                    passInfo.HostURL = setLocation(Index).Text
                    If nUrl.GetPort(setLocation(Index).Text) <> 21 Then
                        passInfo.Port = nUrl.GetPort(setLocation(Index).Text)
                    Else
                        passInfo.Port = myClient.Port
                    End If
                    passInfo.Pasv = (myClient.ConnectionMode = "PASV")
                    passInfo.PortRange = myClient.DataPortRange
                    passInfo.Adapter = myClient.NetAdapter - 1
                    passInfo.SSL = myClient.ImplicitSSL
                    Paswd.ParentHWnd = Me.hwnd
                    
                    If Not AutoLogin Then
                        If Not ShowPassword(Paswd, passInfo) Then
                            If Not Paswd Is Nothing Then
                                Unload Paswd
                                Set Paswd = Nothing
                            End If
                            If IsWindow(openForm) Then
        
                                CancelStatusGUI(Index) = False
                                SetClientGUIState Index, st_ProcessFailed
                            End If
        
                            Exit Sub
                        End If
                    End If
                    If FTPCommand(Index) <> "Stop" Then
    
                        setLocation(Index).Text = passInfo.HostURL
                        URLType(Index) = nUrl.GetType(passInfo.HostURL)
                        myClient.URL = passInfo.HostURL
                        myClient.Username = passInfo.Username
                        myClient.Password = passInfo.Password
                        myClient.Port = passInfo.Port
                        myClient.ConnectionMode = IIf(passInfo.Pasv, "PASV", "PORT")
                        myClient.DataPortRange = passInfo.PortRange
                        myClient.NetAdapter = passInfo.Adapter + 1
                        myClient.ImplicitSSL = passInfo.SSL
                    End If
                End If

                If Not Paswd Is Nothing Then
                    Unload Paswd
                    Set Paswd = Nothing
                End If

            End If

            SetStatus Index, "Connecting to " & myClient.Server & "..."

        End If

        If myClient.ConnectedState() = False Then
            MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Connecting to " & myClient.Server, myClient
            myClient.Connect
            MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Connected to " & myClient.Server, myClient
        End If

        If Not setLocation(Index).Text = nUrl.SetFolder(myClient.URL, myClient.Folder) Then
            setLocation(Index).Text = nUrl.SetFolder(myClient.URL, myClient.Folder)
        End If

        
        
        If Not IsActiveForm(Me) And addUrl And (Not dbSettings.GetProfileSetting("HistoryLock")) Then AddAutoTypeURL setLocation(Index).Text

    End If

    If FTPCommand(Index) <> "Stop" And (Not myClient.ConnectedState() = False) Then
        FTPRefresh Index

        If myClient.URLType = URLTypes.File Then
            If IsOnDriveList(pViewDrives(Index), myClient.Folder) > -1 Then
                pViewDrives(Index).ListIndex = IsOnDriveList(pViewDrives(Index), myClient.Folder)
            End If
        End If
    End If

    Set myClient = Nothing
Exit Sub
catcH:
    If FTPCommand(Index) <> "Stop" Then
        'If Me.Visible Then MsgBox Err.Description, vbCritical, AppName
        If Err.Description <> "" Then
            MaxEvents.AddEvent dbSettings, MyDescription, setLocation(Index).Text, "Error: " & Err.Description
        ElseIf myClient.GetLastError <> "" Then
            MaxEvents.AddEvent dbSettings, MyDescription, setLocation(Index).Text, "Error: " & myClient.GetLastError
        End If
    End If
    Err.Clear

    Set myClient = Nothing
    CancelStatusGUI(Index) = False
    SetClientGUIState Index, st_ProcessFailed
End Sub

Public Sub FTPStop(ByVal Index As Integer)

On Error GoTo clearerr

    CloseChildrenWindows
    
    If FTPCommand(Index) = "CopyMove" Then
        'PasteAbort copyToClientForm(copyToClientIndex), copyToClientIndex
       ' If FTPState(Index) = st_Processing Then
            
            PasteAbort Me, Index
            'copyToClientForm(copyToClientIndex(Index)).FTPStop copyToClientIndex(Index)
            'FTPAbort Index
       ' End If
    Else
       ' If FTPState(Index) = st_Processing Then
            'PasteAbort Me, Index
            FTPAbort Index
       ' End If
    End If

    Exit Sub
clearerr:
    If Err Then Err.Clear
    On Error GoTo 0
End Sub

Private Sub CloseChildrenWindows()

    If Not PopUp Is Nothing Then
        If PopUp.Visible Then PopUp.Hide
        PopUp.ParentHWnd = 0
        Set PopUp = Nothing
    End If
       
    Dim frms
    For Each frms In Forms
        If (TypeName(frms) = "frmPassword") Or (TypeName(frms) = "frmOverwrite") Then
            If frms.ParentHWnd = Me.hwnd Then
                frms.ParentHWnd = 0
                Unload frms
            End If
        End If
    Next
    
End Sub

Public Sub FTPDisconnect(ByVal Index As Integer)

    FTPStop Index
    
On Error GoTo clearerr

    If Index = 0 Then
        If Not (myClient0 Is Nothing) Then
            If Not myClient0.ConnectedState() = False Then
                myClient0.Disconnect
                MaxEvents.AddEvent dbSettings, MyDescription, myClient0.URL, "Disconnected", myClient0
            End If
            
        End If
    Else
        If Not (myClient1 Is Nothing) Then
            If Not myClient1.ConnectedState() = False Then
                myClient1.Disconnect
                MaxEvents.AddEvent dbSettings, MyDescription, myClient1.URL, "Disconnected", myClient1
            End If
            
        End If
    End If

    SetClientGUIState Index, st_ProcessFailed

    Exit Sub
clearerr:
    If Err Then Err.Clear
    On Error GoTo 0
End Sub

Public Sub FTPOpen(ByVal Index As Integer, Optional ByVal CacheOnly As Boolean = False)

    FTPUnloading(Index) = False
    
    If pView(Index).SelectedItem Is Nothing Then
        If Me.Visible Then MsgBox "You must select a file or folder to open.", vbInformation, AppName
    Else
        If Left(pView(Index).SelectedItem.Text, 1) = "/" Then
            FTPChangeFolder Index, Mid(pView(Index).SelectedItem.Text, 2)
            
        Else
            Dim URL As New NTAdvFTP61.URL
            
            Dim myClient As NTAdvFTP61.Client
        
            Select Case URLType(Index)
                Case URLTypes.ftp

                        Dim File As String
                        Dim gID As String
                        File = pView(Index).SelectedItem.Text
                        
                        Dim newClient As frmFTPClientGUI
                        Set newClient = New frmFTPClientGUI
                        newClient.LoadClient "Active App Cache "
                        
                        Set myClient = SetMyClient(Index)
                        
                        gID = frmActiveCache.ExistsInCache(File, setLocation(Index).Text, myClient.Username, myClient.Password)
                        
                        newClient.setLocation(0).Text = Replace(GetTemporaryFolder & "\" & ActiveAppFolder & gID & "\", "\\", "\")
                        newClient.FTPConnect 0
                        
                        CopyFiles Me, Index, "Copy"
                        
                        PasteFromClipboard newClient, 0
                        
                        Do While newClient.GetState(0) = st_Processing
                            modCommon.DoTasks
                        Loop

                        If PathExists(Replace(GetTemporaryFolder & "\" & ActiveAppFolder & gID & "\" & File, "\\", "\"), True) Then
                            frmActiveCache.AddToCache File, setLocation(Index).Text, myClient.Username, myClient.Password, gID
 
                        End If
                        Set myClient = Nothing
                                    
                        If dbSettings.GetClientSetting("ActiveAppOpen") And ((Not frmActiveCache.Visible) And (Not (frmActiveCache.WindowState = 1))) Then
                            frmActiveCache.ShowForm
                        ElseIf (((Not frmActiveCache.Visible) And (Not (frmActiveCache.WindowState = 1)))) Then
                            Unload frmActiveCache
                        End If
                        
                        Unload newClient
                        Set newClient = Nothing

                        If dbSettings.GetClientSetting("ActiveAppRun") Then
                            OpenAssociatedFile Replace(GetTemporaryFolder & "\" & ActiveAppFolder & gID & "\" & File, "\\", "\"), False
                        End If
                            
                Case Else
                
                    Set myClient = SetMyClient(Index)
                    If Right(myClient.Folder, 1) = "\" Then
                        OpenAssociatedFile myClient.Folder & pView(Index).SelectedItem.Text, False
                    Else
                        OpenAssociatedFile myClient.Folder & "\" & pView(Index).SelectedItem.Text, False
                    End If
                    
            End Select
            Set URL = Nothing
            Set myClient = Nothing
        
        End If
    End If

End Sub

Public Sub FTPAbort(ByVal Index As Integer)
    
    Dim oldCommand As String
    oldCommand = FTPCommand(Index)
    
    FTPUnloading(Index) = True
    FTPCommand(Index) = "Stop"

    Dim myClient As NTAdvFTP61.Client
    Set myClient = SetMyClient(Index)
    
    myClient.CancelTransfer

    Set myClient = Nothing
      
    CloseChildrenWindows
   
    SetCaption
    
    CancelStatusGUI(Index) = False
    InRecursiveAction(Index) = False
    SetClientGUIState Index, st_ProcessFailed

    If oldCommand = "CopyMove" Then
        copyToClientForm(Index).SetInRecursiveAction copyToClientIndex(Index), False
        copyToClientForm(Index).SetCancelStatusGUI copyToClientIndex(Index), False
        copyToClientForm(Index).SetClientGUIState copyToClientIndex(Index), st_ProcessFailed
    End If

End Sub

Private Sub ClientError(ByVal Index As Integer, ByVal Number As Long, ByVal Source As String, ByVal Description As String)

    Dim myClient As NTAdvFTP61.Client
    Set myClient = SetMyClient(Index)
    
    If FTPCommand(Index) <> "Stop" Then
        
        MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Error: " & Description, myClient
        'FTPStop Index
    
        If Me.Visible And Not Description = "Disconnected..." Then
            FTPStop Index
            Set PopUp = Nothing
            
            Set PopUp = New NTPopup21.Window
            PopUp.ParentHWnd = Me.hwnd
    
            PopUp.AlwaysOnTop = True
            PopUp.Message = Replace(Replace(Replace(Replace(Description, vbCrLf, vbLf), vbCr, "\n"), vbLf, "\n"), "\n", vbCrLf)
            PopUp.Title = AppName
            PopUp.Icon = vbInformation
            PopUp.LinkText = "Max-FTP Server-Side Error"
            PopUp.Visible = True
        ElseIf Description = "Disconnected..." Then
            WriteToLogView pDummyView(Index), MessageTypes.Incorrect, Description
            If FTPCommand(Index) = "Refresh" Then
                FTPDisconnect Index
                
                FTPConnect Index  ', True
                Exit Sub
            Else
                FTPDisconnect Index
            End If
        End If
        
    End If

    CancelStatusGUI(Index) = False
    InRecursiveAction(Index) = False
    SetClientGUIState Index, st_ProcessFailed

    Set myClient = Nothing

End Sub

Public Sub FTPRefresh(ByVal Index As Integer)
On Error GoTo catcH

    If InRecursiveAction(Index) = False Then
        FTPUnloading(Index) = True
        FTPCommand(Index) = "Refresh"
    End If

    SetClientGUIState Index, st_Processing
    
    Dim myClient As NTAdvFTP61.Client
    Set myClient = SetMyClient(Index)
    
    SetStatus Index, "Listing Folder"

    myClient.ListContents tmpListFile(Index)
    
    Do While InRecursiveAction(Index) And GetState(Index) = st_Processing
        DoEvents
    Loop
    
    If GetState(Index) = st_ProcessFailed Then
        FTPConnect Index
    End If
    
    Debug.Print GetState(Index)
    

    MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Listing of " & myClient.Folder, myClient
    
    Set myClient = Nothing

Exit Sub
catcH:

    If FTPCommand(Index) <> "Stop" Then
        'If Me.Visible Then MsgBox Err.Description, vbCritical, AppName
        If Err.Description <> "" Then
            MaxEvents.AddEvent dbSettings, MyDescription, setLocation(Index).Text, "Error: " & Err.Description
        ElseIf myClient.GetLastError <> "" Then
            MaxEvents.AddEvent dbSettings, MyDescription, setLocation(Index).Text, "Error: " & myClient.GetLastError
        End If
    End If
    Err.Clear
    CancelStatusGUI(Index) = False
    SetClientGUIState Index, st_ProcessFailed
    Set myClient = Nothing
End Sub

Public Sub ClientListProgress(ByVal Index As Integer, ByVal ReceivedBytes As Double)
On Error GoTo catcH
    
        SetStatus Index, "Listing Folder " & ReceivedBytes & " Bytes"

Exit Sub
catcH:
    If FTPCommand(Index) <> "Stop" Then
        'If Me.Visible Then MsgBox Err.Description, vbCritical, AppName
    End If
    Err.Clear
    CancelStatusGUI(Index) = False
    InRecursiveAction(Index) = False
    SetClientGUIState Index, st_ProcessFailed
End Sub

Public Sub ClientListComplete(ByVal Index As Integer)
On Error GoTo catcH
 Debug.Print "ClientListComplete"
    Dim fullText As String
    Dim ListItems() As String
    
    If PathExists(tmpListFile(Index), True) Then fullText = ReadFile(tmpListFile(Index))
    
    Dim myClient As NTAdvFTP61.Client
    Set myClient = SetMyClient(Index)
    
    SetStatus Index, "Reading Contents..."
    
    myClient.ParseListing fullText, ListItems()
    
    RefreshFileView Index, ListItems()
    
    Dim nUrl As New NTAdvFTP61.URL
    
    setLocation(Index).Text = nUrl.SetFolder(setLocation(Index).Text, myClient.Folder)
    
    AddToHistory Index, setLocation(Index).Text
    
    MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Listing of " & myClient.Folder, myClient

    AllowAddToHistory = True

    SetClientGUIState Index, st_ProcessSuccess

    If PathExists(tmpListFile(Index), True) Then Kill tmpListFile(Index)

    Set myClient = Nothing
    Set nUrl = Nothing

Exit Sub
catcH:

    If FTPCommand(Index) <> "Stop" Then
       ' If Me.Visible Then MsgBox Err.Description, vbCritical, AppName
        If Err.Description <> "" Then
            MaxEvents.AddEvent dbSettings, MyDescription, setLocation(Index).Text, "Error: " & Err.Description
        ElseIf myClient.GetLastError <> "" Then
            MaxEvents.AddEvent dbSettings, MyDescription, setLocation(Index).Text, "Error: " & myClient.GetLastError
        End If
    End If
    If Err Then Err.Clear
    CancelStatusGUI(Index) = False
    InRecursiveAction(Index) = False
    SetClientGUIState Index, st_ProcessFailed
End Sub


Public Sub RecursiveClientListComplete(ByVal Index As Integer)
On Error GoTo catcH
 Debug.Print "RecursiveClientListComplete"
    Dim fullText As String
    
    If PathExists(tmpListFile(Index), True) Then fullText = ReadFile(tmpListFile(Index))
    
    Dim myClient As NTAdvFTP61.Client
    Set myClient = SetMyClient(Index)
    
    SetStatus Index, "Reading Contents..."
    
    If Index = 0 Then
        myClient.ParseListing fullText, RecursiveListItems0
    Else
        myClient.ParseListing fullText, RecursiveListItems1
    End If
    
    AllowAddToHistory = True

    If PathExists(tmpListFile(Index), True) Then Kill tmpListFile(Index)

    InRecursiveAction(Index) = False

    Set myClient = Nothing

Exit Sub
catcH:
    If FTPCommand(Index) <> "Stop" Then
       ' If Me.Visible Then MsgBox Err.Description, vbCritical, AppName
        If Err.Description <> "" Then
            MaxEvents.AddEvent dbSettings, MyDescription, setLocation(Index).Text, "Error: " & Err.Description
        ElseIf myClient.GetLastError <> "" Then
            MaxEvents.AddEvent dbSettings, MyDescription, setLocation(Index).Text, "Error: " & myClient.GetLastError
        End If
    End If
    If Err Then Err.Clear
    CancelStatusGUI(Index) = False
    InRecursiveAction(Index) = False
    SetClientGUIState Index, st_ProcessFailed
End Sub

Public Sub ClientFileProgress(ByVal Index As Integer, ByVal ProgressType As NTAdvFTP61.ProgressTypes, ByVal ReceivedBytes As Double)


    If RecursiveFileSize(Index) > 0 Then
        SetCaption RecursiveFileName(Index), CLng(CDbl(ReceivedBytes / RecursiveFileSize(Index)) * 100)
    End If

    If (Timer - ftpTimer) >= 1 Then
        ftpTimer = Timer
        
        If PreviousFileSize(Index) > 0 Then
            If PreviousFileSize(Index) < ReceivedBytes Then
                Dim newRate As Double
                newRate = CLng(CDbl(ReceivedBytes - PreviousFileSize(Index)) / 1000)
                If newRate > 0 Then
                    If ftpRates(Index, 0) = 0 And ftpRates(Index, 1) = 0 And ftpRates(Index, 2) = 0 Then
                        ftpRates(Index, 0) = newRate
                    ElseIf ftpRates(Index, 1) = 0 And ftpRates(Index, 2) = 0 Then
                        ftpRates(Index, 1) = newRate
                    ElseIf ftpRates(Index, 2) = 0 Then
                        ftpRates(Index, 2) = newRate
                        newRate = CLng(CDbl(ftpRates(Index, 0) + ftpRates(Index, 1) + ftpRates(Index, 2)) / 3)
                        ftpRates(Index, 0) = 0
                    ElseIf ftpRates(Index, 1) = 0 Then
                        ftpRates(Index, 1) = newRate
                        newRate = CLng(CDbl(ftpRates(Index, 0) + ftpRates(Index, 1) + ftpRates(Index, 2)) / 3)
                        ftpRates(Index, 2) = 0
                    ElseIf ftpRates(Index, 0) = 0 Then
                        ftpRates(Index, 0) = newRate
                        newRate = CLng(CDbl(ftpRates(Index, 0) + ftpRates(Index, 1) + ftpRates(Index, 2)) / 3)
                        ftpRates(Index, 1) = 0
                    End If

                    PreviousFileRate(Index) = ",  " & newRate & " KB/Sec"
                      
                End If
            End If
        End If
    
        PreviousFileSize(Index) = ReceivedBytes
        
      '  Sleep ((Timer - ftpTimer) / 1000) * 0.1

    End If

    Dim ProgressDisplay As String
    
    If ((ProgressType And NTAdvFTP61.ProgressTypes.TransferingFile) = NTAdvFTP61.ProgressTypes.TransferingFile) Then
        If RecursiveFileSize(Index) = -1 Then
            copyToClientForm(Index).SetStatus copyToClientIndex(Index), "Received " & Trim(FormatFileSize(ReceivedBytes)) & " of Unknown Size" & FormatFileSize(PreviousFileRate(Index))
            SetProgress Index, "Sent " & Trim(FormatFileSize(ReceivedBytes)) & " of Unknown Size" & FormatFileSize(PreviousFileRate(Index))
        Else
            copyToClientForm(Index).SetProgress copyToClientIndex(Index), ReceivedBytes, RecursiveFileSize(Index)
            SetStatus Index, "Sent " & Trim(FormatFileSize(ReceivedBytes)) & " of " & Trim(FormatFileSize(RecursiveFileSize(Index))) & PreviousFileRate(Index)
        End If
    ElseIf ((ProgressType And NTAdvFTP61.ProgressTypes.PositioningFile) = NTAdvFTP61.ProgressTypes.PositioningFile) Then
        'resuming status
        If RecursiveFileSize(Index) = -1 Then
            copyToClientForm(Index).SetStatus copyToClientIndex(Index), "Waiting with " & Trim(FormatFileSize(ReceivedBytes)) & " of Unknown Size" & FormatFileSize(PreviousFileRate(Index))
            SetProgress Index, "Pointing at " & Trim(FormatFileSize(ReceivedBytes)) & " of Unknown Size" & FormatFileSize(PreviousFileRate(Index))
        Else
            copyToClientForm(Index).SetProgress copyToClientIndex(Index), ReceivedBytes, RecursiveFileSize(Index)
            SetStatus Index, "Pointing at " & Trim(FormatFileSize(ReceivedBytes)) & " of " & Trim(FormatFileSize(RecursiveFileSize(Index))) & PreviousFileRate(Index)
        End If
    ElseIf ((ProgressType And NTAdvFTP61.ProgressTypes.AllocatingFile) = NTAdvFTP61.ProgressTypes.AllocatingFile) Then
        'allocating status
        If RecursiveFileSize(Index) = -1 Then
            copyToClientForm(Index).SetStatus copyToClientIndex(Index), "Allocated " & Trim(FormatFileSize(ReceivedBytes)) & " of Unknown Size" & FormatFileSize(PreviousFileRate(Index))
            SetProgress Index, "Allocating " & Trim(FormatFileSize(ReceivedBytes)) & " of Unknown Size" & FormatFileSize(PreviousFileRate(Index))
        Else
            copyToClientForm(Index).SetProgress copyToClientIndex(Index), ReceivedBytes, RecursiveFileSize(Index)
            SetStatus Index, "Allocating " & Trim(FormatFileSize(ReceivedBytes)) & " of " & Trim(FormatFileSize(RecursiveFileSize(Index))) & PreviousFileRate(Index)
        End If
    End If

        
End Sub

Public Sub RecursiveClientFileComplete(ByVal Index As Integer)
On Error GoTo catcH:

    copyToClientForm(Index).SetStatus copyToClientIndex(Index), "Receive Complete " & RecursiveFileName(Index)
    SetStatus Index, "Send Complete " & RecursiveFileName(Index)

    InRecursiveAction(Index) = False
    PreviousFileRate(Index) = ""
    PreviousFileSize(Index) = 0
    ftpRates(Index, 0) = 0
    ftpRates(Index, 1) = 0
    ftpRates(Index, 2) = 0
    If copyToClientIndex(Index) = 0 Then
        MaxEvents.AddEvent dbSettings, MyDescription, copyToClientForm(Index).setLocation(copyToClientIndex(Index)).Text, "Transfered " & RecursiveFileName(Index), copyToClientForm(Index).myClient0
    Else
        MaxEvents.AddEvent dbSettings, MyDescription, copyToClientForm(Index).setLocation(copyToClientIndex(Index)).Text, "Transfered " & RecursiveFileName(Index), copyToClientForm(Index).myClient1
    End If

Exit Sub
catcH:

    If FTPCommand(Index) <> "Stop" Then
        If Me.Visible Then MsgBox Err.Description, vbCritical, AppName
        If copyToClientIndex(Index) = 0 Then
            MaxEvents.AddEvent dbSettings, MyDescription, copyToClientForm(Index).setLocation(copyToClientIndex(Index)).Text, "Error: " & Err.Description, copyToClientForm(Index).myClient0
        Else
            MaxEvents.AddEvent dbSettings, MyDescription, copyToClientForm(Index).setLocation(copyToClientIndex(Index)).Text, "Error: " & Err.Description, copyToClientForm(Index).myClient1
        End If
    End If
    If Err Then Err.Clear
    CancelStatusGUI(Index) = False
    InRecursiveAction(Index) = False
    SetClientGUIState Index, st_ProcessFailed
End Sub

Public Sub ClientAddToLog(ByVal Index As Integer, ByVal MessageType As LogMsgTypes, ByVal AddedText As String)
'On Error GoTo catch
    
    If ((MessageType And LogMsgTypes.log_Outgoing) = LogMsgTypes.log_Outgoing) Then
        If Left(AddedText, 4) = "PASS" Then AddedText = "PASS " & String(Round(3 + (15 * Rnd), 0), IIf(Round(Rnd, 0) = 0, "*", "."))
    End If
    
    WriteToLogView pDummyView(Index), MessageType, AddedText

'Exit Sub
'catch:
'    If FTPCommand(Index) <> "Stop" Then
'        If Me.Visible Then MsgBox Err.Description, vbCritical, AppName
'    End If
'    Err.Clear
'    CancelStatusGUI(Index) = False
'    InRecursiveAction(Index) = False
'    SetClientGUIState Index, st_ProcessFailed
End Sub

Public Sub FTPChangeFolder(ByVal Index As Integer, ByVal RelativeFolder As String)
On Error GoTo catcH

    FTPUnloading(Index) = False
    FTPCommand(Index) = "ChangeFolder"
    
    SetClientGUIState Index, st_Processing

    SetStatus Index, "Changing to folder " & RelativeFolder & "..."
    
    Dim myClient As NTAdvFTP61.Client
    Set myClient = SetMyClient(Index)
    
    myClient.ChangeFolderRelative RelativeFolder
    
    MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Changed Folder to " & RelativeFolder, myClient
    
    FTPRefresh Index

    Set myClient = Nothing
Exit Sub
catcH:
    If FTPCommand(Index) <> "Stop" Then
        If Me.Visible Then MsgBox Err.Description, vbCritical, AppName
        If myClient.GetLastError <> "" Then
            MaxEvents.AddEvent dbSettings, MyDescription, setLocation(Index).Text, "Error: " & myClient.GetLastError, myClient
        ElseIf Err.Description <> "" Then
            MaxEvents.AddEvent dbSettings, MyDescription, setLocation(Index).Text, "Error: " & Err.Description, myClient
        End If
    End If
    Err.Clear
    CancelStatusGUI(Index) = False
    SetClientGUIState Index, st_ProcessFailed
End Sub

Public Sub FTPRename(ByVal Index As Integer, ByVal OldName As String, ByVal NewName As String)
On Error GoTo catcH

    FTPUnloading(Index) = False
    FTPCommand(Index) = "Rename"
    
    SetClientGUIState Index, st_Processing
    SetStatus Index, "Renaming " & OldName & " to " & NewName & "..."
    
    Dim myClient As NTAdvFTP61.Client
    Set myClient = SetMyClient(Index)
    
    If Left(OldName, 1) = "/" Then
        OldName = Mid(OldName, 2)
    End If
    If Left(NewName, 1) = "/" Then
        NewName = Mid(NewName, 2)
    End If
        
    myClient.Rename OldName, NewName

    MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Renamed " & OldName & " to " & NewName, myClient

    SetClientGUIState Index, st_ProcessSuccess

    Set myClient = Nothing
Exit Sub
catcH:
    If FTPCommand(Index) <> "Stop" Then
        If Me.Visible Then MsgBox Err.Description, vbCritical, AppName
        If myClient.GetLastError <> "" Then
            MaxEvents.AddEvent dbSettings, MyDescription, setLocation(Index).Text, "Error: " & myClient.GetLastError, myClient
        ElseIf Err.Description <> "" Then
            MaxEvents.AddEvent dbSettings, MyDescription, setLocation(Index).Text, "Error: " & Err.Description, myClient
        End If
    End If
    Err.Clear
    Set myClient = Nothing
    CancelStatusGUI(Index) = False
    SetClientGUIState Index, st_ProcessFailed
End Sub

Private Function GetNewFolderName(ByVal Index As Integer) As String

    Dim cnt As Integer
    cnt = 1
    Do Until IsOnListItems(pView(Index), "/New Folder " & cnt) = 0
        cnt = cnt + 1
    Loop

    GetNewFolderName = "New Folder " & cnt
    
End Function

Public Sub FTPNewFolder(ByVal Index As Integer, ByVal FolderName As String)
On Error GoTo catcH

    FTPUnloading(Index) = False
    FTPCommand(Index) = "NewFolder"
    
    SetClientGUIState Index, st_Processing
    
    SetStatus Index, "Creating New Folder..."
    
    Dim myClient As NTAdvFTP61.Client
    Set myClient = SetMyClient(Index)
        
    myClient.MakeFolder FolderName
    
    Dim newItem
    
    Set newItem = pView(Index).ListItems.Add(, , "/" & FolderName, , "folder")
    newItem.SubItems(1) = "<DIR>"
    
    MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Create Folder " & FolderName, myClient

    SetClientGUIState Index, st_ProcessSuccess
    
    Dim cnt As Integer
    cnt = 1
    
    Do Until cnt > pView(Index).ListItems.count
        pView(Index).ListItems(cnt).Selected = False
        cnt = cnt + 1
    Loop
    
    newItem.Selected = True
    
    If Me.Visible Then
    
        pView(Index).SetFocus
        newItem.EnsureVisible
    
        pView(Index).StartLabelEdit
    
    End If

    Set myClient = Nothing
    
Exit Sub
catcH:
    If FTPCommand(Index) <> "Stop" Then
        If Me.Visible Then MsgBox Err.Description, vbCritical, AppName
        If myClient.GetLastError <> "" Then
            MaxEvents.AddEvent dbSettings, MyDescription, setLocation(Index).Text, "Error: " & myClient.GetLastError, myClient
        ElseIf Err.Description <> "" Then
            MaxEvents.AddEvent dbSettings, MyDescription, setLocation(Index).Text, "Error: " & Err.Description, myClient
        End If
    End If
    Err.Clear
    CancelStatusGUI(Index) = False
    SetClientGUIState Index, st_ProcessFailed
End Sub

Public Sub FTPDelete(ByVal Index As Integer)
On Error GoTo catcH

    FTPUnloading(Index) = False
    FTPCommand(Index) = "Delete"
    
    If pView(Index).SelectedItem Is Nothing Then
        If Me.Visible Then MsgBox "You must first select the files or folders you want to delete.", vbInformation, AppName
    Else
        If Me.Visible Then
            If MsgBox("Are you sure you want to delete the selected items?  This will remove files under folders you may have selected.", vbQuestion + vbYesNo, AppName) = vbNo Then
                Exit Sub
            End If
        End If

        SetClientGUIState Index, st_Processing

        SetStatus Index, "Deleting..."

        Dim myClient As NTAdvFTP61.Client
        Set myClient = SetMyClient(Index)

        MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Begin Delete", myClient

        Dim newList As New NTNodes10.Collection
        Dim cnt As Integer
        
        cnt = 1
        Do Until cnt > pView(Index).ListItems.count

            If pView(Index).ListItems(cnt).Selected Then
                newList.Add pView(Index).ListItems(cnt).Text
            End If
            
            cnt = cnt + 1
        Loop

        cnt = 1
        Do Until cnt > newList.count

            CancelStatusGUI(Index) = True
            RecursiveAction Index, myClient, "delete", newList(cnt)
            CancelStatusGUI(Index) = False
                
            cnt = cnt + 1
        Loop

        Set newList = Nothing

        MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Finish Delete", myClient

        FTPRefresh Index
    
        Set myClient = Nothing
        
    End If

Exit Sub
catcH:
    If FTPCommand(Index) <> "Stop" Then
        If Me.Visible Then MsgBox Err.Description, vbCritical, AppName
        If myClient.GetLastError <> "" Then
            MaxEvents.AddEvent dbSettings, MyDescription, setLocation(Index).Text, "Error: " & myClient.GetLastError, myClient
        ElseIf Err.Description <> "" Then
            MaxEvents.AddEvent dbSettings, MyDescription, setLocation(Index).Text, "Error: " & Err.Description, myClient
        End If
    End If
    Err.Clear
        Set myClient = Nothing
    CancelStatusGUI(Index) = False
    SetClientGUIState Index, st_ProcessFailed
End Sub

Public Sub FTPCopyMove(ByVal Action As String, ByVal Index As Integer, ByVal copyToForm, ByVal copyToIndex As Integer, ByVal newList As NTNodes10.Collection)
On Error Resume Next
    
    FTPUnloading(Index) = False
    FTPCommand(Index) = "CopyMove"
    copyToForm.SetFTPCommand copyToIndex, "CopyMove"
    
    Action = LCase(Trim(Action))
    
            SetClientGUIState Index, st_Processing
            If LCase(Action) = "copy" Then
                SetStatus Index, "Copying..."
            Else
                SetStatus Index, "Moving..."
            End If
            CancelStatusGUI(Index) = True
    
            copyToForm.SetClientGUIState copyToIndex, st_Processing
            copyToForm.SetCancelStatusGUI copyToIndex, True

            Set copyToClientForm(Index) = copyToForm
            copyToClientIndex(Index) = copyToIndex
            copyIsSource(Index) = True
            
            Set copyToForm.copyToClientForm(copyToIndex) = Me
            copyToForm.copyToClientIndex(copyToIndex) = Index
            copyToForm.copyIsSource(copyToIndex) = False
    
            Dim myClient As NTAdvFTP61.Client
            Set myClient = SetMyClient(Index)

            If LCase(Action) = "copy" Or LCase(Action) = "move" Then

                If (LCase(setLocation(Index).Text) = LCase(copyToForm.setLocation(copyToIndex).Text)) Then
                    Err.Raise 75, App.EXEName, "Unable to copy files to themselves." ', AppName

                End If

            End If
    
            MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Begin " & Action, myClient
            
            Dim oOverwrite As OverwriteTypes
            
            Select Case Left(LCase(MyDescription), 4)
                Case "acti"
                    oOverwrite = or_YesToAll
                Case "clie"
                    Select Case dbSettings.GetClientSetting("Overwrite")
                        Case 1
                            oOverwrite = or_NoToAll
                        Case 2
                            oOverwrite = or_YesToAll
                        Case 3
                            oOverwrite = or_AutoAll
                        Case Else
                            oOverwrite = or_Prompt
                    End Select
                Case Else
                    oOverwrite = or_YesToAll
            End Select

            Dim newToList As NTNodes10.Collection
            If copyToIndex = 0 Then
                Set newToList = RecursiveGetCopyToCollection(Index, copyToForm.myClient0)
            Else
                Set newToList = RecursiveGetCopyToCollection(Index, copyToForm.myClient1)
            End If
            
            Dim cnt As Integer
            cnt = 1
            Do Until cnt > newList.count
                
                If copyToIndex = 0 Then
                    RecursiveAction Index, myClient, Action, newList(cnt), oOverwrite, copyToForm.myClient0, newToList
                Else
                    RecursiveAction Index, myClient, Action, newList(cnt), oOverwrite, copyToForm.myClient1, newToList
                End If
                    
                cnt = cnt + 1
            Loop

            Set newToList = Nothing
            
            CancelStatusGUI(Index) = False
            copyToForm.SetCancelStatusGUI copyToIndex, False

            Set newList = Nothing
            
If FTPUnloading(Index) Then Exit Sub

            MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Finish " & Action, myClient
        
            SetInRecursiveAction Index, False
            FTPRefresh Index
            
            copyToClientForm(Index).SetInRecursiveAction copyToClientIndex(Index), False
            
            copyToClientForm(Index).FTPRefresh copyToClientIndex(Index)
            
            
            Set copyToClientForm(Index).copyToClientForm(copyToClientIndex(Index)) = Nothing
            Set copyToClientForm(Index) = Nothing

        
            Set myClient = Nothing

Exit Sub
catcH:

        If FTPCommand(Index) <> "Stop" Then
            If Me.Visible Then MsgBox Err.Description, vbCritical, AppName
            If myClient.GetLastError <> "" Then
                MaxEvents.AddEvent dbSettings, MyDescription, setLocation(Index).Text, "Error: " & myClient.GetLastError, myClient
            ElseIf Err.Description <> "" Then
                MaxEvents.AddEvent dbSettings, MyDescription, setLocation(Index).Text, "Error: " & Err.Description, myClient
            End If

            Err.Clear
            
            InRecursiveAction(Index) = False
            CancelStatusGUI(Index) = False
            SetClientGUIState Index, st_ProcessFailed
        
            copyToForm.SetInRecursiveAction copyToIndex, False
            copyToForm.SetCancelStatusGUI copyToIndex, False
            copyToForm.SetClientGUIState copyToIndex, st_ProcessFailed
        Else
            Err.Clear
            
            InRecursiveAction(Index) = False
            CancelStatusGUI(Index) = False
            SetClientGUIState Index, st_Finished
            
            copyToForm.SetInRecursiveAction copyToIndex, False
            copyToForm.SetCancelStatusGUI copyToIndex, False
            copyToForm.SetClientGUIState copyToIndex, st_Finished
        End If


        Set copyToClientForm(Index).copyToClientForm(copyToClientIndex(Index)) = Nothing
        Set copyToClientForm(Index) = Nothing
        
End Sub

Public Function RecursiveGetCopyToCollection(ByVal Index As Integer, ByRef myClient As NTAdvFTP61.Client) As NTNodes10.Collection

    copyToClientForm(Index).RecursiveResetArray copyToClientIndex(Index)
    copyToClientForm(Index).SetInRecursiveAction copyToClientIndex(Index), True

    myClient.ListContents tmpListFile(copyToClientIndex(Index))
    Do Until copyToClientForm(Index).InRecursion(copyToClientIndex(Index)) = False
        DoEvents
        Sleep 1
    Loop

    Dim newToList As New NTNodes10.Collection
    copyToClientForm(Index).RecursiveLoopArrayToCollection copyToClientIndex(Index), newToList
    copyToClientForm(Index).RecursiveResetArray copyToClientIndex(Index)

    Set RecursiveGetCopyToCollection = newToList.Clone
    Set newToList = Nothing

End Function

Public Function RecursiveGetCopyCollection(ByVal Index As Integer, ByRef myClient As NTAdvFTP61.Client) As NTNodes10.Collection

    RecursiveResetArray Index
    SetInRecursiveAction Index, True

    myClient.ListContents tmpListFile(Index)
    Do Until InRecursion(Index) = False
        DoEvents
        Sleep 1
    Loop

    Dim newToList As New NTNodes10.Collection
    RecursiveLoopArrayToCollection Index, newToList
    RecursiveResetArray Index

    Set RecursiveGetCopyCollection = newToList.Clone
    Set newToList = Nothing

End Function
Private Sub RecursiveAction(ByVal Index As Integer, ByRef myClient As NTAdvFTP61.Client, ByVal Action As String, ByVal aItem As String, Optional ByRef oOverwrite As OverwriteTypes, Optional ByRef copyToClient As NTAdvFTP61.Client, Optional ByVal copyToList As NTNodes10.Collection)
On Error GoTo catcH

    Dim newToList As NTNodes10.Collection
    Dim newList As NTNodes10.Collection
    Dim fa As clsFileAssoc

    Dim ItmIndex As Integer
    Dim IsFolder As Boolean
    Dim ItemName As String
    Dim ItemSize As String
    Dim ItemDate As String
    Dim DestItemSize As String
    Dim DestItemDate As String

    If Left(aItem, 1) = "/" Then
        IsFolder = True
        aItem = Mid(aItem, 2)
    Else
        IsFolder = False
    End If
    ItemName = RemoveNextArg(aItem, "|")
    ItemSize = RemoveNextArg(aItem, "|")
    ItemDate = RemoveNextArg(aItem, "|")

If FTPUnloading(Index) Then GoTo catcH

    If IsFolder Then
        Dim cnt As Integer
        SetStatus Index, "Changing Folder to " & ItemName & "..."
        myClient.ChangeFolderRelative ItemName
        MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Changed Folder to " & ItemName, myClient
        Select Case LCase(Trim(Action))
            Case "copy", "move"

                ItmIndex = IsFileOnCollection(copyToList, "/" & ItemName)
                If ItmIndex = -1 Then
                    copyToClient.MakeFolder ItemName
                    MaxEvents.AddEvent dbSettings, MyDescription, copyToClient.URL, "Created Folder " & ItemName, myClient
                End If
                copyToClient.ChangeFolderRelative ItemName
                MaxEvents.AddEvent dbSettings, MyDescription, copyToClient.URL, "Changed Folder to " & ItemName, myClient

                Set newToList = RecursiveGetCopyToCollection(Index, copyToClient)

        End Select

If FTPUnloading(Index) Then GoTo catcH


        Set newList = RecursiveGetCopyCollection(Index, myClient)

'
'        RecursiveResetArray Index
'
'        InRecursiveAction(Index) = True
'
'        FTPRefresh Index
'
'        Do While InRecursiveAction(Index)
'            DoTasks
'
'        Loop
'
'If FTPUnloading(Index) Then GoTo catch
'
'        RecursiveLoopArrayToCollection Index, newList
'
'        RecursiveResetArray Index

        cnt = 1
        Do Until cnt > newList.count

            Select Case LCase(Trim(Action))
                Case "delete"
                    RecursiveAction Index, myClient, Action, newList(cnt)
                Case "copy", "move"
                    RecursiveAction Index, myClient, Action, newList(cnt), oOverwrite, copyToClient, newToList
            End Select

            cnt = cnt + 1
        Loop

If FTPUnloading(Index) Then GoTo catcH

        myClient.ChangeFolderRelative ".."
        MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Parent Folder", myClient
        Select Case LCase(Trim(Action))
            Case "copy", "move"
                copyToClient.ChangeFolderRelative ".."
                MaxEvents.AddEvent dbSettings, MyDescription, copyToClient.URL, "Parent Folder", myClient
        End Select

    End If

If FTPUnloading(Index) Then GoTo catcH

    If IsFolder Then
        Select Case LCase(Trim(Action))
            Case "delete", "move"
                SetStatus Index, "Removing Folder " & ItemName
                myClient.RemoveFolder ItemName
                MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Removed Folder " & ItemName, myClient
        End Select
    Else
        Dim FileWasTransfered As Boolean
        FileWasTransfered = False

        Select Case LCase(Trim(Action))
            Case "copy", "move"

                If CDec(CDbl(ItemSize)) > CDec(CDbl(modBitValue.HighBound())) Then
                    Err.Raise 8, App.EXEName, "Transfering for file sizes over " & CDec(CDbl(modBitValue.HighBound())) & " bytes not implementated."
                End If

                ItmIndex = IsFileOnCollection(copyToList, ItemName, DestItemSize, DestItemDate)
                If (ItmIndex > -1) And ((oOverwrite = or_Prompt) Or (oOverwrite = or_Yes) Or (oOverwrite = or_No)) Then

                    If oOverwrite = or_Prompt Then

                        Dim openForm As Long
                        openForm = Me.hwnd

                        Dim OWF As frmOverwrite
                        Set OWF = New frmOverwrite

                        OWF.ParentHWnd = Me.hwnd

                        If Not ShowOverwrite(OWF, oOverwrite, ItemName & "      FileSize: " & GetFileSizeOnCollection(copyToList, ItmIndex) & "      FileDate: " & GetFileDateOnCollection(copyToList, ItmIndex), ItemName & "      FileSize: " & ItemSize & "      FileDate: " & ItemDate) Then
                            If Not OWF Is Nothing Then
                                Unload OWF
                                Set OWF = Nothing
                            End If

                            If IsWindow(openForm) Then
                                If oOverwrite = or_Cancel Then
                                    FTPUnloading(Index) = True
                                    FTPCommand(Index) = "Stop"
                                    GoTo catcH
                                End If
                                GoTo passit
                            End If

                            Exit Sub
                        End If

                        If Not OWF Is Nothing Then
                            Unload OWF
                            Set OWF = Nothing
                        End If

passit:
                    End If

                End If

If FTPUnloading(Index) Then GoTo catcH

                If (oOverwrite = or_Yes Or oOverwrite = or_YesToAll Or oOverwrite = or_Resume Or oOverwrite = or_AutoAll) Or (ItmIndex = -1) Then
                    SetStatus Index, IIf(oOverwrite = or_Resume, "Prepairing resume of " & ItemName & "...", "Starting transfer of " & ItemName & "...")

                    InRecursiveAction(Index) = True

                    If IsNumeric(ItemSize) Then
                        RecursiveFileSize(Index) = CDbl(ItemSize)
                    Else
                        RecursiveFileSize(Index) = -1
                    End If

                    RecursiveFileName(Index) = ItemName

If FTPUnloading(Index) Then GoTo catcH

                    If (CDbl(ItemSize) > modBitValue.LongBound) Or (CDbl(IIf(IsNumeric(DestItemSize), DestItemSize, "0")) > _
                         modBitValue.LongBound) Or dbSettings.GetProfileSetting("LargeFileMode") Then
                        myClient.LargeFileMode = True
                        copyToClient.LargeFileMode = True
                    Else
                        myClient.LargeFileMode = False
                        copyToClient.LargeFileMode = False
                    End If

                    Set fa = New clsFileAssoc

                    myClient.TransferRates(0) = dbSettings.GetProfileSetting("ftpLocalSize")
                    myClient.TransferRates(1) = dbSettings.GetProfileSetting("ftpBufferSize")
                    myClient.TransferRates(2) = dbSettings.GetProfileSetting("ftpPacketSize")
                    copyToClient.TransferRates(0) = dbSettings.GetProfileSetting("ftpLocalSize")
                    copyToClient.TransferRates(1) = dbSettings.GetProfileSetting("ftpBufferSize")
                    copyToClient.TransferRates(2) = dbSettings.GetProfileSetting("ftpPacketSize")
                    myClient.Allocation = IIf(dbSettings.GetProfileSetting("ClientAlloc"), CLng(AllocateSides.Client), 0) + _
                                            IIf(dbSettings.GetProfileSetting("ServerAlloc"), CLng(AllocateSides.Remote), 0)
                    copyToClient.Allocation = IIf(dbSettings.GetProfileSetting("ClientAlloc"), CLng(AllocateSides.Client), 0) + _
                                            IIf(dbSettings.GetProfileSetting("ServerAlloc"), CLng(AllocateSides.Remote), 0)

                    If mnuMultiThread.Checked And mnuMultiThread.Visible Then
                        If oOverwrite = or_Resume Or oOverwrite = or_AutoAll And (Not (ItmIndex = -1)) Then
                            If oOverwrite = or_Resume Then
                                myClient.TransferType = fa.GetTransferType(GetFileExt(ItemName))
                                copyToClient.TransferType = fa.GetTransferType(GetFileExt(ItemName))
                                myClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(ItemName))
                                copyToClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(ItemName))
                                MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Transfering " & ItemName, myClient
                                CreateTransferThread Me, MyDescription, LCase(Action), myClient, copyToClient, ItemName, ItemSize, CDbl(GetFileSizeOnCollection(copyToList, ItmIndex)), CDbl(ItemSize)

                            ElseIf CDbl(GetFileSizeOnCollection(copyToList, ItmIndex)) < CDbl(ItemSize) Then
                                If IsDate(DestItemDate) And IsDate(ItemDate) Then
                                    If (DateDiff("s", DestItemDate, ItemDate) < 0) Then
                                        myClient.TransferType = fa.GetTransferType(GetFileExt(ItemName))
                                        copyToClient.TransferType = fa.GetTransferType(GetFileExt(ItemName))
                                        myClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(ItemName))
                                        copyToClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(ItemName))
                                        MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Transfering " & ItemName, myClient
                                        CreateTransferThread Me, MyDescription, LCase(Action), myClient, copyToClient, ItemName, ItemSize, CDbl(GetFileSizeOnCollection(copyToList, ItmIndex)), CDbl(ItemSize)

                                    End If
                                Else
                                    myClient.TransferType = fa.GetTransferType(GetFileExt(ItemName))
                                    copyToClient.TransferType = fa.GetTransferType(GetFileExt(ItemName))
                                    myClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(ItemName))
                                    copyToClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(ItemName))
                                    MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Transfering " & ItemName, myClient
                                    CreateTransferThread Me, MyDescription, LCase(Action), myClient, copyToClient, ItemName, ItemSize, CDbl(GetFileSizeOnCollection(copyToList, ItmIndex)), CDbl(ItemSize)

                                End If
                            ElseIf oOverwrite = or_AutoAll Then
                                If CDbl(GetFileSizeOnCollection(copyToList, ItmIndex)) <> CDbl(ItemSize) Then
                                    myClient.TransferType = fa.GetTransferType(GetFileExt(ItemName))
                                    copyToClient.TransferType = fa.GetTransferType(GetFileExt(ItemName))
                                    myClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(ItemName))
                                    copyToClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(ItemName))
                                    MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Transfering " & ItemName, myClient
                                    CreateTransferThread Me, MyDescription, LCase(Action), myClient, copyToClient, ItemName, ItemSize, , CDbl(ItemSize)

                                End If
                            Else
                                myClient.TransferType = fa.GetTransferType(GetFileExt(ItemName))
                                copyToClient.TransferType = fa.GetTransferType(GetFileExt(ItemName))
                                myClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(ItemName))
                                copyToClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(ItemName))
                                MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Transfering " & ItemName, myClient
                                CreateTransferThread Me, MyDescription, LCase(Action), myClient, copyToClient, ItemName, ItemSize, , CDbl(ItemSize)

                            End If
                        Else
                            myClient.TransferType = fa.GetTransferType(GetFileExt(ItemName))
                            copyToClient.TransferType = fa.GetTransferType(GetFileExt(ItemName))
                            myClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(ItemName))
                            copyToClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(ItemName))
                            MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Transfering " & ItemName, myClient
                            CreateTransferThread Me, MyDescription, LCase(Action), myClient, copyToClient, ItemName, ItemSize, , CDbl(ItemSize)

                        End If

                        InRecursiveAction(Index) = False

                        FileWasTransfered = False

                    Else

                        If (oOverwrite = or_Resume Or oOverwrite = or_AutoAll) And (Not (ItmIndex = -1)) Then
                            If oOverwrite = or_Resume Then
                                myClient.TransferType = fa.GetTransferType(GetFileExt(ItemName))
                                copyToClient.TransferType = fa.GetTransferType(GetFileExt(ItemName))
                                myClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(ItemName))
                                copyToClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(ItemName))
                                MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Transfering " & ItemName, myClient
                                myClient.TransferFile ItemName, copyToClient, CDbl(GetFileSizeOnCollection(copyToList, ItmIndex)), CDbl(ItemSize)

                            ElseIf CDbl(GetFileSizeOnCollection(copyToList, ItmIndex)) < CDbl(ItemSize) Then
                                If IsDate(DestItemDate) And IsDate(ItemDate) Then
                                    If (DateDiff("s", IIf(IsDate(DestItemDate), DestItemDate, ItemDate), ItemDate) < 0) Then
                                        myClient.TransferType = fa.GetTransferType(GetFileExt(ItemName))
                                        copyToClient.TransferType = fa.GetTransferType(GetFileExt(ItemName))
                                        myClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(ItemName))
                                        copyToClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(ItemName))
                                        MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Transfering " & ItemName, , myClient
                                        myClient.TransferFile ItemName, copyToClient, CDbl(GetFileSizeOnCollection(copyToList, ItmIndex)), CDbl(ItemSize)

                                    Else
                                        InRecursiveAction(Index) = False
                                    End If
                                Else
                                    myClient.TransferType = fa.GetTransferType(GetFileExt(ItemName))
                                    copyToClient.TransferType = fa.GetTransferType(GetFileExt(ItemName))
                                    myClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(ItemName))
                                    copyToClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(ItemName))
                                    MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Transfering " & ItemName, myClient
                                    myClient.TransferFile ItemName, copyToClient, CDbl(GetFileSizeOnCollection(copyToList, ItmIndex)), CDbl(ItemSize)

                                End If
                            ElseIf oOverwrite = or_AutoAll Then
                                If CDbl(GetFileSizeOnCollection(copyToList, ItmIndex)) <> CDbl(ItemSize) Then
                                    myClient.TransferType = fa.GetTransferType(GetFileExt(ItemName))
                                    copyToClient.TransferType = fa.GetTransferType(GetFileExt(ItemName))
                                    myClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(ItemName))
                                    copyToClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(ItemName))
                                    MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Transfered " & ItemName, myClient
                                    myClient.TransferFile ItemName, copyToClient, , CDbl(ItemSize)

                                Else
                                    InRecursiveAction(Index) = False
                                End If
                            Else
                                myClient.TransferType = fa.GetTransferType(GetFileExt(ItemName))
                                copyToClient.TransferType = fa.GetTransferType(GetFileExt(ItemName))
                                myClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(ItemName))
                                copyToClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(ItemName))
                                MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Transfered " & ItemName, myClient
                                myClient.TransferFile ItemName, copyToClient, , CDbl(ItemSize)

                            End If
                        Else
                            myClient.TransferType = fa.GetTransferType(GetFileExt(ItemName))
                            copyToClient.TransferType = fa.GetTransferType(GetFileExt(ItemName))
                            myClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(ItemName))
                            copyToClient.AssumeLineFeed = fa.GetAssumeLineFeed(GetFileExt(ItemName))
                            MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Transfered " & ItemName, myClient
                            myClient.TransferFile ItemName, copyToClient, , CDbl(ItemSize)

                        End If

                        
                        Do While InRecursiveAction(Index) ' And myClient.GetLastError = "" And copyToClient.GetLastError = ""

                            DoTasks
            
                        Loop

                        FileWasTransfered = True

                    End If

                    Set fa = Nothing

If FTPUnloading(Index) Then GoTo catcH

                End If

        End Select

        Select Case Trim(LCase(Action))
            Case "move"

                If FileWasTransfered Then
                    SetStatus Index, "Removing File " & ItemName
                    myClient.RemoveFile ItemName

                    MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Removed File " & ItemName, myClient
                End If

            Case "delete"

                SetStatus Index, "Removing File " & ItemName
                myClient.RemoveFile ItemName

                MaxEvents.AddEvent dbSettings, MyDescription, myClient.URL, "Removed File " & ItemName, myClient

        End Select

        If oOverwrite = or_Yes Or oOverwrite = or_No Or oOverwrite = or_Resume Then
            oOverwrite = or_Prompt
        End If

    End If

catcH:
 '   newToList.Clear
    Set newToList = Nothing
 '   newList.Clear
    Set newList = Nothing
    If Err.Number <> 0 Then
        If oOverwrite = or_AutoAll Then
            Err.Clear
            InRecursiveAction(Index) = False
        Else
            Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
        End If
    End If
End Sub

Private Function IsFileOnCollection(ByRef tList As NTNodes10.Collection, ByVal Item As String, Optional ByRef DestItemSize As String, Optional ByRef DestItemDate As String) As Integer

    Dim cnt As Integer
    Dim ItemFound As Integer
    Dim ItemData As String
    cnt = 1
    ItemFound = -1
    Do Until cnt > tList.count Or ItemFound > -1
        ItemData = tList(cnt)
    
        If LCase(Trim(RemoveNextArg(ItemData, "|"))) = LCase(Trim(Item)) Then
            ItemFound = cnt
            DestItemSize = RemoveNextArg(ItemData, "|")
            DestItemDate = RemoveNextArg(ItemData, "|")
        End If
        cnt = cnt + 1
    Loop
    IsFileOnCollection = ItemFound

End Function

Private Function GetFileSizeOnCollection(ByRef tList As NTNodes10.Collection, ByVal lstIndex As Integer) As String
    Dim ItemData As String
    Dim ItemSize As String

    ItemData = tList(lstIndex)
    ItemSize = RemoveNextArg(ItemData, "|")
    ItemSize = RemoveNextArg(ItemData, "|")
    GetFileSizeOnCollection = ItemSize

End Function

Private Function GetFileDateOnCollection(ByRef tList As NTNodes10.Collection, ByVal lstIndex As Integer) As String
    Dim ItemData As String
    Dim ItemDate As String

    ItemData = tList(lstIndex)
    ItemDate = RemoveNextArg(ItemData, "|")
    ItemDate = RemoveNextArg(ItemData, "|")
    ItemDate = RemoveNextArg(ItemData, "|")
    GetFileDateOnCollection = ItemDate

End Function

Private Function IsOnDriveList(ByVal tList As Control, ByVal DriveLetter As String) As Integer

    Dim cnt As Integer
    Dim ItemFound As Integer
    cnt = 0
    ItemFound = -1
    Do Until cnt = tList.ListCount Or ItemFound > -1
        If LCase(Trim(Left(tList.List(cnt), 2))) = LCase(Trim(Left(DriveLetter, 2))) Then ItemFound = cnt
        cnt = cnt + 1
    Loop
    IsOnDriveList = ItemFound

End Function

Private Sub RefreshThreadMenu()
    Dim AllTransfering As Boolean, AllStopped As Boolean, HasStopped As Boolean, HasFinished As Boolean
    Dim xItem
    AllStopped = True
    AllTransfering = True
    HasStopped = False
    HasFinished = False
    Dim myTransfer As clsTransfer
    Dim cnt As Long
    cnt = 1
    Do Until cnt > ListView1.ListItems.count
        
        Set myTransfer = GetTransferThread(cnt)
        If Not (myTransfer Is Nothing) Then
            If ListView1.ListItems(cnt).Selected Then
                If Not myTransfer.ThreadState = st_Stopped Then
                    AllStopped = False
                End If
                If Not myTransfer.ThreadState = st_Transfering Then
                    AllTransfering = False
                End If
            End If
            If myTransfer.ThreadState = st_Stopped Then
                HasStopped = True
            End If
            If myTransfer.ThreadState = st_Finished Then
                HasFinished = True
            End If
        End If
        cnt = cnt + 1
    Loop
    
    mnuClearFinished.enabled = HasFinished
    mnuClearStopped.enabled = HasStopped
    mnuCancelTransfer.enabled = (Not ListView1.SelectedItem Is Nothing) And AllTransfering
    mnuRetryTransfer.enabled = (Not ListView1.SelectedItem Is Nothing) And AllStopped
    
    
    Me.PopupMenu mnuThread
End Sub

Private Sub CancelAllTransfers()

    Dim cnt As Integer
    cnt = 1
    Do Until cnt > ListView1.ListItems.count
        If Not (GetTransferThread(cnt) Is Nothing) Then
            GetTransferThread(cnt).CancelFileTransfer
            DestroyTransferThread cnt
        End If
        cnt = cnt + 1
    Loop

End Sub
Private Sub mnuClearFinished_Click()
    Dim cnt As Integer
    cnt = 1
    Do Until cnt > ListView1.ListItems.count
        If Not (GetTransferThread(cnt) Is Nothing) Then
            If GetTransferThread(cnt).ThreadState = st_Finished Then
                If GetTransferThread(cnt).GUICancel Then
                    DestroyTransferThread cnt
                    ListView1.ListItems.Remove cnt
                    modCommon.DoTasks
                Else
                    cnt = cnt + 1
                End If
            Else
                cnt = cnt + 1
            End If
        Else
            cnt = cnt + 1
        End If
    Loop
End Sub
Private Sub mnuClearStopped_Click()
    Dim cnt As Integer
    cnt = 1
    Do Until cnt > ListView1.ListItems.count
        If Not (GetTransferThread(cnt) Is Nothing) Then
            If GetTransferThread(cnt).ThreadState = st_Stopped Then
                If GetTransferThread(cnt).GUICancel Then
                    DestroyTransferThread cnt
                    ListView1.ListItems.Remove cnt
                    modCommon.DoTasks
                Else
                    cnt = cnt + 1
                End If
            Else
                cnt = cnt + 1
            End If
        Else
            cnt = cnt + 1
        End If
        
    Loop

End Sub
Private Sub mnuCancelTransfer_Click()
    Dim cnt As Integer
    cnt = 1
    Do Until cnt > ListView1.ListItems.count
        If Not (GetTransferThread(cnt) Is Nothing) Then
            If GetTransferThread(cnt).ThreadState = st_Transfering And ListView1.ListItems(cnt).Selected Then
                GetTransferThread(cnt).GUICancel
            End If
        End If
        cnt = cnt + 1
    Loop
End Sub
Private Sub mnuRetryTransfer_Click()
    Dim cnt As Integer
    cnt = 1
    Do Until cnt > ListView1.ListItems.count
        If Not (GetTransferThread(cnt) Is Nothing) Then
            If GetTransferThread(cnt).ThreadState = st_Stopped And ListView1.ListItems(cnt).Selected Then
                GetTransferThread(cnt).StartTransfer
            End If
        End If
        cnt = cnt + 1
    Loop
    
End Sub

Private Function ThreadStateCount(ByVal tState As Integer) As Integer
    Dim cnt As Integer
    Dim retVal As Integer
    retVal = 0
    For cnt = 1 To ListView1.ListItems.count
        If Not (GetTransferThread(cnt) Is Nothing) Then
            If GetTransferThread(cnt).ThreadState = tState Then
                retVal = retVal + 1
            End If
        End If
    Next
    ThreadStateCount = retVal
End Function
Private Function MessageFormal(ByVal Text As String) As String
    Dim cnt As Long
    For cnt = 1 To Len(Text)
        Select Case Mid(Text, cnt, 1)
            Case "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", _
                "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", _
                "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", ",", " ", "(", ")", "\", "/", "-"

                MessageFormal = MessageFormal & Mid(Text, cnt, 1)
        End Select
    Next
End Function
Public Sub WriteToLogView(ByRef LogView As Variant, ByVal LogType As Integer, ByVal InData As String)
'    On Error GoTo redo
    
    Dim useWrite As Integer
    Static MaxLogBytes As Long
    If MaxLogBytes = 0 Then MaxLogBytes = dbSettings.GetClientSetting("LogFileSize")

    Dim InCmd As String
    useWrite = True
    If useWrite Then
        With LogView
            If Len(.Text) + Len(InData) > MaxLogBytes Then
                .SelStart = 0
                .SelLength = Len(InData)
                .SelText = ""
            End If
                
            Dim Blue As Long
            Dim Green As Long
            Dim Red As Long
            Dim gray As Long
            Dim black As Long
            Blue = GetSkinColor("LogView_IncommingColor")
            Green = GetSkinColor("LogView_OutGoingColor")
            gray = GetSkinColor("LogView_TextColor")
            Red = GetSkinColor("LogView_ErrorColor")
            black = GetSkinColor("LogView_HighlightColor")
        
            Select Case LogType
                Case log_Outgoing
                    If InStr(InData, " ") > 0 Then
                        InCmd = Left(InData, InStr(InData, " ") - 1)
                        InData = " " & Mid(InData, InStr(InData, " ") + 1) & vbLf
                    Else
                        InCmd = InData
                        InData = vbLf
                    End If
                    
                    .SelStart = Len(.Text)
                    .SelLength = 0
                    .SelColor = Green
                    .SelText = InCmd
                    
                    Select Case (InCmd)
                        Case "RETR", "STOR", "APPE", "RNFR", "RNTO", "CWD", "PWD", "MKD", "DELE", "RMD", "NLST", "LIST", "ALLO", "SMNT"
                            .SelStart = Len(.Text)
                            .SelLength = 0
                            .SelColor = black
                            .SelText = InData
                        
                        Case Else
                        
                            .SelStart = Len(.Text)
                            .SelLength = 0
                            .SelColor = gray
                            .SelText = InData
                    End Select
                
                Case log_Incomming
                    InCmd = Left(InData, 3)
                    If IsNumeric(InCmd) Then
                        InData = Mid(InData, 4) & vbLf
                    
                        .SelStart = Len(.Text)
                        .SelLength = 0
                        .SelColor = Blue
                        .SelText = InCmd
                    Else
                        InData = InData & vbLf
                    End If
                    
                    .SelStart = Len(.Text)
                    .SelLength = 0
                    .SelColor = gray
                    .SelText = InData
                    
                Case log_Error
                    .SelStart = Len(.Text)
                    .SelLength = 0
                    .SelColor = Red
                    .SelText = InData & vbLf
            
            End Select
            
        End With
    End If
'Exit Sub
'redo:
'    Err.Clear
'    Resume
End Sub

Public Sub SetTooltip()
    With Me
   
        Dim cnt As Integer
        If dbSettings.GetProfileSetting("ViewToolTips") Then
            .mnuShowToolTips.Checked = True
            
            For cnt = 0 To -dbSettings.GetClientSetting("ViewDoubleWindow")
                .Picture1(cnt).ToolTipText = "This displays any logging with the server."
                .Image1(cnt).ToolTipText = "This displays any logging with the server."
                .pViewDrives(cnt).ToolTipText = "Connects you to your local drives."
                .setLocation(cnt).ToolTipText = "Allows you to eaisly type in any URL."
                .pView(cnt).ToolTipText = "This displays files and folders."
                .pDummyView(cnt).ToolTipText = "This displays any logging with the server."
                
                .UserControls(cnt).Buttons(1).ToolTipText = "Up One Level. Navigates to the parent directory."
                .UserControls(cnt).Buttons(2).ToolTipText = "Back. Navigates back to the last directory."
                .UserControls(cnt).Buttons(3).ToolTipText = "Forward. Navigates forward to a directory."
                .UserControls(cnt).Buttons(5).ToolTipText = "Stop/Disconnect.  Stops any file action."
                .UserControls(cnt).Buttons(6).ToolTipText = "Refresh.  Refreshes your directory view."
                .UserControls(cnt).Buttons(8).ToolTipText = "New Folder. Creates a new folder."
                .UserControls(cnt).Buttons(9).ToolTipText = "Delete. Deletes selected item(s)."
                .UserControls(cnt).Buttons(11).ToolTipText = "Cut. Cuts selected items to clipboard."
                .UserControls(cnt).Buttons(12).ToolTipText = "Copy. Copies selected items to clipboard."
                .UserControls(cnt).Buttons(13).ToolTipText = "Paste. Pastes the clipboard contents."
            
                .UserGo(cnt).Buttons(1).ToolTipText = "Connects you to a resource or URL."
                .UserGo(cnt).Buttons(2).ToolTipText = "Disconnect the current connection."
                .UserGo(cnt).Buttons(3).ToolTipText = "Opens a dialog to easily access your directories."
                
                .pStatus(cnt).ToolTipText = "Displays the transfer status."
                .pProgress(cnt).ToolTipText = "Displays the transfer status."
            
            Next
            
            .ListView1.ToolTipText = ""
        Else
            .mnuShowToolTips.Checked = False
            
            For cnt = 0 To -dbSettings.GetClientSetting("ViewDoubleWindow")
                .Picture1(cnt).ToolTipText = ""
                .Image1(cnt).ToolTipText = ""
                .pViewDrives(cnt).ToolTipText = ""
                .setLocation(cnt).ToolTipText = ""
                .pView(cnt).ToolTipText = ""
                .pDummyView(cnt).ToolTipText = ""
                
                .UserControls(cnt).Buttons(1).ToolTipText = ""
                .UserControls(cnt).Buttons(2).ToolTipText = ""
                .UserControls(cnt).Buttons(3).ToolTipText = ""
                .UserControls(cnt).Buttons(5).ToolTipText = ""
                .UserControls(cnt).Buttons(6).ToolTipText = ""
                .UserControls(cnt).Buttons(8).ToolTipText = ""
                .UserControls(cnt).Buttons(9).ToolTipText = ""
                .UserControls(cnt).Buttons(11).ToolTipText = ""
                .UserControls(cnt).Buttons(12).ToolTipText = ""
                .UserControls(cnt).Buttons(13).ToolTipText = ""
            
                .UserGo(cnt).Buttons(1).ToolTipText = ""
                .UserGo(cnt).Buttons(2).ToolTipText = ""
                .UserGo(cnt).Buttons(3).ToolTipText = ""
                
                .pStatus(cnt).ToolTipText = ""
                .pProgress(cnt).ToolTipText = ""
            
            Next
            
            .ListView1.ToolTipText = ""
            
        End If
    End With
    
End Sub

Public Function LoadGUI(ByVal Index As Integer)
    With Me
   
        .NoResize = True
        
        .Image1(Index).Stretch = False
        SetPicture "list_background_graphic", .Image1(Index)
        
        .pView(Index).SortKey = dbSettings.GetClientSetting("wColumnKey" & Trim(Index))
        .pView(Index).SortOrder = dbSettings.GetClientSetting("wColumnSort" & Trim(Index))
        
        .vSizer.BackColor = GetSkinColor("sizers_normal")
        .hSizer.BackColor = GetSkinColor("sizers_normal")
    
        .pDummyView(Index).BackColor = GetSkinColor("logview_backcolor")
        
        .pView(Index).BackColor = GetSkinColor("list_backcolor")
        .pView(Index).ForeColor = GetSkinColor("list_textcolor")
        
        .Picture1(Index).BackColor = GetSkinColor("list_backcolor")
        
        .setLocation(Index).BackColor = GetSkinColor("address_backcolor")
        .setLocation(Index).ForeColor = GetSkinColor("address_textcolor")
        
        .pViewDrives(Index).BackColor = GetSkinColor("drivelist_backcolor")
        .pViewDrives(Index).ForeColor = GetSkinColor("drivelist_textcolor")
        
        .ListView1.BackColor = GetSkinColor("transferlist_backcolor")
        .ListView1.ForeColor = GetSkinColor("transferlist_textcolor")
       
        SetIcecue .Line1(0), "icecue_shadow"
        SetIcecue .Line1(2), "icecue_shadow"
        SetIcecue .Line5(0), "icecue_shadow"
        SetIcecue .Line4(0), "icecue_shadow"
        SetIcecue .Line7(0), "icecue_shadow"
        SetIcecue .Line1(1), "icecue_shadow"
        SetIcecue .Line5(1), "icecue_shadow"
        SetIcecue .Line4(1), "icecue_shadow"
        SetIcecue .Line7(1), "icecue_shadow"
        SetIcecue .Line7(2), "icecue_shadow"
        SetIcecue .Line4(2), "icecue_shadow"
        SetIcecue .Line5(2), "icecue_shadow"
        SetIcecue .Line5(3), "icecue_shadow"
        SetIcecue .Line1(3), "icecue_shadow"
        SetIcecue .Line4(3), "icecue_shadow"
        SetIcecue .Line7(3), "icecue_shadow"
        SetIcecue .Line1(4), "icecue_shadow"
        SetIcecue .Line7(4), "icecue_shadow"
        SetIcecue .Line4(4), "icecue_shadow"
        SetIcecue .Line5(4), "icecue_shadow"
        SetIcecue .Line1(5), "icecue_shadow"
        SetIcecue .Line5(5), "icecue_shadow"
        SetIcecue .Line4(5), "icecue_shadow"
        SetIcecue .Line7(5), "icecue_shadow"
                
        SetIcecue .Line6(0), "icecue_hilite"
        SetIcecue .Line2(0), "icecue_hilite"
        SetIcecue .Line3(0), "icecue_hilite"
        SetIcecue .Line8(0), "icecue_hilite"
        SetIcecue .Line2(0), "icecue_hilite"
        SetIcecue .Line2(1), "icecue_hilite"
        SetIcecue .Line6(1), "icecue_hilite"
        SetIcecue .Line3(1), "icecue_hilite"
        SetIcecue .Line8(1), "icecue_hilite"
        SetIcecue .Line2(2), "icecue_hilite"
        SetIcecue .Line8(2), "icecue_hilite"
        SetIcecue .Line6(5), "icecue_hilite"
        SetIcecue .Line3(2), "icecue_hilite"
        SetIcecue .Line6(2), "icecue_hilite"
        SetIcecue .Line2(3), "icecue_hilite"
        SetIcecue .Line6(3), "icecue_hilite"
        SetIcecue .Line3(3), "icecue_hilite"
        SetIcecue .Line8(3), "icecue_hilite"
        SetIcecue .Line2(4), "icecue_hilite"
        SetIcecue .Line3(4), "icecue_hilite"
        SetIcecue .Line6(4), "icecue_hilite"
        SetIcecue .Line8(4), "icecue_hilite"
        SetIcecue .Line2(5), "icecue_hilite"
        SetIcecue .Line3(5), "icecue_hilite"
        SetIcecue .Line8(5), "icecue_hilite"
        SetIcecue .Line6(5), "icecue_hilite"
        
        Set .ListView1.SmallIcons = frmMain.imgOperations(0)

        .UserControls(Index).Buttons.Clear
        Set .UserControls(Index).ImageList = frmMain.imgClient(0)
        Set .UserControls(Index).DisabledImageList = frmMain.imgClient(0)
        Set .UserControls(Index).HotImageList = frmMain.imgClient(1)
        
        .UserGo(Index).Buttons.Clear
        Set .UserGo(Index).ImageList = frmMain.imgClient16(0)
        Set .UserGo(Index).DisabledImageList = frmMain.imgClient16(0)
        Set .UserGo(Index).HotImageList = frmMain.imgClient16(1)
        
        Dim btnX As Button
        
        Set btnX = .UserControls(Index).Buttons.Add(1, "uplevel", , 0, "uplevelout")
        Set btnX = .UserControls(Index).Buttons.Add(2, "back", , 0, "backout")
        Set btnX = .UserControls(Index).Buttons.Add(3, "forward", , 0, "forwardout")
        Set btnX = .UserControls(Index).Buttons.Add(4, , , 3)
        If GetCollectSkinValue("toolbarbutton_spacer") = "none" Then btnX.Visible = False
        Set btnX = .UserControls(Index).Buttons.Add(5, "stop", , 0, "stopout")
        Set btnX = .UserControls(Index).Buttons.Add(6, "refresh", , 0, "refreshout")
        Set btnX = .UserControls(Index).Buttons.Add(7, , , 3)
        If GetCollectSkinValue("toolbarbutton_spacer") = "none" Then btnX.Visible = False
        Set btnX = .UserControls(Index).Buttons.Add(8, "newfolder", , 0, "newfolderout")
        Set btnX = .UserControls(Index).Buttons.Add(9, "delete", , 0, "deleteout")
        Set btnX = .UserControls(Index).Buttons.Add(10, , , 3)
        If GetCollectSkinValue("toolbarbutton_spacer") = "none" Then btnX.Visible = False
        Set btnX = .UserControls(Index).Buttons.Add(11, "cut", , 0, "cutout")
        Set btnX = .UserControls(Index).Buttons.Add(12, "copy", , 0, "copyout")
        Set btnX = .UserControls(Index).Buttons.Add(13, "paste", , 0, "pasteout")
        
        Set btnX = .UserGo(Index).Buttons.Add(1, "go", , 0, "goout")
        Set btnX = .UserGo(Index).Buttons.Add(2, "close", , 0, "closeout")
        Set btnX = .UserGo(Index).Buttons.Add(3, "browse", , 0, "browseout")
        
        .SetClientGUIState Index, st_StartUp
        
        .ClientGUILoaded(Index) = True
        
        .NoResize = False
        
        EnableTransferInfo dbSettings.GetClientSetting("MultiThread")

    End With
End Function
Public Function UnloadGUI(ByVal Index As Integer)
    With Me

        .ClearFileView Index
        
        .UserControls(Index).Buttons.Clear
        Set .UserControls(Index).ImageList = Nothing
        Set .UserControls(Index).DisabledImageList = Nothing
        Set .UserControls(Index).HotImageList = Nothing
        
        .UserGo(Index).Buttons.Clear
        Set .UserGo(Index).ImageList = Nothing
        Set .UserGo(Index).DisabledImageList = Nothing
        Set .UserGo(Index).HotImageList = Nothing
        
        
        Set .ListView1.SmallIcons = Nothing
        
        .ClientGUILoaded(Index) = False

    End With
    
End Function
Public Sub PaintClientGUI(ByRef myForm, ByVal Index As Integer)
    With myForm

        If .NoResize Then Exit Sub
        
        Dim xPixel As Integer
        Dim yPixel As Integer
        xPixel = (Screen.TwipsPerPixelX)
        yPixel = (Screen.TwipsPerPixelY)
                
        Dim cHeight As Integer
        Dim cTop As Integer
        Dim cWidth As Integer
        
        If .dContainer(GetContainerIndex(0, Index)).Visible Then
            .dContainer(GetContainerIndex(0, Index)).Move 0, 0, .userGUI(Index).Width, ((GetSkinDimension("toolbarbutton_height") + 6) * Screen.TwipsPerPixelY) + (yPixel * 4)
            .UserControls(Index).Move yPixel * 2, yPixel * 2, .dContainer(GetContainerIndex(0, Index)).Width - (yPixel * 4)
        End If
            
        If .dContainer(GetContainerIndex(1, Index)).Visible Then
            If .dContainer(GetContainerIndex(0, Index)).Visible Then
                .dContainer(GetContainerIndex(1, Index)).Top = .dContainer(GetContainerIndex(0, Index)).Height + (yPixel)
            Else
                .dContainer(GetContainerIndex(1, Index)).Top = 0
            End If
            .dContainer(GetContainerIndex(1, Index)).Left = 0
            .dContainer(GetContainerIndex(1, Index)).Width = .userGUI(Index).Width
            .dContainer(GetContainerIndex(1, Index)).Height = (22 * Screen.TwipsPerPixelY) + (yPixel * 4)
            .pViewDrives(Index).Move yPixel * 2, yPixel * 2, .dContainer(GetContainerIndex(1, Index)).Width - (yPixel * 4)
        End If
        
        If .dContainer(GetContainerIndex(2, Index)).Visible Then
            If .dContainer(GetContainerIndex(1, Index)).Visible Then
                .dContainer(GetContainerIndex(2, Index)).Top = .dContainer(GetContainerIndex(1, Index)).Top + .dContainer(GetContainerIndex(1, Index)).Height + (yPixel)
            Else
                If .dContainer(GetContainerIndex(0, Index)).Visible Then
                    .dContainer(GetContainerIndex(2, Index)).Top = .dContainer(GetContainerIndex(0, Index)).Height + (yPixel)
                Else
                    .dContainer(GetContainerIndex(2, Index)).Top = 0
                End If
            End If
            .dContainer(GetContainerIndex(2, Index)).Left = 0
            .dContainer(GetContainerIndex(2, Index)).Width = .userGUI(Index).Width
            .dContainer(GetContainerIndex(2, Index)).Height = (22 * Screen.TwipsPerPixelY) + (yPixel * 4)
            .pAddressBar(Index).Move yPixel * 2, yPixel * 2, .dContainer(GetContainerIndex(2, Index)).Width - (yPixel * 4)
            .UserGo(Index).Width = 47 * xPixel
            .UserGo(Index).Height = 36 * yPixel
            .UserGo(Index).Left = .dContainer(GetContainerIndex(2, Index)).Width - .UserGo(Index).Width - (yPixel * 4)
            .UserGo(Index).Top = 0
        End If
        
        cTop = yPixel * 2
        cHeight = .userGUI(Index).Height '- 20
        cWidth = .userGUI(Index).Width
        
        If .dContainer(GetContainerIndex(0, Index)).Visible Then
            cTop = cTop + .dContainer(GetContainerIndex(0, Index)).Height
            cHeight = cHeight - .dContainer(GetContainerIndex(0, Index)).Height
        End If
        If .dContainer(GetContainerIndex(1, Index)).Visible Then
            cTop = cTop + .dContainer(GetContainerIndex(1, Index)).Height
            cHeight = cHeight - .dContainer(GetContainerIndex(1, Index)).Height
        End If
        If .dContainer(GetContainerIndex(2, Index)).Visible Then
            cTop = cTop + .dContainer(GetContainerIndex(2, Index)).Height
            cHeight = cHeight - .dContainer(GetContainerIndex(2, Index)).Height
        End If
            
        If .mnuShowStatusBar.Checked Then
            .pStatus(Index).Visible = True
            .pProgress(Index).Visible = False
            cHeight = cHeight - (26 * Screen.TwipsPerPixelY)
            
            .pStatus(Index).Move 0, (.userGUI(Index).Height - .pStatus(Index).Height), .userGUI(Index).Width, (20 * Screen.TwipsPerPixelY)
            .pProgress(Index).Move 0, (.userGUI(Index).Height - .pProgress(Index).Height), .userGUI(Index).Width, (20 * Screen.TwipsPerPixelY)
            
        Else
            cHeight = cHeight - (2 * Screen.TwipsPerPixelY)
            .pStatus(Index).Visible = False
            .pProgress(Index).Visible = False
        End If
        
        .pView(Index).Move 0, cTop, cWidth, cHeight
        .Picture1(Index).Move 0, cTop, cWidth, cHeight
        SizeInnerImage myForm, CByte(Index)
        
        .pDummyView(Index).Move 0, 0, .Picture1(Index).ScaleWidth, .Picture1(Index).ScaleHeight
    
    End With
End Sub
Private Sub SizeInnerImage(ByRef myForm, ByRef Index As Byte)
    With myForm
        Select Case LCase(GetCollectSkinValue("list_background_resize"))
            Case "auto"
                .Image1(Index).Stretch = True
                If .Picture1(Index).ScaleWidth > 0 Then
                    .Image1(Index).Left = (.Picture1(Index).ScaleWidth / 2) - (.Image1(Index).Width / 2)
                End If
                If .Picture1(Index).ScaleHeight > 0 Then
                    .Image1(Index).Top = (.Picture1(Index).ScaleHeight / 2) - (.Image1(Index).Height / 2)
                End If
            Case "fit"
                .Image1(Index).Stretch = True
                
                .Image1(Index).Left = 0
                If .Picture1(Index).ScaleWidth > 0 Then
                    .Image1(Index).Width = .Picture1(Index).ScaleWidth
                End If
                .Image1(Index).Top = 0
                If .Picture1(Index).ScaleHeight > 0 Then
                    .Image1(Index).Height = .Picture1(Index).ScaleHeight
                End If
            
        End Select
    End With
End Sub
Public Sub PaintInfoGUI(ByRef myForm)
    With myForm
        
        If .NoResize Then Exit Sub
        
        .ListView1.Move 0, 0, .userInfo.Width, .userInfo.Height
    
    End With

End Sub

Public Function GetContainerIndex(ByVal ConNum As Integer, GUIIndex As Integer) As Integer

    Dim retVal As Integer
    Select Case ConNum
        Case 0
        Select Case GUIIndex
            Case 0
                retVal = 0
            Case 1
                retVal = 3
        End Select
        Case 1
        Select Case GUIIndex
            Case 0
                retVal = 1
            Case 1
                retVal = 4
        End Select
        Case 2
        Select Case GUIIndex
            Case 0
                retVal = 2
            Case 1
                retVal = 5
        End Select
    End Select
    GetContainerIndex = retVal
    
End Function

Public Sub dContainerResize(ByRef myForm, Index As Integer)
    With myForm
        
        If .NoResize Then Exit Sub
        
        Dim xPixel As Integer
        Dim yPixel As Integer
        xPixel = (Screen.TwipsPerPixelX)
        yPixel = (Screen.TwipsPerPixelY)
        
        .Line1(Index).X1 = 0
        .Line1(Index).X2 = 0
        .Line1(Index).Y1 = yPixel
        .Line1(Index).Y2 = (.dContainer(Index).Height - (2 * yPixel))
        
        .Line2(Index).X1 = xPixel
        .Line2(Index).X2 = xPixel
        .Line2(Index).Y1 = yPixel
        .Line2(Index).Y2 = (.dContainer(Index).Height - (2 * yPixel))
        
        .Line3(Index).X1 = .dContainer(Index).Width - yPixel
        .Line3(Index).X2 = .dContainer(Index).Width - yPixel
        .Line3(Index).Y1 = 0
        .Line3(Index).Y2 = .dContainer(Index).Height - yPixel
        
        .Line4(Index).X1 = .dContainer(Index).Width - (xPixel * 2)
        .Line4(Index).X2 = .dContainer(Index).Width - (xPixel * 2)
        .Line4(Index).Y1 = yPixel
        .Line4(Index).Y2 = (.dContainer(Index).Height - (2 * yPixel))
        
        .Line5(Index).X1 = 0
        .Line5(Index).X2 = .dContainer(Index).Width
        .Line5(Index).Y1 = 0
        .Line5(Index).Y2 = 0
        
        .Line6(Index).X1 = xPixel
        .Line6(Index).X2 = .dContainer(Index).Width
        .Line6(Index).Y1 = yPixel
        .Line6(Index).Y2 = yPixel
        
        .Line7(Index).X1 = xPixel
        .Line7(Index).X2 = .dContainer(Index).Width - (xPixel * 2)
        .Line7(Index).Y1 = .dContainer(Index).Height - (yPixel * 2)
        .Line7(Index).Y2 = .dContainer(Index).Height - (yPixel * 2)
        
        .Line8(Index).X1 = 0
        .Line8(Index).X2 = .dContainer(Index).Width
        .Line8(Index).Y1 = .dContainer(Index).Height - yPixel
        .Line8(Index).Y2 = .dContainer(Index).Height - yPixel

    End With
End Sub

Public Sub FormResize(ByRef myForm)
    With myForm
        On Error Resume Next
        
        If .NoResize Then Exit Sub
        
        If .mnuDoubleWin.Checked Then
            If Not .userGUI(0).Visible Then .userGUI(0).Visible = True
            If Not .userGUI(1).Visible Then .userGUI(1).Visible = True
        
            .vSizer.Visible = True
            If .mnuMultiThread.Checked Then
                If Not .userInfo.Visible = True Then .userInfo.Visible = True
                .hSizer.Visible = True
            
                If .ReCenterSizers Then
                    .hSizer.Top = (0.68 * .ScaleHeight)
                End If
            
                .hSizer.Left = (Border * Screen.TwipsPerPixelX)
                .hSizer.Width = .ScaleWidth - ((Border * Screen.TwipsPerPixelX) * 2)
                If .hSizer.Top < (Border * Screen.TwipsPerPixelY) * 2 Then .hSizer.Top = (Border * Screen.TwipsPerPixelY) * 2
                If .hSizer.Top + ((Border * Screen.TwipsPerPixelY) * 2) > .ScaleHeight Then .hSizer.Top = .ScaleHeight - ((Border * Screen.TwipsPerPixelY) * 2)
            
            Else
                If Not .userInfo.Visible = False Then .userInfo.Visible = False
                .hSizer.Visible = False
            End If
            
            If .ReCenterSizers Then
                .vSizer.Left = (.ScaleWidth / 2) - ((Border * Screen.TwipsPerPixelX) / 2)
            End If
            
            If .vSizer.Left < (Border * Screen.TwipsPerPixelX) * 2 Then .vSizer.Left = (Border * Screen.TwipsPerPixelX) * 2
            If .vSizer.Left + ((Border * Screen.TwipsPerPixelX) * 2) > .ScaleWidth Then .vSizer.Left = .ScaleWidth - ((Border * Screen.TwipsPerPixelX) * 2)
            
            .vSizer.Top = (Border * Screen.TwipsPerPixelY)
            .userInfo.Left = (Border * Screen.TwipsPerPixelX)
            .userInfo.Width = (.ScaleWidth) - ((Border * Screen.TwipsPerPixelX) * 2)
                        
            .userGUI(0).Move (Border * Screen.TwipsPerPixelX), 0, (.vSizer.Left) - ((Border * Screen.TwipsPerPixelX)) - Screen.TwipsPerPixelX
            
            .userGUI(1).Move (.vSizer.Left + .vSizer.Width) + Screen.TwipsPerPixelX, 0, (.ScaleWidth) - (.vSizer.Left + .vSizer.Width) - (Border * Screen.TwipsPerPixelX)
            
            If .mnuMultiThread.Checked Then
                .userInfo.Top = (.hSizer.Top + .hSizer.Height + Screen.TwipsPerPixelY)
                .userGUI(0).Height = (.hSizer.Top)
                .userGUI(1).Height = (.hSizer.Top)
                .userInfo.Height = .ScaleHeight - .userInfo.Top - (Border * Screen.TwipsPerPixelY)
                .vSizer.Height = .userInfo.Top - (Border * Screen.TwipsPerPixelY) - Screen.TwipsPerPixelY
                PaintInfoGUI Me
            Else
                .userGUI(0).Height = (.ScaleHeight) - (Border * Screen.TwipsPerPixelY) - Screen.TwipsPerPixelY
                .userGUI(1).Height = (.ScaleHeight) - (Border * Screen.TwipsPerPixelY) - Screen.TwipsPerPixelY
                .vSizer.Height = (.ScaleHeight) - (Border * Screen.TwipsPerPixelY)
            End If
            
            PaintClientGUI Me, 0
            PaintClientGUI Me, 1
        
        Else
            If Not .userGUI(0).Visible Then .userGUI(0).Visible = True
            If .userGUI(1).Visible Then .userGUI(1).Visible = False
            
            .vSizer.Visible = False
            If .mnuMultiThread.Checked Then
                If Not .userInfo.Visible = True Then .userInfo.Visible = True
                .hSizer.Visible = True
                
                If .hSizer.Top < (Border * Screen.TwipsPerPixelY) * 2 Then .hSizer.Top = (Border * Screen.TwipsPerPixelY) * 2
                If .hSizer.Top + ((Border * Screen.TwipsPerPixelY) * 2) > .ScaleHeight Then .hSizer.Top = .ScaleHeight - ((Border * Screen.TwipsPerPixelY) * 2)
                .hSizer.Move (Border * Screen.TwipsPerPixelX), , .ScaleWidth - ((Border * Screen.TwipsPerPixelX) * 2)
                
                If .ReCenterSizers Then
                    .hSizer.Top = (0.68 * .ScaleHeight)
                End If
                
                .userInfo.Move (Border * Screen.TwipsPerPixelX), , (.ScaleWidth) - ((Border * Screen.TwipsPerPixelX) * 2)
            
            Else
                If Not .userInfo.Visible = False Then .userInfo.Visible = False
                .hSizer.Visible = False
            End If
           
            .userGUI(0).Move (Border * Screen.TwipsPerPixelX), (Border * Screen.TwipsPerPixelY), .ScaleWidth - ((Border * Screen.TwipsPerPixelX) * 2)
        
            If .mnuMultiThread.Checked Then
                .userInfo.Top = (.hSizer.Top + .hSizer.Height + Screen.TwipsPerPixelY)
                .userGUI(0).Height = (.hSizer.Top) - ((Border * Screen.TwipsPerPixelY)) - Screen.TwipsPerPixelY
                .userInfo.Height = .ScaleHeight - .userInfo.Top - (Border * Screen.TwipsPerPixelY)
                PaintInfoGUI Me
            Else
                .userGUI(0).Height = (.ScaleHeight) - (Border * Screen.TwipsPerPixelY) - (Screen.TwipsPerPixelY * 2)
            End If
        
            PaintClientGUI Me, 0
        End If

        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        
    End With
End Sub

Public Sub ViewDoubleWindow(ByRef myForm, ByVal IsDouble As Boolean, Optional ByVal SetValue As Boolean = True)
    With Me
    
        .mnuDoubleWin.Checked = IsDouble
        
        If IsDouble Then
            LoadGUI 1
        Else
            UnloadGUI 1
        End If
        
        FormResize Me
    
    End With
End Sub
Public Sub ViewTransferInfo(ByRef myForm, ByVal IsVisible As Boolean, Optional SetValue As Boolean = True)
    With Me

        .mnuMultiThread.Checked = IsVisible
            
         .ReCenterSizers = True
        FormResize Me
        .ReCenterSizers = False
        FormResize Me
    
    End With
End Sub

Public Sub EnableTransferInfo(ByVal enabled As Boolean)
    With Me
        
        .mnuMultiThread.Checked = enabled
        
        FormResize Me

    End With
End Sub

Public Sub ViewStatusBar(ByRef myForm, ByVal IsVisible As Boolean, Optional SetValue As Boolean = True)
    With Me
    
        .mnuShowStatusBar.Checked = IsVisible

        PaintClientGUI Me, 0
        If .mnuDoubleWin.Checked Then PaintClientGUI Me, 1

    End With
End Sub
Public Sub ViewToolBar(ByRef myForm, ByVal IsVisible As Boolean, Optional SetValue As Boolean = True)
    With Me
    
        .mnuShowToolBar.Checked = IsVisible

        .dContainer(GetContainerIndex(0, 1)).Visible = IsVisible
        .dContainer(GetContainerIndex(0, 0)).Visible = IsVisible
        
        PaintClientGUI Me, 0
        If .mnuDoubleWin.Checked Then PaintClientGUI Me, 1
        
    End With
End Sub
Public Sub ViewDriveList(ByRef myForm, ByVal IsVisible As Boolean, Optional SetValue As Boolean = True)
    With Me
        
        .mnuShowDriveList.Checked = IsVisible

        .dContainer(GetContainerIndex(1, 1)).Visible = IsVisible
        .dContainer(GetContainerIndex(1, 0)).Visible = IsVisible
        PaintClientGUI Me, 0
        If .mnuDoubleWin.Checked Then PaintClientGUI Me, 1
    
    End With
End Sub
Public Sub ViewAddressBar(ByRef myForm, ByVal IsVisible As Boolean, Optional SetValue As Boolean = True)
    With Me
    
        .mnuShowAddressBar.Checked = IsVisible

        .dContainer(GetContainerIndex(2, 1)).Visible = IsVisible
        .dContainer(GetContainerIndex(2, 0)).Visible = IsVisible
        PaintClientGUI Me, 0
        If .mnuDoubleWin.Checked Then PaintClientGUI Me, 1

    End With
End Sub

Public Sub EnableFileMenu(ByVal Index As Integer)
    
    Me.FocusIndex = Index
    
    Dim selCount As Integer
    Dim IsConnected As Boolean
    
    Dim myClient As NTAdvFTP61.Client
    Set myClient = Me.SetMyClient(Index)
    IsConnected = Not myClient.ConnectedState() = False
    
    selCount = GetSelectedCount(Me.pView(Index))
    
    Me.mnuOpen.enabled = (IsConnected And Me.GetState(Index) = st_ProcessSuccess) And (selCount = 1)
    'Me.mnuCache.enabled = (IsConnected And Me.GetState(Index) = st_ProcessSuccess) And (selCount > 0)
    Me.mnuCut.enabled = (IsConnected And Me.GetState(Index) = st_ProcessSuccess) And (selCount > 0)
    Me.mnuCopy.enabled = (IsConnected And Me.GetState(Index) = st_ProcessSuccess) And (selCount > 0)
    Me.mnuPaste.enabled = (IsConnected And Me.GetState(Index) = st_ProcessSuccess)
    Me.mnuNewFolder.enabled = (IsConnected And Me.GetState(Index) = st_ProcessSuccess)
    Me.mnuDelete.enabled = (IsConnected And Me.GetState(Index) = st_ProcessSuccess) And (selCount > 0)
    Me.mnuRename.enabled = (IsConnected And Me.GetState(Index) = st_ProcessSuccess) And (selCount = 1)

    Me.mnuSelAll.enabled = (IsConnected And Me.GetState(Index) = st_ProcessSuccess)
    Me.mnuSelFiles.enabled = (IsConnected And Me.GetState(Index) = st_ProcessSuccess)
    Me.mnuSelFolders.enabled = (IsConnected And Me.GetState(Index) = st_ProcessSuccess)
    Me.mnuWildCard.enabled = (IsConnected And Me.GetState(Index) = st_ProcessSuccess)

    Me.mnuRefresh.enabled = IsConnected
    Me.mnuStop.enabled = (IsConnected And Me.GetState(Index) = st_Processing)

    Me.mnuConnect.enabled = Not IsConnected
    Me.mnuDisconnect = IsConnected

    Me.mnuInfoBanner.enabled = (IsConnected And Me.GetState(Index) = st_ProcessSuccess) And (myClient.URLType = URLTypes.ftp)
    Me.mnuInfoHelp.enabled = (IsConnected And Me.GetState(Index) = st_ProcessSuccess) And (myClient.URLType = URLTypes.ftp)
    Me.mnuInfoMOTD.enabled = (IsConnected And Me.GetState(Index) = st_ProcessSuccess) And (myClient.URLType = URLTypes.ftp)
    Me.mnuInfoStat.enabled = (IsConnected And Me.GetState(Index) = st_ProcessSuccess) And (myClient.URLType = URLTypes.ftp)
    Me.mnuInfoSystem.enabled = (IsConnected And Me.GetState(Index) = st_ProcessSuccess) And (myClient.URLType = URLTypes.ftp)

    Set myClient = Nothing

End Sub




