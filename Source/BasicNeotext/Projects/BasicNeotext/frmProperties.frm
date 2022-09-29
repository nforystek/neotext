VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Project Properties"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   Icon            =   "frmProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4050
      Index           =   0
      Left            =   210
      TabIndex        =   4
      Top             =   540
      Width           =   5955
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3765
         TabIndex        =   29
         Top             =   1530
         Width           =   315
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "frmProperties.frx":000C
         Left            =   3195
         List            =   "frmProperties.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   210
         Width           =   2550
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmProperties.frx":002C
         Left            =   165
         List            =   "frmProperties.frx":003C
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   210
         Width           =   2805
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   4440
         TabIndex        =   26
         Top             =   1530
         Width           =   1305
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   195
         TabIndex        =   25
         Top             =   2235
         Width           =   5550
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   195
         TabIndex        =   24
         Top             =   1530
         Width           =   3570
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   180
         TabIndex        =   23
         Top             =   840
         Width           =   3900
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Retained In Memor&y"
         Height          =   225
         Left            =   210
         TabIndex        =   22
         Top             =   3735
         Width           =   2115
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Require &License Key"
         Height          =   195
         Left            =   210
         TabIndex        =   21
         Top             =   3420
         Width           =   1995
      End
      Begin VB.CheckBox Check2 
         Caption         =   "&Upgrade ActiveX Controls"
         Height          =   195
         Left            =   210
         TabIndex        =   20
         Top             =   3090
         Width           =   2145
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Unatt&ended Execution"
         Height          =   225
         Left            =   210
         TabIndex        =   19
         Top             =   2760
         Width           =   2070
      End
      Begin VB.Frame Frame2 
         Caption         =   "Threading &Model"
         Height          =   1305
         Left            =   2475
         TabIndex        =   15
         Top             =   2610
         Width           =   3300
         Begin VB.VScrollBar VScroll1 
            Height          =   240
            Left            =   2145
            TabIndex        =   52
            Top             =   900
            Width           =   225
         End
         Begin VB.TextBox Text12 
            Height          =   315
            Left            =   1560
            TabIndex        =   51
            Text            =   "1"
            Top             =   870
            Width           =   840
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Thread P&ool                         threads"
            Height          =   195
            Left            =   270
            TabIndex        =   18
            Top             =   900
            Width           =   2850
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Th&read per Object"
            Height          =   195
            Left            =   255
            TabIndex        =   17
            Top             =   630
            Width           =   2100
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmProperties.frx":0079
            Left            =   225
            List            =   "frmProperties.frx":0083
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   225
            Width           =   2835
         End
      End
      Begin VB.Label Label6 
         Caption         =   "&Project Description:"
         Height          =   225
         Left            =   195
         TabIndex        =   14
         Top             =   2010
         Width           =   2205
      End
      Begin VB.Label Label5 
         Caption         =   "Project Help Context &ID:"
         Height          =   525
         Left            =   4455
         TabIndex        =   13
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "&Help File Name:"
         Height          =   255
         Left            =   195
         TabIndex        =   12
         Top             =   1305
         Width           =   1995
      End
      Begin VB.Label Label3 
         Caption         =   "Project &Name:"
         Height          =   195
         Left            =   195
         TabIndex        =   11
         Top             =   615
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "&Startup Object:"
         Height          =   255
         Left            =   3225
         TabIndex        =   10
         Top             =   -15
         Width           =   1710
      End
      Begin VB.Label Label1 
         Caption         =   "Project &Type:"
         Height          =   210
         Left            =   180
         TabIndex        =   9
         Top             =   -15
         Width           =   1425
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4050
      Index           =   1
      Left            =   210
      TabIndex        =   5
      Top             =   480
      Width           =   5955
      Begin VB.CheckBox Check6 
         Caption         =   "R&emove information about unused ActiveX Controls"
         Height          =   240
         Left            =   135
         TabIndex        =   50
         Top             =   3795
         Width           =   4320
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   2640
         TabIndex        =   49
         Top             =   3420
         Width           =   3150
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   2640
         TabIndex        =   47
         Top             =   3015
         Width           =   3150
      End
      Begin VB.Frame Frame3 
         Caption         =   "Application"
         Height          =   1440
         Left            =   2550
         TabIndex        =   31
         Top             =   30
         Width           =   3165
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   630
            TabIndex        =   40
            Top             =   420
            Width           =   2370
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Left            =   630
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   855
            Width           =   1515
         End
         Begin VB.Image Image1 
            Height          =   420
            Left            =   2385
            Top             =   840
            Width           =   480
         End
         Begin VB.Label Label9 
            Caption         =   "Ic&on:"
            Height          =   240
            Left            =   135
            TabIndex        =   39
            Top             =   900
            Width           =   495
         End
         Begin VB.Label Label8 
            Caption         =   "&Title:"
            Height          =   240
            Left            =   135
            TabIndex        =   38
            Top             =   450
            Width           =   435
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Version Number"
         Height          =   1455
         Left            =   165
         TabIndex        =   30
         Top             =   30
         Width           =   2250
         Begin VB.CheckBox Check5 
            Caption         =   "A&uto Increment"
            Height          =   195
            Left            =   150
            TabIndex        =   37
            Top             =   1095
            Width           =   1755
         End
         Begin VB.TextBox Text7 
            Height          =   300
            Left            =   1440
            TabIndex        =   36
            Text            =   "0"
            Top             =   660
            Width           =   525
         End
         Begin VB.TextBox Text6 
            Height          =   300
            Left            =   780
            TabIndex        =   35
            Text            =   "0"
            Top             =   660
            Width           =   525
         End
         Begin VB.TextBox Text5 
            Height          =   300
            Left            =   135
            TabIndex        =   34
            Text            =   "1"
            Top             =   660
            Width           =   525
         End
         Begin VB.Label Label16 
            Caption         =   "&Major:"
            Height          =   285
            Left            =   135
            TabIndex        =   65
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "M&inor:"
            Height          =   285
            Left            =   765
            TabIndex        =   64
            Top             =   375
            Width           =   525
         End
         Begin VB.Label Label7 
            Caption         =   "&Revision:"
            Height          =   285
            Left            =   1425
            TabIndex        =   33
            Top             =   375
            Width           =   720
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Version Information"
         Height          =   1335
         Left            =   165
         TabIndex        =   32
         Top             =   1560
         Width           =   5565
         Begin VB.ListBox List2 
            Height          =   450
            ItemData        =   "frmProperties.frx":00AC
            Left            =   195
            List            =   "frmProperties.frx":00C2
            TabIndex        =   44
            Top             =   510
            Width           =   1695
         End
         Begin VB.TextBox Text9 
            Height          =   630
            Left            =   2010
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   45
            Top             =   510
            Width           =   3360
         End
         Begin VB.Label Label11 
            Caption         =   "&Value:"
            Height          =   225
            Left            =   2055
            TabIndex        =   43
            Top             =   270
            Width           =   1380
         End
         Begin VB.Label Label10 
            Caption         =   "T&ype:"
            Height          =   255
            Left            =   210
            TabIndex        =   42
            Top             =   270
            Width           =   525
         End
      End
      Begin VB.Label Label13 
         Caption         =   "Con&ditional Compilation Arguments:"
         Height          =   225
         Left            =   105
         TabIndex        =   48
         Top             =   3450
         Width           =   2700
      End
      Begin VB.Label Label12 
         Caption         =   "&Command Line Arguments:"
         Height          =   270
         Left            =   105
         TabIndex        =   46
         Top             =   3075
         Width           =   1965
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4050
      Index           =   2
      Left            =   210
      TabIndex        =   6
      Top             =   495
      Width           =   5955
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   1755
         TabIndex        =   63
         Text            =   "&H11000000"
         Top             =   3270
         Width           =   2010
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Compile to &Native Code"
         Height          =   285
         Left            =   315
         TabIndex        =   55
         Top             =   405
         Width           =   1980
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Compile to &P-Code"
         Height          =   210
         Left            =   330
         TabIndex        =   54
         Top             =   90
         Width           =   1905
      End
      Begin VB.Frame Frame6 
         Height          =   2250
         Left            =   105
         TabIndex        =   53
         Top             =   450
         Width           =   5745
         Begin VB.CommandButton Command5 
            Caption         =   "Ad&vanced Optimizations..."
            Height          =   510
            Left            =   525
            TabIndex        =   61
            Top             =   1545
            Width           =   2445
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Create Symbolic &Debug Info"
            Height          =   255
            Left            =   2715
            TabIndex        =   60
            Top             =   750
            Width           =   2370
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Favo&r Pentium Pro(tm)"
            Height          =   225
            Left            =   2715
            TabIndex        =   59
            Top             =   375
            Width           =   2055
         End
         Begin VB.OptionButton Option7 
            Caption         =   "N&o Optimization"
            Height          =   255
            Left            =   405
            TabIndex        =   58
            Top             =   1125
            Width           =   1620
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Optimize for &Small Code"
            Height          =   225
            Left            =   390
            TabIndex        =   57
            Top             =   735
            Width           =   2010
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Optimize for &Fast Code"
            Height          =   225
            Left            =   390
            TabIndex        =   56
            Top             =   360
            Width           =   1995
         End
      End
      Begin VB.Label Label14 
         Caption         =   "DLL &Base Address:"
         Height          =   270
         Left            =   210
         TabIndex        =   62
         Top             =   3315
         Width           =   1485
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4050
      Index           =   3
      Left            =   210
      TabIndex        =   7
      Top             =   495
      Width           =   5955
      Begin VB.Frame Frame9 
         Caption         =   "Version Compatibility"
         Height          =   1710
         Left            =   150
         TabIndex        =   68
         Top             =   2025
         Width           =   5640
         Begin VB.CommandButton Command6 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5220
            TabIndex        =   76
            Top             =   1275
            Width           =   330
         End
         Begin VB.TextBox Text14 
            Height          =   285
            Left            =   240
            TabIndex        =   75
            Top             =   1275
            Width           =   4935
         End
         Begin VB.OptionButton Option12 
            Caption         =   "&Binary Compatibility"
            Height          =   300
            Left            =   210
            TabIndex        =   74
            Top             =   900
            Width           =   2010
         End
         Begin VB.OptionButton Option11 
            Caption         =   "&Project Compatibility"
            Height          =   255
            Left            =   210
            TabIndex        =   73
            Top             =   615
            Width           =   2550
         End
         Begin VB.OptionButton Option10 
            Caption         =   "&No Compatibility"
            Height          =   270
            Left            =   225
            TabIndex        =   72
            Top             =   300
            Width           =   2445
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Remote Server"
         Height          =   765
         Left            =   150
         TabIndex        =   67
         Top             =   1215
         Width           =   2805
         Begin VB.CheckBox Check9 
            Caption         =   "R&emote Server Files"
            Height          =   240
            Left            =   195
            TabIndex        =   71
            Top             =   360
            Width           =   2235
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Start Mode"
         Height          =   1065
         Left            =   135
         TabIndex        =   66
         Top             =   105
         Width           =   2820
         Begin VB.OptionButton Option9 
            Caption         =   "&ActiveX Component"
            Height          =   270
            Left            =   180
            TabIndex        =   70
            Top             =   675
            Width           =   1920
         End
         Begin VB.OptionButton Option8 
            Caption         =   "S&tandalone"
            Height          =   240
            Left            =   165
            TabIndex        =   69
            Top             =   360
            Width           =   1815
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4050
      Index           =   4
      Left            =   195
      TabIndex        =   8
      Top             =   555
      Width           =   5955
      Begin VB.Frame Frame10 
         Caption         =   "When this project starts"
         Height          =   3675
         Left            =   195
         TabIndex        =   77
         Top             =   60
         Width           =   5580
         Begin VB.CheckBox Check10 
            Caption         =   "&Use existing browser"
            Height          =   195
            Left            =   195
            TabIndex        =   86
            Top             =   3300
            Width           =   2655
         End
         Begin VB.TextBox Text16 
            Height          =   285
            Left            =   510
            TabIndex        =   85
            Top             =   2760
            Width           =   4860
         End
         Begin VB.OptionButton Option16 
            Caption         =   "Start &browser with URL:"
            Height          =   210
            Left            =   195
            TabIndex        =   84
            Top             =   2430
            Width           =   2100
         End
         Begin VB.CommandButton Command7 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5010
            TabIndex        =   83
            Top             =   1815
            Width           =   315
         End
         Begin VB.TextBox Text15 
            Height          =   285
            Left            =   525
            TabIndex        =   82
            Top             =   1815
            Width           =   4365
         End
         Begin VB.OptionButton Option15 
            Caption         =   "Start &program:"
            Height          =   255
            Left            =   195
            TabIndex        =   81
            Top             =   1455
            Width           =   1380
         End
         Begin VB.ComboBox Combo5 
            Height          =   315
            Left            =   1905
            TabIndex        =   80
            Text            =   "Combo5"
            Top             =   840
            Width           =   3405
         End
         Begin VB.OptionButton Option14 
            Caption         =   "Start &component:"
            Height          =   255
            Left            =   195
            TabIndex        =   79
            Top             =   885
            Width           =   1635
         End
         Begin VB.OptionButton Option13 
            Caption         =   "&Wait for components to be created"
            Height          =   285
            Left            =   195
            TabIndex        =   78
            Top             =   390
            Width           =   3225
         End
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   4800
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3780
      TabIndex        =   2
      Top             =   4800
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Help"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   4800
      Width           =   1170
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4620
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   8149
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Make"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Compile"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Component"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Debugging"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private myProj As Project

'mk:@MSITStore:C:\Program%20Files\Microsoft%20Visual%20Studio\MSDN98\98VS\1033\vbenlr98.chm::/html/vamsgInvalidConstValSyntax.htm
'mk:@MSITStore:C:\Program%20Files\Microsoft%20Visual%20Studio\MSDN98\98VS\1033\vbenlr98.chm::/html/vamsgNoHelp.htm

'mk:@MSITStore:C:\Program%20Files\Microsoft%20Visual%20Studio\MSDN98\98VS\1033\vb98.chm::/html/vbdlgGeneralTabProjectSettingsDialog.htm
'mk:@MSITStore:C:\Program%20Files\Microsoft%20Visual%20Studio\MSDN98\98VS\1033\vb98.chm::/html/vbdlgBuildOptionsTabProjectSettingsDialog.htm
'mk:@MSITStore:C:\Program%20Files\Microsoft%20Visual%20Studio\MSDN98\98VS\1033\vb98.chm::/html/vbdlgCompileTabProjectProperties.htm
'mk:@MSITStore:C:\Program%20Files\Microsoft%20Visual%20Studio\MSDN98\98VS\1033\vb98.chm::/html/vbdlgStartUpTabProjectProperties.htm
'mk:@MSITStore:C:\Program%20Files\Microsoft%20Visual%20Studio\MSDN98\98VS\1033\vb98.chm::/html/vadlgDebuggingTab(ProjectPropertiesDialogBox).htm

'mk:@MSITStore:C:\Program%20Files\Microsoft%20Visual%20Studio\MSDN98\98VS\1033\vb98.chm::/html/vbdlgAdvancedOptimizationsCompile.htm

Public Sub ShowProperties(ByRef Project As VBProject, Optional Proj As Project = Nothing)
   
    
    If Proj Is Nothing Then Set Proj = Projs

    If LCase(Proj.Location) = LCase(Project.FileName) Then
        Set myProj = Proj
    End If

    If myProj Is Nothing Then
        Dim p As Project
        For Each p In Proj.Includes
            ShowProperties Project, p
        Next
    Else
        ParseProperties myProj, Project.FileName
        
        Me.Show
        
    End If
End Sub

Public Function ParseProperties(ByRef Self As Project, ByVal URI As String) As String

    
    Dim inText As String
    Dim inLine As String
    Dim inPath As String
    Dim inName As String
    
    inText = Self.Contents
    Do Until inText = ""
        inLine = RemoveNextArg(inText, vbCrLf)
        inName = LCase(RemoveNextArg(inLine, "="))
        Select Case inName
'            Case "reference", "object"
'                Do Until inLine = ""
'                    If InStr(NextArg(inLine, "#"), "..\") > 0 Then
'                        inLine = NextArg(inLine, "#")
'                        Exit Do
'                    Else
'                        Select Case GetFilePath(NextArg(inLine, "#"))
'                            Case "dll", "ocx", "exe", "tlb"
'                                inLine = NextArg(inLine, "#")
'                                Exit Do
'                            Case Else
'                                RemoveNextArg inLine, "#"
'                        End Select
'                    End If
'                Loop
'                If inLine <> "" Then
'                    If Mid(inLine, 2, 1) = "\" And Mid(inLine, 4, 2) = ".." Or Mid(inLine, 5, 1) = ":" And (Not Left(inLine, 1) = ".") Then
'                        inLine = Mid(inLine, 3)
'                    End If
'                    If PathExists(MapPaths(, NextArg(inLine, "#"), Self.Location), True) Then
'                        inLine = MapPaths(, NextArg(inLine, "#"), Self.Location)
'                    ElseIf PathExists(MapPaths(, NextArg(inLine, "#")), True) Then
'                        inLine = MapPaths(, NextArg(inLine, "#"))
'                    End If
'                    If InStr(1, ParseProject, inLine, vbTextCompare) = 0 Then
'                        ParseProject = ParseProject & inLine & vbCrLf
'                    End If
'                    inLine = ""
'                End If
'            Case "designer", "module", "class", "userdocument", "form", "relateddoc", "usercontrol"
'                If InStr(inLine, ";") > 0 Then inLine = RemoveArg(inLine, ";")
'                If PathExists(MapPaths(, inLine, Self.Location), True) Then
'                    inLine = MapPaths(, inLine, Self.Location)
'                ElseIf PathExists(MapPaths(, inLine), True) Then
'                    inLine = MapPaths(, inLine)
'                Else
'                    inLine = ""
'                End If
'                If inLine <> "" Then
'                    If InStr(1, ParseProject, inLine, vbTextCompare) = 0 Then
'                        ParseProject = ParseProject & inLine & vbCrLf
'                    End If
'                    inLine = ""
'                End If
'            Case "exename32"
'                Self.Compiled = MapPaths(Self.Location, RemoveQuotedArg(inLine, """", """"), inPath)
'            Case "path32"
'                inPath = Replace(inLine, """", "")
'                Self.Compiled = MapPaths(MapPaths(Self.Location, inPath), GetFileName(Self.Compiled))
            
            
            Case "compatibleexe32"
                Self.Compiled = MapPaths(Self.Location, RemoveQuotedArg(inLine, """", """"), inPath)
            Case "command32"
                Self.CmdLine = RemoveQuotedArg(inLine, """", """")
            Case "resfile32"
            
            Case "condcomp"
                Self.CondComp = RemoveQuotedArg(inLine, """", """")

            Case "type"
            Case "iconform"
            Case "startup"
            Case "helpfile"
            Case "title"
            Case "name"
            Case "helpcontextid"
            Case "description"
            Case "compatiblemode"
            Case "compcond"
            Case "majorver"
            Case "minorver"
            Case "revisionver"
            Case "autoincrementver"
            Case "serversupportfiles"
            Case "versioncomments"
            Case "versioncompanyname"
            Case "versionlegaltrademarks"
            Case "versionfiledescription"
            Case "versionlegalcopyright"
            Case "versionproductname"
            Case "versioncompatible32"
            Case "compilationtype"
            Case "optimizationtype"
            Case "favorpentiumpro(tm)"
            Case "removeunusedcontrolinfo"
            Case "codeviewdebuginfo"
            Case "noaliasing"
            Case "boundscheck"
            Case "overflowcheck"
            Case "flpointcheck"
            Case "fdivcheck"
            Case "unroundedfp"
            Case "startmode"
            Case "unattended"
            Case "threadingmodel"
            Case "retained"
            Case "threadperobject"
            Case "maxnumberofthreads"
            Case "debugstartupoption"
            Case "useexistingbrowser"
            Case "[ms transaction server]"
            Case "autorefresh"

            

        End Select
    Loop

End Function

Private Sub TabStrip1_Click()
    Frame1(0).Visible = (TabStrip1.SelectedItem.Index - 1) = 0
    Frame1(1).Visible = (TabStrip1.SelectedItem.Index - 1) = 1
    Frame1(2).Visible = (TabStrip1.SelectedItem.Index - 1) = 2
    Frame1(3).Visible = (TabStrip1.SelectedItem.Index - 1) = 3
    Frame1(4).Visible = (TabStrip1.SelectedItem.Index - 1) = 4
End Sub
