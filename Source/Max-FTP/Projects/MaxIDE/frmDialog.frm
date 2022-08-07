VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDialog 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8235
   Icon            =   "frmDialog.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   360
      Index           =   1
      Left            =   7290
      TabIndex        =   2
      Top             =   5910
      Width           =   885
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Next"
      Default         =   -1  'True
      Height          =   360
      Index           =   0
      Left            =   6315
      TabIndex        =   1
      Top             =   5910
      Width           =   885
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5280
      Left            =   150
      TabIndex        =   0
      Top             =   450
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   9313
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ClassID"
         Object.Width           =   4180
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   10029
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5760
      Left            =   60
      TabIndex        =   3
      Top             =   75
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   10160
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public IsOk As Boolean
Public ItemName As String
Public ItemClass As String

Public Function Finish(ByVal pOK As Boolean)
    IsOk = pOK
    If IsOk Then
        ItemName = ListView1.SelectedItem.Text
        ItemClass = ListView1.SelectedItem.SubItems(1)
    End If
    Me.Hide
End Function

Public Function InitDialog(ByVal Dialog As Integer)
    TabStrip1.Tabs.Clear
    
    Set ListView1.Icons = frmMainIDE.CommonIcons
    Set ListView1.SmallIcons = frmMainIDE.CommonIcons
    
    Select Case Dialog
        Case 0
            Me.Caption = AppName
            TabStrip1.Tabs.Add , , "New Project"
        
            AddIcon ScriptFile, "JScript", "Project", "Start a new Max-FTP JavaScript Project."
            AddIcon ScriptFile, "VBScript", "Project", "Start a new Max-FTP VBScript Project."
            'AddIcon "FolderOpen", "PGP Batch", "Script", "Start a new Max-FTP PGP Batch Script."
            If IsDebugger Then
                AddIcon "FolderOpen", "Template", "Script", "Start a new Max-FTP Template Script."
            End If
            
            ListView1.Height = 1635
            TabStrip1.Height = 2145
            Command1(0).Top = 2265
            Command1(1).Top = 2265
            Me.Height = 3030
            
        Case 1
            Me.Caption = "Add Item to Project"
            TabStrip1.Tabs.Add , , "New Item"
            
            AddIcon ModuleFile, "Module", "MaxIDE.Module", "Generic script module where you can write subs and functions."
            AddIcon ModuleFile, "Generic", "MaxIDE.Generic", "Generic non-event driven object template to add your own objects."
            AddIcon ObjectFile, "Debug", "MaxIDE.Debug", "Modifies the text in the MaxIDE Debug Window for run-time debugging."
            AddIcon ObjectFile, "Events", "MaxIDE.Events", "Allows you to add events to the MaxFTP Global Event log."
            AddIcon ObjectFile, "Collect", "MaxIDE.Collect", "An object orientated wrapper for maintaining a collection of strings."
            AddIcon ObjectFile, "Client", "NTAdvFTP61.Client", "Advanced event driven FTP/Local client object which MaxFTP uses."
            AddIcon ObjectFile, "Socket", "NTAdvFTP61.Socket", "Uses Winsock and the TCP/IP protocol to manage data connections."
            AddIcon ObjectFile, "URL", "NTAdvFTP61.URL", "Modify, parse and validate functions for remote or local URL strings."
            AddIcon ObjectFile, "Window", "NTPopup21.Window", "Popup window with large message and web link capabilities."
            AddIcon ObjectFile, "Timer", "NTSchedule20.Timer", "Event driven timer object that is based on milliseconds."
            AddIcon ObjectFile, "Schedule", "NTSchedule20.Schedule", "Event driven schedule object with date settings and abilities."
            AddIcon ObjectFile, "Process", "NTShell22.Process", "Manages and executes applications (exe, bat, com) with wait options."
            AddIcon ObjectFile, "Internet", "NTShell22.Internet", "Retrieve html text from websites with the option to post form data."
            AddIcon ObjectFile, "Player", "NTSound20.Player", "Plays many types of audio files with optional play back looping."
            AddIcon ObjectFile, "NCode", "NTCipher10.NCode", "Encrypts and Decrypts data key strands using NForystek's method."
            AddIcon ObjectFile, "UUCode", "NTCipher10.UUCode", "Used to encode or decode binary files to UU code format."
            AddIcon ObjectFile, "BWord", "NTCipher10.BWord", "Used to preforming bitwise, hi, and low set/get on any number."
            AddIcon ObjectFile, "PoolID", "NTCipher10.PoolID", "Generates unique ID's with in a pool context of instance only."
            AddIcon ObjectFile, "GUID", "NTCipher10.GUID", "Allows for easily generating of a globally unique identifier."
            
            ListView1.Height = 5280
            TabStrip1.Height = 5760
            Command1(0).Top = 5910
            Command1(1).Top = 5910
            Me.Height = 6675
    
    End Select

    ListView1.ListItems(1).Selected = True
    Me.Show 1
End Function

Public Function AddIcon(ByVal nIconKey As String, ByVal nName As String, ByVal nClass As String, ByVal nDescription As String)
    Dim Node As ListItem
    Set Node = ListView1.ListItems.Add(, , nName, nIconKey, nIconKey)
    Node.SubItems(1) = nClass
    Node.SubItems(2) = nDescription
    Set Node = Nothing
End Function

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            If Not ListView1.SelectedItem Is Nothing Then
                Finish True
            End If
        Case 1
            Finish False
    End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = True
        Finish False
    End If
End Sub

Private Sub ListView1_DblClick()
    Command1_Click 0
End Sub



