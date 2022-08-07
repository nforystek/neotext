VERSION 5.00
Object = "{C98B112F-745F-4542-B5B3-DDFADF1F6E2F}#1160.0#0"; "NTControls22.ocx"
Begin VB.Form frmSchOpProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operation Info"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   Icon            =   "frmSchOpProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   960
      Left            =   75
      TabIndex        =   16
      Top             =   1395
      Visible         =   0   'False
      Width           =   5925
      Begin VB.CheckBox Check1 
         Height          =   195
         Left            =   3135
         TabIndex        =   23
         Top             =   255
         Width           =   210
      End
      Begin VB.CheckBox chkNoInterface 
         Caption         =   "Do not allow user interface"
         Height          =   225
         Left            =   525
         TabIndex        =   22
         Top             =   555
         Width           =   2265
      End
      Begin VB.CheckBox chkForceTerminate 
         Caption         =   "Force script to termination."
         Height          =   210
         Left            =   3135
         TabIndex        =   21
         Top             =   555
         Width           =   2220
      End
      Begin VB.TextBox txtSeconds 
         Height          =   285
         Left            =   4125
         TabIndex        =   19
         Text            =   "0"
         Top             =   195
         Width           =   450
      End
      Begin VB.CheckBox chkWaitForScript 
         Caption         =   "Wait for script execution."
         Height          =   240
         Left            =   510
         TabIndex        =   17
         Top             =   240
         Width           =   2235
      End
      Begin VB.Label Label4 
         Caption         =   "seconds."
         Height          =   180
         Left            =   4650
         TabIndex        =   20
         Top             =   255
         Width           =   705
      End
      Begin VB.Label Label2 
         Caption         =   "Time out"
         Height          =   195
         Left            =   3420
         TabIndex        =   18
         Top             =   255
         Width           =   765
      End
   End
   Begin NTControls22.SiteInformation SiteInformation2 
      Height          =   1695
      Left            =   0
      TabIndex        =   7
      Top             =   3180
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   2990
   End
   Begin NTControls22.SiteInformation SiteInformation1 
      Height          =   1695
      Left            =   0
      TabIndex        =   6
      Top             =   1485
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   2990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   345
      Index           =   1
      Left            =   6225
      TabIndex        =   8
      Top             =   1500
      Width           =   1260
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   345
      Index           =   0
      Left            =   6225
      TabIndex        =   9
      Top             =   1965
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      Caption         =   "Operation Properties"
      Height          =   1395
      Left            =   75
      TabIndex        =   10
      Top             =   -15
      Width           =   7440
      Begin VB.CheckBox chkOnlyNewer 
         Caption         =   "Only New"
         Height          =   195
         Left            =   6315
         TabIndex        =   15
         Top             =   1080
         Width           =   1050
      End
      Begin VB.CheckBox chkOverwrite 
         Caption         =   "Overwrite"
         Height          =   195
         Left            =   6315
         TabIndex        =   5
         Top             =   810
         Width           =   990
      End
      Begin VB.CheckBox chkSubFolders 
         Caption         =   "Include Sub Folders"
         Height          =   195
         Left            =   4410
         TabIndex        =   3
         Top             =   810
         Width           =   1770
      End
      Begin VB.TextBox txtRename 
         Height          =   285
         Index           =   1
         Left            =   4485
         TabIndex        =   4
         Top             =   765
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.TextBox txtWildCard 
         Height          =   285
         Left            =   2445
         TabIndex        =   2
         Top             =   765
         Width           =   1695
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   3255
         TabIndex        =   1
         Top             =   345
         Width           =   4050
      End
      Begin VB.ComboBox cmbOperation 
         Height          =   288
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   330
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Action:"
         Height          =   240
         Left            =   210
         TabIndex        =   14
         Top             =   375
         Width           =   570
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Name of new Folder to Create:"
         Height          =   240
         Index           =   2
         Left            =   150
         TabIndex        =   13
         Top             =   810
         Width           =   2220
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Display Note:"
         Height          =   240
         Left            =   2130
         TabIndex        =   12
         Top             =   375
         Width           =   1020
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "to"
         Height          =   240
         Left            =   4110
         TabIndex        =   11
         Top             =   810
         Visible         =   0   'False
         Width           =   390
      End
   End
End
Attribute VB_Name = "frmSchOpProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public myScheduleHWnd As Long

Public OpAction As String
Public OpId As Long
Public SchId As Long

Public Sub ShowOperation(frm As Form, ByVal Action As String, ByVal ScheduleID As Long, Optional ByVal OperationID As Long = -1)
    
    myScheduleHWnd = frm.hwnd
    
    OpAction = LCase(Action)
    
    SchId = ScheduleID
    OpId = OperationID
    If OpAction = "add" Then
        Me.Caption = Action & " Operation [Schedule ID: " & SchId & "]"

    ElseIf OpAction = "edit" Then
        Me.Caption = Action & " Operation [Operation ID: " & OpId & ", Schedule ID: " & ScheduleID & "]"
        
        Dim enc As New NTCipher10.ncode
        Dim dbSchedule As New clsDBSchedule
        
        Dim tmpValue As String
        Dim tmpName As String
        tmpName = dbSettings.CryptKey("", dbSchedule.GetScheduleValue(SchId, "ParentID"))
            
        With dbSchedule
    
            txtName.Text = .GetOperationValue(OpId, "OperationName")
            cmbOperation.ListIndex = IsOnList(cmbOperation, .GetOperationValue(OpId, "Action"))
            chkOverwrite.Value = BoolToCheck(.GetOperationValue(OpId, "Overwrite"))
            chkWaitForScript.Value = chkOverwrite.Value
            chkOnlyNewer.Value = BoolToCheck(.GetOperationValue(OpId, "OnlyNewer"))
            chkForceTerminate.Value = chkOnlyNewer.Value
            chkSubFolders.Value = BoolToCheck(.GetOperationValue(OpId, "SubFolders"))
            chkNoInterface.Value = chkSubFolders.Value
            txtWildCard.Tag = False
            txtWildCard.Text = .GetOperationValue(OpId, "WildCard")
            If cmbOperation.Text = "Script" Then
                If IsNumeric(.GetOperationValue(OpId, "RenameNew")) Then
                    Check1.Value = Abs((CLng(.GetOperationValue(OpId, "RenameNew")) > 0))
                End If
                txtSeconds.Text = .GetOperationValue(OpId, "RenameNew")
            Else
                txtRename(1).Text = .GetOperationValue(OpId, "RenameNew")
            End If
            If IsNumeric(txtRename(1).Text) Then txtSeconds.Text = txtRename(1).Text
            txtWildCard.Tag = True

            tmpValue = .GetOperationValue(OpId, "SURL")
            If Not (tmpValue = "") Then tmpValue = enc.DecryptString(tmpValue, tmpName)
            SiteInformation1.sHostURL.Text = tmpValue
            
            tmpValue = .GetOperationValue(OpId, "SLogin")
            If Not (tmpValue = "") Then tmpValue = enc.DecryptString(tmpValue, tmpName)
            SiteInformation1.sUserName.Text = tmpValue
            
            tmpValue = .GetOperationValue(OpId, "SPass")
            If Not (tmpValue = "") Then tmpValue = enc.DecryptString(tmpValue, tmpName)
            SiteInformation1.sPassword.Text = tmpValue
            
            SiteInformation1.sPort.Text = .GetOperationValue(OpId, "SPort")
            SiteInformation1.sPassive.Value = BoolToCheck(.GetOperationValue(OpId, "SPasv"))
            SiteInformation1.sPortRange.Text = .GetOperationValue(OpId, "SData")
            SiteInformation1.sAdapter.ListIndex = (.GetOperationValue(OpId, "SAdap") - 1)
            SiteInformation1.sssl.Value = .GetOperationValue(OpId, "SSSL")
            
            SiteInformation1.Refresh

            tmpValue = .GetOperationValue(OpId, "DURL")
            If Not (tmpValue = "") Then tmpValue = enc.DecryptString(tmpValue, tmpName)
            SiteInformation2.sHostURL.Text = tmpValue
            
            tmpValue = .GetOperationValue(OpId, "DLogin")
            If Not (tmpValue = "") Then tmpValue = enc.DecryptString(tmpValue, tmpName)
            SiteInformation2.sUserName.Text = tmpValue
            
            tmpValue = .GetOperationValue(OpId, "DPass")
            If Not (tmpValue = "") Then tmpValue = enc.DecryptString(tmpValue, tmpName)
            SiteInformation2.sPassword.Text = tmpValue
            
            SiteInformation2.sPort.Text = .GetOperationValue(OpId, "DPort")
            SiteInformation2.sPassive.Value = BoolToCheck(.GetOperationValue(OpId, "DPasv"))
            SiteInformation2.sPortRange.Text = .GetOperationValue(OpId, "DData")
            SiteInformation2.sAdapter.ListIndex = (.GetOperationValue(OpId, "DAdap") - 1)
            SiteInformation2.sssl.Value = .GetOperationValue(OpId, "DSSL")
            SiteInformation2.Refresh

        End With
        
        Set dbSchedule = Nothing
        Set enc = Nothing
    
    End If
    
    Me.Show
End Sub

Private Sub OpActionDone()
    If OpAction = "add" Or OpId = -1 Then
        Dim NewOp As ScheduleOperationType
        
        NewOp.ScheduleID = SchId
        NewOp.OperationName = txtName.Text
        NewOp.Action = cmbOperation.Text
        
        OpId = AddOperation(NewOp)
        
    End If

    Dim enc As New NTCipher10.ncode
    Dim dbSchedule As New clsDBSchedule

    Dim tmpValue As String
    Dim tmpName As String
    tmpName = dbSettings.CryptKey("", dbSchedule.GetScheduleValue(SchId, "ParentID"))
        
    With dbSchedule
                    
        tmpValue = SiteInformation1.sHostURL.Text
        If Not (tmpValue = "") Then tmpValue = enc.EncryptString(tmpValue, tmpName)
        .SetOperationValue OpId, "SURL", tmpValue
        
        tmpValue = SiteInformation1.sUserName.Text

        If Not (tmpValue = "") Then tmpValue = enc.EncryptString(tmpValue, tmpName)
        .SetOperationValue OpId, "SLogin", tmpValue

        tmpValue = SiteInformation1.sPassword.Text

        If Not (tmpValue = "") Then tmpValue = enc.EncryptString(tmpValue, tmpName)
        .SetOperationValue OpId, "SPass", tmpValue
        
        .SetOperationValue OpId, "SPort", SiteInformation1.sPort.Text
        .SetOperationValue OpId, "SPasv", CBool(SiteInformation1.sPassive.Value)
        .SetOperationValue OpId, "SData", SiteInformation1.sPortRange.Text
        .SetOperationValue OpId, "SAdap", (SiteInformation1.sAdapter.ListIndex + 1)
        .SetOperationValue OpId, "SSSL", SiteInformation1.sssl.Value


        tmpValue = SiteInformation2.sHostURL.Text
        If Not (tmpValue = "") Then tmpValue = enc.EncryptString(tmpValue, tmpName)
        .SetOperationValue OpId, "DURL", tmpValue

        tmpValue = SiteInformation2.sUserName.Text
        If Not (tmpValue = "") Then tmpValue = enc.EncryptString(tmpValue, tmpName)
        .SetOperationValue OpId, "DLogin", tmpValue

        tmpValue = SiteInformation2.sPassword.Text
        If Not (tmpValue = "") Then tmpValue = enc.EncryptString(tmpValue, tmpName)
        .SetOperationValue OpId, "DPass", tmpValue
        
        .SetOperationValue OpId, "DPort", SiteInformation2.sPort.Text
        .SetOperationValue OpId, "DPasv", CBool(SiteInformation2.sPassive.Value)
        .SetOperationValue OpId, "DData", SiteInformation2.sPortRange.Text
        .SetOperationValue OpId, "DAdap", (SiteInformation2.sAdapter.ListIndex + 1)
        .SetOperationValue OpId, "DSSL", SiteInformation2.sssl.Value
        
        .SetOperationValue OpId, "OperationName", txtName.Text
        If OpAction = "add" Then
            .SetOperationValue OpId, "OperationOrder", CLng(CountOperation(SchId) + 1)
        End If
        .SetOperationValue OpId, "Action", cmbOperation.Text
        .SetOperationValue OpId, "Overwrite", CBool(chkOverwrite.Value)
        .SetOperationValue OpId, "OnlyNewer", CBool(chkOnlyNewer.Value)
        .SetOperationValue OpId, "SubFolders", CBool(chkSubFolders.Value)
        .SetOperationValue OpId, "WildCard", txtWildCard.Text
        
        .SetOperationValue OpId, "RenameNew", IIf(cmbOperation.Text = "Script", txtSeconds.Text, txtRename(1).Text)

    End With
    
    Set dbSchedule = Nothing
    Set enc = Nothing
    
    If IsFormVisible(myScheduleHWnd) Then
        GetForm(myScheduleHWnd).RefreshSchedule
    End If

    If Not (ProcessRunning(ServiceFileName) = 0) Then
        MessageQueueAdd ServiceFileName, "/loadschedule " & SchId
    End If
    
End Sub

Public Sub RefreshDisabilities()
    chkNoInterface.enabled = dbSettings.GetPublicSetting("ServiceInterface")
End Sub
Private Sub SetOperationType(ByVal OpType As String)

    Select Case Trim(LCase(OpType))
    
    Case "copy"
        Label6(2).Caption = "Filename, Folder or Wildcard:"
        SiteInformation1.Caption = "Source URL"
        SiteInformation2.Caption = "Destination URL"
        SiteInformation1.Visible = True
        Frame2.Visible = False
        
        txtWildCard.Width = 1695
        
        chkOverwrite.Visible = True
        chkOnlyNewer.Visible = True
        chkOnlyNewer.enabled = (chkOverwrite.Value = 1)
        chkSubFolders.Visible = True
        Label9.Visible = False
        txtRename(1).Visible = False
        
        Height = 5265
        Command1(1).Top = 4020
        Command1(0).Top = 4440
        SiteInformation2.Visible = True
        
    Case "move"
        Label6(2).Caption = "Filename, Folder or Wildcard:"
        SiteInformation1.Caption = "Source URL"
        SiteInformation2.Caption = "Destination URL"
        SiteInformation1.Visible = True
        Frame2.Visible = False
        
        txtWildCard.Width = 1695
        
        chkOverwrite.Visible = True
        chkOnlyNewer.Visible = True
        chkOnlyNewer.enabled = (chkOverwrite.Value = 1)
        chkSubFolders.Visible = True
        Label9.Visible = False
        txtRename(1).Visible = False
        
        Height = 5265
        Command1(1).Top = 4020
        Command1(0).Top = 4440
        SiteInformation2.Visible = True
        
    Case "delete"
        Label6(2).Caption = "Filename, Folder or Wildcard:"
        SiteInformation1.Caption = "Location URL"
        SiteInformation1.Visible = True
        Frame2.Visible = False
        
        txtWildCard.Width = 1695
        
        chkOverwrite.Visible = False
        chkOnlyNewer.Visible = False
        chkSubFolders.Visible = True
        Label9.Visible = False
        txtRename(1).Visible = False
    
        SiteInformation2.Visible = False
        Command1(1).Top = 2325
        Command1(0).Top = 2745
        Height = 3570
    
    Case "folder"
        Label6(2).Caption = "Name of new Folder to Create:"
        SiteInformation1.Caption = "Location URL"
        SiteInformation1.Visible = True
        Frame2.Visible = False
        
        txtWildCard.Width = 1695
        
        chkOverwrite.Visible = False
        chkOnlyNewer.Visible = False
        chkSubFolders.Visible = False
        Label9.Visible = False
        txtRename(1).Visible = False
        
        SiteInformation2.Visible = False
        Command1(1).Top = 2325
        Command1(0).Top = 2745
        Height = 3570
    
    Case "rename"
        Label6(2).Caption = "Rename File or Folder from:"
        SiteInformation1.Caption = "Location URL"
        SiteInformation1.Visible = True
        Frame2.Visible = False
        
        txtWildCard.Width = 1695
        
        chkOverwrite.Visible = False
        chkOnlyNewer.Visible = False
        chkSubFolders.Visible = False
        Label9.Visible = True
        txtRename(1).Visible = True
        
        SiteInformation2.Visible = False
        Command1(1).Top = 2325
        Command1(0).Top = 2745
        Height = 3570
    Case "script"
        Label6(2).Caption = "Path of .mprj script file:"
        Frame2.Visible = True
        
        SiteInformation1.Visible = False
        SiteInformation2.Visible = False

        txtWildCard.Width = 4845
        
        chkNoInterface.enabled = dbSettings.GetPublicSetting("ServiceInterface")
        
        If chkWaitForScript.Value = 1 And chkWaitForScript.enabled Then
        
            Check1.enabled = (chkWaitForScript.Value = 1) Or (IsNumeric(txtSeconds.Text) And txtSeconds.Text <> "")
    
            txtSeconds.enabled = (Check1.Value = 1) And Check1.enabled Or (IsNumeric(txtSeconds.Text) And txtSeconds.Text <> "")
            txtSeconds_Change
            If ((chkWaitForScript.Value = 1) And Check1.enabled And (IsNumeric(txtSeconds.Text) And txtSeconds.Text <> "")) Then
                Check1.Value = 0
            End If
            
            chkForceTerminate.enabled = Check1.enabled And Check1.Value = 1
        Else
            Check1.enabled = False
            Check1.Value = 0
            txtSeconds.enabled = False
            txtSeconds.Text = ""
            chkForceTerminate.enabled = False
            chkForceTerminate.Value = 0
        End If
        
        chkOverwrite.Visible = False
        chkOnlyNewer.Visible = False
        chkSubFolders.Visible = False
        Label9.Visible = False
        txtRename(1).Visible = False
        
        Command1(1).Top = 1500
        Command1(0).Top = 1965
        Height = 2805
    End Select

End Sub

Private Sub Check1_Click()
    txtSeconds.enabled = (Check1.Value = 1) And Check1.enabled
    If Check1.Value = 0 Then txtSeconds.Text = ""
    chkForceTerminate.enabled = Check1.enabled And Check1.Value = 1
'    If Check1.Value = 0 Or Not Check1.enabled Then
'        txtSeconds.Text = ""
    If Check1.enabled And Check1.Value = 1 Then
        If Not IsNumeric(txtRename(1).Text) Then
            txtSeconds.Text = "1"
        Else
            If CLng(txtSeconds.Text) <= 0 Then txtSeconds.Text = "1"
        End If
        If Not IsNumeric(txtSeconds.Text) Then txtSeconds.Text = "1"
    End If
   ' SetOperationType cmbOperation.Text
End Sub

Private Sub chkWaitForScript_Click()
    chkOverwrite.Value = chkWaitForScript.Value
    SetOperationType cmbOperation.Text
End Sub

Private Sub chkForceTerminate_Click()
    chkOnlyNewer.Value = chkForceTerminate.Value
    'SetOperationType cmbOperation.Text
End Sub

Private Sub chkNoInterface_Click()
    chkSubFolders.Value = chkNoInterface.Value
    SetOperationType cmbOperation.Text
End Sub

Private Sub chkOverwrite_Click()
    SetOperationType cmbOperation.Text
End Sub

Private Sub cmbOperation_Click()
    SetOperationType cmbOperation.Text
End Sub

Private Sub Combo1_Click()
    SetOperationType cmbOperation.Text
End Sub

Private Sub Command1_Click(Index As Integer)
    frmMain.ValidDataPortRange SiteInformation1.sPortRange
    frmMain.ValidDataPortRange SiteInformation2.sPortRange

    Select Case Index
        Case 1

            
            If cmbOperation.Text = "Folder" And txtWildCard.Text = "" Then
                MsgBox "You must enter a folder name to perform this operation on.", vbInformation, AppName
                txtWildCard.SetFocus
                Exit Sub
            End If
                
            If cmbOperation.Text = "Rename" And txtWildCard.Text = "" Then
                MsgBox "You must enter a folder or file name to rename from.", vbInformation, AppName
                txtWildCard.SetFocus
                Exit Sub
            End If
            If cmbOperation.Text = "Script" Then
                If txtWildCard.Text = "" Then
                    MsgBox "You must enter a fully qualified path to an existing .mprj script file.", vbInformation, AppName
                    txtWildCard.SetFocus
                    Exit Sub
                ElseIf Not PathExists(txtWildCard.Text, True) Then
                    MsgBox "You must enter a fully qualified path to an existing .mprj script file.", vbInformation, AppName
                    txtWildCard.SetFocus
                    Exit Sub
                End If
            End If
            
            If cmbOperation.Text = "Rename" And txtRename(1).Text = "" Then
                MsgBox "You must enter a folder or file name to rename to.", vbInformation, AppName
                txtRename(1).SetFocus
                Exit Sub
            End If
            
            If (cmbOperation.Text = "Copy" Or cmbOperation.Text = "Move" Or cmbOperation.Text = "Delete") And txtWildCard.Text = "" Then
                    
                MsgBox "You must enter a filename or wildcard to perform this operation on.", vbInformation, AppName
                
                txtWildCard.SetFocus
                Exit Sub
            End If
                
            If (cmbOperation.Text = "Copy" Or cmbOperation.Text = "Move") Then
                
                With SiteInformation2
                
                    If .sHostURL.Text = "" Then
                        MsgBox "You must enter the destination URL for this operation.", vbInformation, AppName
                        .sHostURL.SetFocus
                        Exit Sub
                    End If
            
                    If Not IsNumeric(.sPort.Text) Then
                        MsgBox "You must enter a numeric value for the destination port.", vbInformation, AppName
                        .sPort.SetFocus
                        Exit Sub
                    End If
            
                    If .sUserName.Text = "" And Not .sUserName.Locked Then
                        MsgBox "You must enter a username for this URL.", vbInformation, AppName
                        .sUserName.SetFocus
                        Exit Sub
                    End If

                    If .sPassword.Text = "" And Not .sPassword.Locked Then
                        MsgBox "You must enter a password for this URL.", vbInformation, AppName
                        .sPassword.SetFocus
                        Exit Sub
                    End If
                
                End With
            
            End If
            
            If Not cmbOperation.Text = "Script" Then
            
                With SiteInformation1
                
                    If .sHostURL.Text = "" Then
                        MsgBox "You must enter the source URL for this operation.", vbInformation, AppName
                        .sHostURL.SetFocus
                        Exit Sub
                    End If
                
                        If Not IsNumeric(.sPort.Text) Then
                            MsgBox "You must enter a numeric value for the source port.", vbInformation, AppName
                            .sPort.SetFocus
                            Exit Sub
                        End If
                
                    If .sUserName.Text = "" And Not .sUserName.Locked Then
                        MsgBox "You must enter a username for this URL.", vbInformation, AppName
                        .sUserName.SetFocus
                        Exit Sub
                    End If
    
                    If .sPassword.Text = "" And Not .sPassword.Locked Then
                        MsgBox "You must enter a password for this URL.", vbInformation, AppName
                        .sPassword.SetFocus
                        Exit Sub
                    End If
                
                End With
            
            End If
            
            OpActionDone
            
        Case 0
    
    End Select
    Unload Me
End Sub

Private Sub Form_Load()
    
    Me.Left = ((Screen.Width / 2) - (Me.Width / 2))
    Me.Top = ((Screen.Height / 2) - (Me.Height / 2))
    Dim frm As Form
    For Each frm In Forms
        If (TypeName(frm) = TypeName(Me)) Then
            If Not (frm.hwnd = Me.hwnd) Then
                Me.Left = frm.Left + (Screen.TwipsPerPixelX * 32)
                Me.Top = frm.Top + (Screen.TwipsPerPixelY * 32)
            End If
        End If
    Next
    
    If ((Me.Left + Me.Width) > Screen.Width) Or _
        ((Me.Top + Me.Height) > Screen.Height) Then
        Me.Left = (32 * Screen.TwipsPerPixelX)
        Me.Top = (32 * Screen.TwipsPerPixelY)
    End If
    
    SiteInformation1.sSavePass.Visible = False
    SiteInformation2.sSavePass.Visible = False

    cmbOperation.AddItem "Copy"
    cmbOperation.AddItem "Move"
    cmbOperation.AddItem "Delete"
    cmbOperation.AddItem "Folder"
    cmbOperation.AddItem "Rename"
    cmbOperation.AddItem "Script"
    cmbOperation.ListIndex = 0

    SetAutoTypeList Me, SiteInformation1.AutoTypeCombo
    SetAutoTypeList Me, SiteInformation2.AutoTypeCombo

    SiteInformation1.sPassive.Value = BoolToCheck(dbSettings.GetProfileSetting("ConnectionMode") = 0)
    SiteInformation2.sPassive.Value = BoolToCheck(dbSettings.GetProfileSetting("ConnectionMode") = 0)
   
    SiteInformation1.sssl.Value = BoolToCheck(dbSettings.GetProfileSetting("SSL") = 0)
    SiteInformation2.sssl.Value = BoolToCheck(dbSettings.GetProfileSetting("SSL") = 0)
   
    SiteInformation1.sPortRange.Text = dbSettings.GetProfileSetting("DefaultPortRange")
    SiteInformation2.sPortRange.Text = dbSettings.GetProfileSetting("DefaultPortRange")
    
    SiteInformation1.sAdapter.ListIndex = (dbSettings.GetProfileSetting("AdapterIndex") - 1)
    SiteInformation2.sAdapter.ListIndex = (dbSettings.GetProfileSetting("AdapterIndex") - 1)
    
    SiteInformation1.ShowAdvSettings = dbSettings.GetProfileSetting("ShowAdvSettings")
    If Not SiteInformation1.ShowAdvSettings Then
        SiteInformation1.sPassive.Value = IIf((dbSettings.GetProfileSetting("ConnectionMode") = 0), 1, 0)
    End If
    SiteInformation2.ShowAdvSettings = dbSettings.GetProfileSetting("ShowAdvSettings")
    If Not SiteInformation2.ShowAdvSettings Then
        SiteInformation2.sPassive.Value = IIf((dbSettings.GetProfileSetting("ConnectionMode") = 0), 1, 0)
    End If

End Sub

Private Sub txtRename_Change(Index As Integer)

    If txtRename(1).Visible Then
        If (InStr(txtWildCard.Text, "/") > 0) Then
            txtRename(1).Text = "/" & Replace(Replace(txtRename(1).Text, "\", ""), "/", "")
            If Len(txtRename(1).Text) = 1 Then txtRename(1).SelStart = 1
        ElseIf (InStr(txtWildCard.Text, "\") > 0) Then
            txtRename(1).Text = "/" & Replace(Replace(txtRename(1).Text, "\", ""), "/", "")
            If Len(txtRename(1).Text) = 1 Then txtRename(1).SelStart = 1
        Else
            txtRename(1).Text = Replace(Replace(txtRename(1).Text, "/", ""), "\", "")
        End If
    Else
        'txtRename(1).Text = ""
    End If

End Sub

Private Sub txtSeconds_Change()
    If IsNumeric(txtSeconds.Text) And (txtSeconds.Text <> "") Then
        chkForceTerminate.enabled = (CLng(txtSeconds.Text) > 0) And (Check1.Value = 1)
    Else
        chkForceTerminate.enabled = False
    End If
    chkForceTerminate.enabled = txtSeconds.enabled
    If cmbOperation.Text = "Script" Then txtRename(1).Text = txtSeconds.Text
End Sub

Private Sub txtWildCard_Change()
    
    If Not cmbOperation.Text = "Script" Then
    
        If (InStr(txtWildCard.Text, "/") > 0) Then
            txtWildCard.Text = "/" & Replace(Replace(txtWildCard.Text, "/", ""), "\", "")
            If Len(txtWildCard.Text) = 1 Then txtWildCard.SelStart = 1
        ElseIf (InStr(txtWildCard.Text, "\") > 0) Then
            txtWildCard.Text = "/" & Replace(Replace(txtWildCard.Text, "/", ""), "\", "")
            If Len(txtWildCard.Text) = 1 Then txtWildCard.SelStart = 1
        Else
            txtWildCard.Text = Replace(Replace(txtWildCard.Text, "/", ""), "\", "")
        End If
        
        txtRename_Change 1
    End If
End Sub
