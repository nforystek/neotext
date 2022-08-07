VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Install"
   ClientHeight    =   4020
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   7140
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   36
      Top             =   3492
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   3870
      Left            =   1605
      ScaleHeight     =   3870
      ScaleWidth      =   5490
      TabIndex        =   0
      Top             =   72
      Width           =   5496
      Begin VB.CommandButton Command1 
         Caption         =   "&View Log"
         Height          =   375
         Index           =   2
         Left            =   1815
         TabIndex        =   7
         Top             =   3495
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Exit"
         Height          =   375
         Index           =   1
         Left            =   4305
         TabIndex        =   3
         Top             =   3495
         Width           =   1155
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Kill App"
         Height          =   375
         Index           =   0
         Left            =   3060
         TabIndex        =   2
         Top             =   3495
         Width           =   1155
      End
      Begin VB.TextBox BulkText 
         Height          =   3225
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Text            =   "frmMain.frx":0E42
         Top             =   210
         Width           =   5448
      End
      Begin VB.PictureBox CloseFInish 
         BorderStyle     =   0  'None
         Height          =   3480
         Left            =   0
         ScaleHeight     =   3480
         ScaleWidth      =   5445
         TabIndex        =   5
         Top             =   0
         Width           =   5448
         Begin VB.Label Message 
            Alignment       =   2  'Center
            Caption         =   "Please close all "
            Height          =   624
            Left            =   84
            TabIndex        =   6
            Top             =   1308
            Width           =   5232
         End
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H8000000D&
         FillColor       =   &H8000000D&
         FillStyle       =   0  'Solid
         Height          =   228
         Left            =   48
         Top             =   2952
         Visible         =   0   'False
         Width           =   12
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000010&
         Height          =   240
         Left            =   36
         Top             =   2952
         Visible         =   0   'False
         Width           =   5400
      End
      Begin VB.Label TextStatus 
         Height          =   204
         Left            =   72
         TabIndex        =   4
         Top             =   3204
         Visible         =   0   'False
         Width           =   5340
      End
      Begin VB.Label TopText 
         Caption         =   "Accept the license agreement below by clicking the agree button to install:"
         Height          =   240
         Left            =   36
         TabIndex        =   1
         Top             =   -24
         Width           =   5460
      End
   End
   Begin VB.Image Image1 
      Height          =   3870
      Left            =   75
      Picture         =   "frmMain.frx":2348
      Stretch         =   -1  'True
      Top             =   75
      Width           =   1425
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Public Enum SystemStates
    Installation = 0
    Uninstallation = 1
End Enum

Public Enum WizardStates
    CloseAllWindow = 0
    AgreementWindow = 1
    SummaryInformation = 2
    LogSystemChanges = 3
    RollbackChanges = 4
    CompletedWindow = 5
    CanceledWindow = 6
End Enum

Public Enum ButtonStates
    IdleWizard = 0
    ActingWizard = 1
    CacnelingWizard = 2
    FinishedWizard = 3
    CanceledWizard = 4
End Enum

Private Type WindowView
    SystemState As SystemStates
    WizardState As WizardStates
    ButtonState As ButtonStates
End Type

Private MyView As WindowView

Private Sub ChangeViewState(ByRef View As WindowView)
    Select Case View.WizardState
        Case CanceledWindow
            Select Case View.SystemState
                Case Installation
                    TopText.Caption = "The installation process was aborted, any changes have been reversed:"
                    BulkText.text = BulkText.text & vbCrLf & "Roll back of changes has finished... " & Now
                Case Uninstallation
                    TopText.Caption = "The uninstallation process was aborted, full removal has not completed:"
                    BulkText.text = BulkText.text & vbCrLf & "Stopping of uninstalltion has finished... " & Now
            End Select
        Case AgreementWindow
            Select Case View.SystemState
                Case Installation
                    TopText.Caption = "Accept the license agreement below by clicking the agree button to isntall:"
                Case Uninstallation
            End Select
        Case SummaryInformation
            Select Case View.SystemState
                Case Installation
                    TopText.Caption = "Below is a list of gathered information that will be installed to the system:"
                    BulkText.text = Program.Summary
                Case Uninstallation
                    TopText.Caption = "Below is a list of gathered information that will be removed from the system:"
                    BulkText.text = Program.Summary
            End Select
        Case LogSystemChanges
            Select Case View.ButtonState
                Case ButtonStates.ActingWizard
                    Select Case View.SystemState
                        Case Installation
                            TopText.Caption = "Install is currently preforming following logged changes made to the system:"
                            BulkText.text = "Starting installation process... " & Now
                        Case Uninstallation
                            TopText.Caption = "Uninstall is currently preforming following logged changes made to the system:"
                            BulkText.text = "Starting uninstallation process... " & Now
                    End Select
                Case ButtonStates.IdleWizard
                    Select Case View.SystemState
                        Case Installation
                            TopText.Caption = "Canceling current process and rolling back any changes made to the system:"
                            BulkText.text = BulkText.text & vbCrLf & "Cancel requested by user... " & Now
                        Case Uninstallation
                            TopText.Caption = "Canceling current process and rolling back any changes made to the system:"
                            BulkText.text = BulkText.text & vbCrLf & "Cancel requested by user... " & Now
                    End Select
            End Select
        Case RollbackChanges

        Case CompletedWindow
            Select Case View.SystemState
                Case Installation
                    TopText.Caption = "Installation is completed.  Below is a log of all cahnges made to your system:"
                    BulkText.text = BulkText.text & vbCrLf & "Installation process has finished... " & Now
                    Command1(2).Visible = True
                    CloseFInish.Visible = True
                    BulkText.Visible = False
                    Message.Caption = "Installation of " & Program.Display & " has completed successfully."
                Case Uninstallation
                    TopText.Caption = "Uninstallation finished.  Below is a log of all cahnges made to your system:"
                    BulkText.text = BulkText.text & vbCrLf & "Uninstallation process has finished... " & Now
                    Command1(2).Visible = True
                    CloseFInish.Visible = True
                    BulkText.Visible = False
                    Message.Caption = "Uninstallation of " & Program.Display & " has completed successfully."
            End Select
    End Select

    Select Case View.ButtonState
        Case IdleWizard
            Select Case View.SystemState
                Case Installation
                    Command1(0).Caption = "&Install"
                    If MyView.WizardState = SummaryInformation Then
                        Command1(1).Caption = "&Exit"
                    Else
                        Command1(1).Caption = "&Disagree"
                    End If
                Case Uninstallation
                    Command1(0).Caption = "&Uninstall"
                    Command1(1).Caption = "&Exit"
            End Select
            Command1(0).Enabled = True
            Command1(0).Visible = True
            Command1(1).Enabled = True
            Command1(1).Visible = True
            BulkText.Height = 3225
            Shape1.Visible = False
            Shape2.Visible = False
            TextStatus.Visible = False
        Case ActingWizard
            Select Case View.SystemState
                Case Installation
                    Command1(0).Caption = "&Installings"
                    Command1(1).Caption = "&Cancel"
                    Command1(1).Enabled = True
                Case Uninstallation
                    Command1(0).Caption = "&Uninstalling"
                    Command1(1).Caption = "&Working"
                    Command1(1).Enabled = False
            End Select
            Command1(0).Enabled = False
            Command1(0).Visible = True
            Command1(1).Visible = True
            BulkText.Height = 2793
            Shape1.Visible = True
            Shape2.Visible = True
            TextStatus.Visible = True
        Case CacnelingWizard
            Select Case View.SystemState
                Case Installation
                    Command1(0).Caption = "&Installings"
                    Command1(1).Caption = "&Working"
                Case Uninstallation
                    Command1(0).Caption = "&Uninstalling"
                    Command1(1).Caption = "&Working"
            End Select
            Command1(0).Enabled = False
            Command1(0).Visible = True
            Command1(1).Enabled = False
            Command1(1).Visible = True
            BulkText.Height = 2793
            Shape1.Visible = True
            Shape2.Visible = True
            TextStatus.Visible = True
        Case FinishedWizard
            Select Case View.SystemState
                Case Installation
                    Command1(0).Caption = "&Installed"
                    Command1(1).Caption = "&Exit"
                Case Uninstallation
                    Command1(0).Caption = "&Uninstalled"
                    Command1(1).Caption = "&Exit"
            End Select
            Command1(0).Enabled = False
            Command1(0).Visible = True
            Command1(1).Enabled = True
            Command1(1).Visible = True
            BulkText.Height = 3225
            Shape1.Visible = False
            Shape2.Visible = False
            TextStatus.Visible = False
        Case CanceledWizard
            Select Case View.SystemState
                Case Installation
                    Command1(0).Caption = "&Aborted"
                    Command1(1).Caption = "&Exit"
                Case Uninstallation
                    Command1(0).Caption = "&Aborted"
                    Command1(1).Caption = "&Exit"
            End Select
            Command1(0).Enabled = False
            Command1(0).Visible = True
            Command1(1).Enabled = True
            Command1(1).Visible = True
            BulkText.Height = 2793
            Shape1.Visible = True
            Shape2.Visible = True
            TextStatus.Visible = True
    End Select
End Sub

Private Sub SetProgress()
    Dim percentage As Single
    If Program.TotalOfItem > 0 Then
        percentage = (((Program.CurrentItem / 100) / Program.TotalOfItem) * 100)
        percentage = (Shape1.Width - Shape1.Left) * percentage
        If Not Shape2.Width = percentage Then
            If percentage > 0 Then
                Shape2.Width = percentage
            Else
                Shape2.Width = 0
            End If
        End If
    ElseIf Program.TotalOfBytes > 0 Then
        percentage = (((Program.ByteProgress / 100) / Program.TotalOfBytes) * 100)
        percentage = (Shape1.Width - Shape1.Left) * percentage
        If Not Shape2.Width = percentage Then
            If percentage > 0 Then
                Shape2.Width = percentage
            Else
                Shape2.Width = 0
            End If
        End If
    End If
End Sub

Public Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 2
            If CloseFInish.Visible Then
                Command1(2).Caption = "&Save Log"
                CloseFInish.Visible = False
                BulkText.Visible = True
            Else
                On Error Resume Next
                WriteFile GetWindowsTempFolder & Program.AppValue & "_InstallStar.log", BulkText.text
                If Err Then
                    MsgBox "Unable to save log file: " & GetWindowsTempFolder & Program.AppValue & "_InstallStar.log"
                Else
                    Select Case MsgBox("Log file saved to: " & GetWindowsTempFolder & Program.AppValue & "_InstallStar.log", vbOKCancel)
                        Case vbOK
                        Case vbCancel
                            Kill GetWindowsTempFolder & Program.AppValue & "_InstallStar.log"
                    End Select
                End If
                On Error GoTo 0
                Command1(2).Visible = False
            End If
        Case 0
            If MyView.WizardState = CloseAllWindow Then
                Program.CloseAll
                CloseChecking
            ElseIf MyView.WizardState = AgreementWindow Then
            
                MyView.ButtonState = IdleWizard
                MyView.WizardState = SummaryInformation
                BulkText.text = ""
                ChangeViewState MyView

            ElseIf MyView.WizardState = SummaryInformation Then
            
                If MyView.SystemState = Uninstallation Then
                    If Program.LockUninstalling Then
                        MyView.ButtonState = ActingWizard
                        MyView.WizardState = LogSystemChanges
                        ChangeViewState MyView
                        Timer1.Enabled = True
                    Else
                        MsgBox "There was a problem starting installation lock.", vbCritical
                    End If
                ElseIf MyView.SystemState = Installation Then

                    If Program.LockInstalling Then
                        MyView.ButtonState = ActingWizard
                        MyView.WizardState = LogSystemChanges
                        ChangeViewState MyView
                        Timer1.Enabled = True
                    Else
                        MsgBox "There was a problem starting uninstallation lock.", vbCritical
                    End If
                End If
            End If
        Case 1
            If MyView.ButtonState = ActingWizard Then
                If MyView.SystemState = Uninstallation Then
                    MyView.ButtonState = CacnelingWizard
                    Program.UnlockUninstalling
                    ChangeViewState MyView
                ElseIf MyView.SystemState = Installation Then
                    MyView.ButtonState = CacnelingWizard
                    Program.UnlockInstalling
                    ChangeViewState MyView
                End If
            Else

                Unload Me

            End If
    End Select
End Sub

Public Sub StartWizard()
    If SimSilence = InstallMode.Normal And Not Program.Installed Then
        BulkText.text = StrConv(LoadResData(LicenseAgreement, "CUSTOM"), vbUnicode)

    End If

    Message.Caption = "Please close all applications related to " & Program.Display & "," & vbCrLf & "Including prior version and services that were installed."

    MyView.ButtonState = IdleWizard
    Timer1.Enabled = True

End Sub


Private Sub Form_Load()
        If Program.Installed Then
            Me.Caption = "Uninstall " & Program.Display

        Else
            Me.Caption = "Install " & Program.Display
        End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If (MyView.ButtonState <> CanceledWizard) And (MyView.ButtonState <> FinishedWizard) And (MyView.WizardState <> AgreementWindow) Then
        If MsgBox("Are you sure you want to exit the wizard?", vbYesNo + vbQuestion, "Installation") = vbNo Then
            Cancel = True
        Else
            Timer1.Enabled = False
            
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Program.Finish
    End
End Sub
Private Sub CloseChecking()
    If Not Program.IsRunning Then
        CloseFInish.Visible = False
        
        Command1(0).Caption = "&Agree"
        If MyView.WizardState = SummaryInformation Then
            Command1(1).Caption = "&Exit"
        Else
            Command1(1).Caption = "&Disagree"
        End If
        
        Timer1.Enabled = False
        MyView.ButtonState = ButtonStates.IdleWizard
        If Program.Installed Then
            Me.Caption = "Uninstall " & Program.Display
            MyView.WizardState = WizardStates.SummaryInformation
            MyView.SystemState = SystemStates.Uninstallation
            ChangeViewState MyView
        Else
            Me.Caption = "Install " & Program.Display
            MyView.WizardState = WizardStates.AgreementWindow
            MyView.SystemState = SystemStates.Installation
        End If
    
    End If
End Sub

Private Sub Timer1_Timer()
    If Timer1.Interval = 100 Then
        If SimSilence = InstallMode.Normal Then
            Me.Show
        Else
            Do Until MyView.WizardState = SummaryInformation
                Command1_Click 0
            Loop
            Command1_Click 0
        End If
        Timer1.Interval = 3
    Else
    
        If CloseFInish.Visible Then
        
            CloseChecking
        Else
            
            Dim txtInfo As String
            Dim diff As Single
            Dim newtext As String
            If (MyView.WizardState = RollbackChanges) Then
        
                If Program.NextProgressive(txtInfo) Then
                
                    If txtInfo <> "" Then
                        If Me.TextWidth(txtInfo) > TextStatus.Width Then
                            diff = Me.TextWidth(NextArg(txtInfo, ":") & ": ")
                            newtext = RemoveArg(txtInfo, ":")
                            
                            Do While ((Me.TextWidth("...." & newtext) + diff) > TextStatus.Width) And (Not (newtext = ""))
                                newtext = Mid(newtext, 2)
                            Loop
                            If newtext = "" Then newtext = StrReverse(NextArg(StrReverse(RemoveArg(txtInfo, ":")), "\"))
                            newtext = NextArg(txtInfo, ":") & ": " & " ..." & newtext
                        Else
                            newtext = txtInfo
                        End If
                            
                        BulkText.text = BulkText.text & vbCrLf & txtInfo
                        BulkText.SelStart = InStrRev(BulkText.text, vbCrLf) + 2
                        TextStatus.Caption = newtext
                        SetProgress
                    End If
                    
                ElseIf Program.Installed Then
                    If Program.UnlockUninstalling Then
                        MyView.WizardState = CanceledWindow
                        MyView.ButtonState = CanceledWizard
                        ChangeViewState MyView
                        If SimSilence <> Normal Then Unload Me
                    End If
                Else
                    If Program.UnlockInstalling Then
                        MyView.WizardState = CanceledWindow
                        MyView.ButtonState = CanceledWizard
                        ChangeViewState MyView
                        If SimSilence <> Normal Then Unload Me
                    End If
                End If
        
                
            ElseIf MyView.WizardState = LogSystemChanges Then
                If MyView.ButtonState = CacnelingWizard Then
                    BulkText.text = BulkText.text & "Rolling back any changes... " & vbCrLf
                    
                    MyView.WizardState = RollbackChanges
                    ChangeViewState MyView
                    
                    
                Else
                    If Program.NextProgressive(txtInfo) Then
                    
                        If txtInfo <> "" Then
                        
                            If Me.TextWidth(txtInfo) > TextStatus.Width Then
                                diff = Me.TextWidth(NextArg(txtInfo, ":") & ": ")
                                newtext = RemoveArg(txtInfo, ":")
                                
                                Do While ((Me.TextWidth("...." & newtext) + diff) > TextStatus.Width) And (Not (newtext = ""))
                                    newtext = Mid(newtext, 2)
                                Loop
                                If newtext = "" Then newtext = StrReverse(NextArg(StrReverse(RemoveArg(txtInfo, ":")), "\"))
                                newtext = NextArg(txtInfo, ":") & ": " & " ..." & newtext
                            Else
                                newtext = txtInfo
                            End If
                                
                            BulkText.text = BulkText.text & vbCrLf & txtInfo
                            BulkText.SelStart = InStrRev(BulkText.text, vbCrLf) + 2
                            TextStatus.Caption = newtext
                            SetProgress
                        End If
                    ElseIf Program.Installed Then
                        If Program.UnlockUninstalling Then
                            MyView.WizardState = CompletedWindow
                            MyView.ButtonState = FinishedWizard
                            ChangeViewState MyView
                            
                            Timer1.Enabled = False
                            If SimSilence <> Normal Then
                                Command1_Click 1
                            End If
                        End If
                    Else
                        If Program.UnlockInstalling Then
                            MyView.WizardState = CompletedWindow
                            MyView.ButtonState = FinishedWizard
                            ChangeViewState MyView
                            Timer1.Enabled = False
                            If SimSilence <> Normal Then
                                Command1_Click 1
                            End If
                        End If
                    End If
                    
                End If
            
            End If
        End If
    End If
End Sub
