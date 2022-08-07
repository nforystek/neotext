VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSchOperations 
   AutoRedraw      =   -1  'True
   Caption         =   "Operations"
   ClientHeight    =   2640
   ClientLeft      =   168
   ClientTop       =   456
   ClientWidth     =   7452
   Icon            =   "frmSchOperations.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   7452
   Tag             =   "schedule"
   Visible         =   0   'False
   Begin VB.PictureBox dContainer 
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   3
      Left            =   1740
      ScaleHeight     =   732
      ScaleWidth      =   5196
      TabIndex        =   1
      Top             =   495
      Width           =   5190
      Begin MSComctlLib.Toolbar SchControls 
         Height          =   450
         Left            =   1695
         TabIndex        =   2
         Top             =   150
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   804
         ButtonWidth     =   826
         ButtonHeight    =   804
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
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
      Begin VB.Line Line8 
         BorderColor     =   &H80000014&
         Index           =   3
         Tag             =   "highlight"
         X1              =   540
         X2              =   1680
         Y1              =   705
         Y2              =   705
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
      Begin VB.Line Line6 
         BorderColor     =   &H80000014&
         Index           =   3
         Tag             =   "highlight"
         X1              =   465
         X2              =   1605
         Y1              =   180
         Y2              =   180
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
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         Index           =   3
         Tag             =   "shadow"
         X1              =   1815
         X2              =   1815
         Y1              =   135
         Y2              =   630
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
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         Tag             =   "shadow"
         X1              =   120
         X2              =   120
         Y1              =   60
         Y2              =   720
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1170
      Left            =   315
      TabIndex        =   0
      Top             =   810
      Width           =   6090
      _ExtentX        =   10732
      _ExtentY        =   2074
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Operation"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Last Run"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Display Note"
         Object.Width           =   7832
      EndProperty
   End
   Begin VB.Menu mnuMax 
      Caption         =   "&Max"
      Begin VB.Menu mnuNewClient 
         Caption         =   "&New Client Window"
      End
      Begin VB.Menu mnuNewScheduleWin 
         Caption         =   "Schedule &Manager"
      End
      Begin VB.Menu mnuActiveAppCache 
         Caption         =   "&Active App Cache"
      End
      Begin VB.Menu mnuDash62 
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
      Begin VB.Menu mnudash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuOperations 
      Caption         =   "O&perations"
      Begin VB.Menu mnuRefresh 
         Caption         =   "Re&fresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnudash14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddOperation 
         Caption         =   "&Add"
      End
      Begin VB.Menu mnuEditOperation 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuDeleteOperation 
         Caption         =   "De&lete"
      End
      Begin VB.Menu mnuDash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMoveUp 
         Caption         =   "Move &Up"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuMoveDown 
         Caption         =   "Move &Down"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuDash36 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRunOperations 
         Caption         =   "&Run Selected"
         Shortcut        =   ^R
         Visible         =   0   'False
      End
      Begin VB.Menu mnuStopOperations 
         Caption         =   "&Stop Selected"
         Shortcut        =   ^S
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Setup"
      Begin VB.Menu mnuPreferences 
         Caption         =   "Setup &Options"
      End
      Begin VB.Menu mnuFileAssoc 
         Caption         =   "&File Associations"
      End
      Begin VB.Menu mnuNetDrives 
         Caption         =   "&Network Drives"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuContents 
         Caption         =   "&Documentation..."
      End
      Begin VB.Menu mnuWebSite 
         Caption         =   "&Neotext.org..."
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmSchOperations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public Sid As Long

Private InOperation As Boolean

Private MyFileName As String

Private Sub LoadGUI()

    Set SchControls.ImageList = frmMain.imgSchedule(0)
    Set SchControls.DisabledImageList = frmMain.imgSchedule(0)
    Set SchControls.HotImageList = frmMain.imgSchedule(1)
    
    Dim btnX As Button
    
    Set btnX = SchControls.Buttons.Add(1, "schedule_add", , 0, "schedule_addout")
    Set btnX = SchControls.Buttons.Add(2, "schedule_edit", , 0, "schedule_editout")
    Set btnX = SchControls.Buttons.Add(3, "schedule_delete", , 0, "schedule_deleteout")
    Set btnX = SchControls.Buttons.Add(4, , , 3)
    If GetCollectSkinValue("toolbarbutton_spacer") = "none" Then btnX.Visible = False
    Set btnX = SchControls.Buttons.Add(5, "schedule_up", , 0, "schedule_upout")
    Set btnX = SchControls.Buttons.Add(6, "schedule_down", , 0, "schedule_downout")

    SetTooltip
    
End Sub
Private Sub UnloadGUI()
    
    SchControls.Buttons.Clear
    Set SchControls.ImageList = Nothing
    Set SchControls.DisabledImageList = Nothing
    Set SchControls.HotImageList = Nothing

End Sub
Public Sub ShowSchedule(ByVal ScheduleID As Long)

    Sid = ScheduleID
    
    If dbSettings.GetScheduleSetting("wState") = -1 Then
        dbSettings.SetScheduleSetting "wState", 0
        dbSettings.SetScheduleSetting "wLeft", ((Screen.Width / 2) - (Me.Width / 2))
        dbSettings.SetScheduleSetting "wTop", ((Screen.Height / 2) - (Me.Height / 2))
        dbSettings.SetScheduleSetting "wWidth", Me.Width
        dbSettings.SetScheduleSetting "wHeight", Me.Height
    Else
       
        Dim frm As Form
        For Each frm In Forms
            If (TypeName(frm) = TypeName(Me)) Then
                If Not (frm.hwnd = Me.hwnd) Then
                    dbSettings.SetScheduleSetting "wLeft", frm.Left + (Screen.TwipsPerPixelX * 32)
                    dbSettings.SetScheduleSetting "wTop", frm.Top + (Screen.TwipsPerPixelY * 32)
                    dbSettings.SetScheduleSetting "wWidth", frm.Width
                    dbSettings.SetScheduleSetting "wHeight", frm.Height
                End If
            End If
        Next
    End If
    
    If ((dbSettings.GetScheduleSetting("wLeft") + dbSettings.GetScheduleSetting("wWidth")) > Screen.Width) Or _
        ((dbSettings.GetScheduleSetting("wTop") + dbSettings.GetScheduleSetting("wHeight")) > Screen.Height) Then
        dbSettings.SetScheduleSetting "wLeft", (32 * Screen.TwipsPerPixelX)
        dbSettings.SetScheduleSetting "wTop", (32 * Screen.TwipsPerPixelY)
    End If
    
    Me.Move dbSettings.GetScheduleSetting("wLeft"), dbSettings.GetScheduleSetting("wTop"), dbSettings.GetScheduleSetting("wWidth"), dbSettings.GetScheduleSetting("wHeight")
    Me.WindowState = dbSettings.GetScheduleSetting("wState")
    
    ListView1.ColumnHeaders(1).Width = dbSettings.GetScheduleSetting("wColumn1")
    ListView1.ColumnHeaders(2).Width = dbSettings.GetScheduleSetting("wColumn2")
    ListView1.ColumnHeaders(3).Width = dbSettings.GetScheduleSetting("wColumn3")
    ListView1.ColumnHeaders(4).Width = dbSettings.GetScheduleSetting("wColumn4")

    LoadGUI
    
    RefreshSchedule
   
    Me.Show

End Sub

Private Sub mnuActiveAppCache_Click()
        frmActiveCache.ShowForm

End Sub

Private Sub dContainer_Resize(Index As Integer)

    On Error Resume Next
    
    Dim xPixel As Integer
    Dim yPixel As Integer
    xPixel = (Screen.TwipsPerPixelX)
    yPixel = (Screen.TwipsPerPixelY)
    
    Line1(Index).X1 = 0
    Line1(Index).X2 = 0
    Line1(Index).Y1 = yPixel
    Line1(Index).Y2 = (dContainer(Index).Height - (2 * yPixel))
    
    Line2(Index).X1 = xPixel
    Line2(Index).X2 = xPixel
    Line2(Index).Y1 = yPixel
    Line2(Index).Y2 = (dContainer(Index).Height - (2 * yPixel))
    
    Line3(Index).X1 = dContainer(Index).Width - yPixel
    Line3(Index).X2 = dContainer(Index).Width - yPixel
    Line3(Index).Y1 = 0
    Line3(Index).Y2 = dContainer(Index).Height - yPixel
    
    Line4(Index).X1 = dContainer(Index).Width - (xPixel * 2)
    Line4(Index).X2 = dContainer(Index).Width - (xPixel * 2)
    Line4(Index).Y1 = yPixel
    Line4(Index).Y2 = (dContainer(Index).Height - (2 * yPixel))
    
    Line5(Index).X1 = 0
    Line5(Index).X2 = dContainer(Index).Width
    Line5(Index).Y1 = 0
    Line5(Index).Y2 = 0
    
    Line6(Index).X1 = xPixel
    Line6(Index).X2 = dContainer(Index).Width
    Line6(Index).Y1 = yPixel
    Line6(Index).Y2 = yPixel
    
    Line7(Index).X1 = xPixel
    Line7(Index).X2 = dContainer(Index).Width - (xPixel * 2)
    Line7(Index).Y1 = dContainer(Index).Height - (yPixel * 2)
    Line7(Index).Y2 = dContainer(Index).Height - (yPixel * 2)
    
    Line8(Index).X1 = 0
    Line8(Index).X2 = dContainer(Index).Width
    Line8(Index).Y1 = dContainer(Index).Height - yPixel
    Line8(Index).Y2 = dContainer(Index).Height - yPixel
    
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Private Sub Form_Activate()
    
    SetIcecue Line4(3), "icecue_shadow"
    SetIcecue Line5(3), "icecue_shadow"
    SetIcecue Line1(3), "icecue_shadow"
    SetIcecue Line7(3), "icecue_shadow"

    SetIcecue Line2(3), "icecue_hilite"
    SetIcecue Line6(3), "icecue_hilite"
    SetIcecue Line3(3), "icecue_hilite"
    SetIcecue Line8(3), "icecue_hilite"
    
    mnuNewScriptWin.Visible = IsScriptIDEInstalled
    mnuRunScript.Visible = IsScriptIDEInstalled
    mnuDash62.Visible = mnuNewScriptWin.Visible

End Sub

Private Sub Form_Load()
    
    Set ListView1.SmallIcons = frmMain.imgOperations(0)
            
    InOperation = False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        If IsAnyOperationWindowOpen Then
            Cancel = (MsgBox("There is one or more operation property windows open, are you sure you want to close this window?", vbYesNo + vbQuestion, AppName) = vbNo)
        End If
    End If
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    If Me.WindowState <> 1 Then
    
        Dim xPixel As Integer
        Dim yPixel As Integer
        xPixel = (Screen.TwipsPerPixelX)
        yPixel = (Screen.TwipsPerPixelY)

        dContainer(3).Move (xPixel * 4), 0, ScaleWidth - (xPixel * 8), ((GetSkinDimension("toolbarbutton_height") + 6) * Screen.TwipsPerPixelY) + (yPixel * 4)
        SchControls.Move yPixel * 2, yPixel * 2, dContainer(3).Width - (yPixel * 4)
    
        ListView1.Move (xPixel * 3), (yPixel * 3) + dContainer(3).Height, ScaleWidth - (6 * xPixel), ScaleHeight - dContainer(3).Height - (7 * yPixel)
    
    End If

    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    UnloadAllOperationWindows
    
    dbSettings.SetScheduleSetting "wState", Me.WindowState
    
    If Me.WindowState = 0 And Me.Visible Then
        dbSettings.SetScheduleSetting "wTop", IIf(Me.Top < 0, (Screen.TwipsPerPixelY * 32), Me.Top)
        dbSettings.SetScheduleSetting "wLeft", IIf(Me.Left < 0, (Screen.TwipsPerPixelX * 32), Me.Left)
        dbSettings.SetScheduleSetting "wHeight", IIf(Me.Height < 0, (Screen.TwipsPerPixelY * 157), Me.Height)
        dbSettings.SetScheduleSetting "wWidth", IIf(Me.Width < 0, (Screen.TwipsPerPixelX * 702), Me.Width)
    End If
    
    dbSettings.SetScheduleSetting "wColumn1", ListView1.ColumnHeaders(1).Width
    dbSettings.SetScheduleSetting "wColumn2", ListView1.ColumnHeaders(2).Width
    dbSettings.SetScheduleSetting "wColumn3", ListView1.ColumnHeaders(3).Width
    dbSettings.SetScheduleSetting "wColumn4", ListView1.ColumnHeaders(4).Width

    UnloadGUI

End Sub

Private Sub ListView1_DblClick()
    mnuEditOperation_Click
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then

        Me.PopupMenu mnuOperations
    End If

End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show

End Sub

Private Sub mnuAddOperation_Click()
    Dim NewOp As New frmSchOpProperties
    NewOp.ShowOperation Me, "Add", Sid

End Sub

Private Sub DisableOperation(ByVal OpId As Long, ByVal setDisable As Boolean)
    Dim dbSchedule As New clsDBSchedule
    dbSchedule.SetOperationValue OpId, "Disabled", setDisable
    Set dbSchedule = Nothing
End Sub

Private Sub mnuContents_Click()
    GotoHelp

End Sub

Private Sub mnuDeleteOperation_Click()
    If ListView1.SelectedItem Is Nothing Then
        MsgBox "You must select the operations you want to delete.", vbInformation, AppName
    Else
        If MsgBox("Are you sure you want to delete the " & GetSelectedCount(ListView1) & " selected operations?", vbQuestion + vbYesNo, AppName) = vbYes Then
            Dim cnt As Integer
            cnt = 1
            Do
        
                If ListView1.ListItems(cnt).Selected Then
                    UnloadOperationWindow CLng(ListView1.ListItems(cnt).Tag)
                    RemoveOperation CLng(ListView1.ListItems(cnt).Tag)
                End If
                cnt = cnt + 1
            Loop Until cnt > ListView1.ListItems.Count
            
            RefreshSchedule
        End If
    End If

    If Not (ProcessRunning(ServiceFileName) = 0) Then
        MessageQueueAdd ServiceFileName, "/loadschedule " & Sid
    End If

End Sub

Private Sub mnuEditOperation_Click()
    If ListView1.SelectedItem Is Nothing Then
        MsgBox "You must select the operations you want to edit.", vbInformation, AppName
    Else
        Dim NewOp As frmSchOpProperties
        Dim cnt As Integer
        cnt = 1
        Do
        
            If ListView1.ListItems(cnt).Selected Then

                Set NewOp = IsOperationWindowOpen(CLng(ListView1.ListItems(cnt).Tag))
                If NewOp Is Nothing Then
                    Set NewOp = New frmSchOpProperties
                    NewOp.ShowOperation Me, "Edit", Sid, CLng(ListView1.ListItems(cnt).Tag)
                Else
                    NewOp.Show
                End If
                Set NewOp = Nothing
    
            End If
            cnt = cnt + 1

        Loop Until cnt > ListView1.ListItems.Count
    End If

End Sub

Private Sub mnuNewScriptWin_Click()
    RunProcess AppPath & MaxIDEFileName, "", vbNormalFocus, False
End Sub

Private Sub mnuRefresh_Click()
    RefreshSchedule
    
End Sub

Private Sub mnuClose_Click()

    Unload Me

End Sub

Private Sub mnuFileAssoc_Click()
    frmFileAssoc.Show
End Sub



Private Sub mnuMoveDown_Click()
    If GetSelectedCount(ListView1) = 1 Then
    
        Dim DownIndex As Integer
        DownIndex = ListView1.SelectedItem.Index
        If DownIndex < ListView1.ListItems.Count Then
            Dim swapStr2 As String
        
            swapStr2 = ListView1.ListItems(DownIndex).Text
            ListView1.ListItems(DownIndex).Text = ListView1.ListItems(DownIndex + 1).Text
            ListView1.ListItems(DownIndex + 1).Text = swapStr2
        
            swapStr2 = ListView1.ListItems(DownIndex).Tag
            ListView1.ListItems(DownIndex).Tag = ListView1.ListItems(DownIndex + 1).Tag
            ListView1.ListItems(DownIndex + 1).Tag = swapStr2
        
            swapStr2 = ListView1.ListItems(DownIndex).SubItems(1)
            ListView1.ListItems(DownIndex).SubItems(1) = ListView1.ListItems(DownIndex + 1).SubItems(1)
            ListView1.ListItems(DownIndex + 1).SubItems(1) = swapStr2
        
            swapStr2 = ListView1.ListItems(DownIndex).SubItems(2)
            ListView1.ListItems(DownIndex).SubItems(2) = ListView1.ListItems(DownIndex + 1).SubItems(2)
            ListView1.ListItems(DownIndex + 1).SubItems(2) = swapStr2
            
            swapStr2 = ListView1.ListItems(DownIndex).SubItems(3)
            ListView1.ListItems(DownIndex).SubItems(3) = ListView1.ListItems(DownIndex + 1).SubItems(3)
            ListView1.ListItems(DownIndex + 1).SubItems(3) = swapStr2
            
            ListView1.ListItems(DownIndex).Selected = False
            ListView1.ListItems(DownIndex + 1).Selected = True
            ListView1.ListItems(DownIndex + 1).EnsureVisible
            
            Dim dbSchedule As New clsDBSchedule
            dbSchedule.SetOperationValue CLng(ListView1.ListItems(DownIndex).Tag), "OperationOrder", DownIndex
            dbSchedule.SetOperationValue CLng(ListView1.ListItems(DownIndex + 1).Tag), "OperationOrder", DownIndex + 1
            Set dbSchedule = Nothing
        
            If Not (ProcessRunning(ServiceFileName) = 0) Then
                MessageQueueAdd ServiceFileName, "/loadschedule " & Sid
            End If
        
        End If
        
    Else
        If GetSelectedCount(ListView1) = 0 Then
            MsgBox "Please select an operation to move down.", vbInformation, AppName
        Else
            MsgBox "You can only select one operation to move down.", vbInformation, AppName
        End If
    End If
End Sub

Private Sub mnuMoveUp_Click()
    If GetSelectedCount(ListView1) = 1 Then
    
        Dim UpIndex As Integer
        UpIndex = ListView1.SelectedItem.Index
        If UpIndex > 1 Then
            Dim swapStr1 As String
        
            swapStr1 = ListView1.ListItems(UpIndex).Text
            ListView1.ListItems(UpIndex).Text = ListView1.ListItems(UpIndex - 1).Text
            ListView1.ListItems(UpIndex - 1).Text = swapStr1
        
            swapStr1 = ListView1.ListItems(UpIndex).Tag
            ListView1.ListItems(UpIndex).Tag = ListView1.ListItems(UpIndex - 1).Tag
            ListView1.ListItems(UpIndex - 1).Tag = swapStr1
        
            swapStr1 = ListView1.ListItems(UpIndex).SubItems(1)
            ListView1.ListItems(UpIndex).SubItems(1) = ListView1.ListItems(UpIndex - 1).SubItems(1)
            ListView1.ListItems(UpIndex - 1).SubItems(1) = swapStr1
        
            swapStr1 = ListView1.ListItems(UpIndex).SubItems(2)
            ListView1.ListItems(UpIndex).SubItems(2) = ListView1.ListItems(UpIndex - 1).SubItems(2)
            ListView1.ListItems(UpIndex - 1).SubItems(2) = swapStr1
            
            swapStr1 = ListView1.ListItems(UpIndex).SubItems(3)
            ListView1.ListItems(UpIndex).SubItems(3) = ListView1.ListItems(UpIndex - 1).SubItems(3)
            ListView1.ListItems(UpIndex - 1).SubItems(3) = swapStr1
            
            ListView1.ListItems(UpIndex).Selected = False
            ListView1.ListItems(UpIndex - 1).Selected = True
            ListView1.ListItems(UpIndex - 1).EnsureVisible
        
            Dim dbSchedule As New clsDBSchedule
            dbSchedule.SetOperationValue CLng(ListView1.ListItems(UpIndex).Tag), "OperationOrder", UpIndex
            dbSchedule.SetOperationValue CLng(ListView1.ListItems(UpIndex - 1).Tag), "OperationOrder", UpIndex - 1
            Set dbSchedule = Nothing
        
            If Not (ProcessRunning(ServiceFileName) = 0) Then
                MessageQueueAdd ServiceFileName, "/loadschedule " & Sid
            End If
        
        End If
    
    Else
        If GetSelectedCount(ListView1) = 0 Then
            MsgBox "Please select an operation to move up.", vbInformation, AppName
        Else
            MsgBox "You can only select one operation to move up.", vbInformation, AppName
        End If
    End If
End Sub

Private Sub mnuNetDrives_Click()
    frmNetDrives.Show
End Sub

Private Sub mnuPreferences_Click()
    frmSetup.Show

End Sub

Private Sub mnuNewClient_Click()
    Dim newClient As New frmFTPClientGUI
    newClient.LoadClient
    newClient.ShowClient
End Sub

Private Sub mnuNewScheduleWin_Click()
        frmSchManager.Show
End Sub

Private Function EqualTime(ByVal Time1, ByVal Time2) As Boolean
    EqualTime = (Hour(Time1) = Hour(Time2) And Minute(Time1) = Minute(Time2))
End Function

Private Function EqualDate(ByVal Date1, ByVal Date2) As Boolean
    EqualDate = (Month(Date1) = Month(Date2) And Day(Date1) = Day(Date2) And Year(Date1) = Year(Date2))
End Function



Private Sub mnuRunOperations_Click()

    If (ProcessRunning(ServiceFileName) = 0) Then
        MsgBox "The schedule service is not running, please start it from the schedule manager.", vbInformation, AppName
    Else

        If MsgBox("Are you sure you want to run the selected operations?", vbQuestion + vbYesNo, AppName) = vbYes Then
            Dim cnt As Integer
            Dim cmds As String
            For cnt = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(cnt).Selected Then cmds = cmds & "/runoperation " & Sid & " " & ListView1.ListItems(cnt).Tag
            Next
            MessageQueueAdd ServiceFileName, cmds
        End If

    End If

End Sub

Private Sub mnuRunScript_Click()
    frmMain.RunScript
End Sub

Private Sub mnuStopOperations_Click()

    If (ProcessRunning(ServiceFileName) = 0) Then
        MsgBox "The schedule service is not running, please start it from the schedule manager.", vbInformation, AppName
    Else

        If PromptAbortClose("Are you sure you want to stop selected operations?") Then
            Dim cnt As Integer
            Dim cmds As String
            For cnt = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(cnt).Selected Then cmds = cmds & "/stopoperation " & Sid & " " & ListView1.ListItems(cnt).Tag
            Next
            MessageQueueAdd ServiceFileName, cmds
        End If

    End If

End Sub

Private Sub mnuWebSite_Click()
    RunFile AppPath & "Neotext.org.url"

End Sub

Private Sub SchControls_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Key
        Case "schedule_add"
            mnuAddOperation_Click
        Case "schedule_edit"
            mnuEditOperation_Click
        Case "schedule_delete"
            mnuDeleteOperation_Click
        
        Case "schedule_up"
            mnuMoveUp_Click
            
        Case "schedule_down"
            mnuMoveDown_Click
        
    End Select
End Sub

Public Sub RefreshSchedule()
    Dim nodX As ListItem
    
    Dim dbConn As New clsDBConnection
    
    Dim rsOperation As New ADODB.Recordset

    dbConn.rsQuery rsOperation, "SELECT * FROM Schedules WHERE ID=" & Sid & ";"
    
    Me.Caption = "[" & rsOperation("ScheduleName") & "] - Operations"

    Dim selItem As Long
    If Not (ListView1.SelectedItem Is Nothing) Then
        selItem = ListView1.SelectedItem.Index
    End If
    
    Dim cnt As Long
    
    ListView1.ListItems.Clear
    
    dbConn.rsQuery rsOperation, "SELECT * FROM Operations WHERE ParentID=" & Sid & " ORDER BY OperationOrder;"
    
    If Not rsOperation.EOF And Not rsOperation.BOF Then
    
        rsOperation.MoveFirst
        Do
            If Len(CStr(rsOperation("ID"))) > cnt Then cnt = Len(CStr(rsOperation("ID")))
            
            Set nodX = ListView1.ListItems.Add(, , rsOperation("ID"))
            nodX.Tag = rsOperation("ID")
            nodX.SmallIcon = CStr(rsOperation("Status") & "")
            
            nodX.SubItems(1) = rsOperation("Action") & ""
            nodX.SubItems(2) = rsOperation("LastRun") & ""
            nodX.SubItems(3) = rsOperation("OperationName") & ""
           
            nodX.Selected = (selItem = nodX.Index) Or ((selItem = 0) And (nodX.Index = 1))
            
            rsOperation.MoveNext
        Loop Until rsOperation.EOF Or rsOperation.BOF
    
        For Each nodX In ListView1.ListItems
            nodX.Text = String(cnt - Len(Trim(nodX.Text)), "0") & nodX.Text
        Next
    
    End If

    If rsOperation.State <> 0 Then rsOperation.Close
    Set rsOperation = Nothing
    
    Set dbConn = Nothing

End Sub

Public Function IsAnyOperationWindowOpen() As Boolean
    Dim frm
    For Each frm In Forms
        If TypeName(frm) = "frmSchOpProperties" Then
            If frm.SchId = Sid Then
                IsAnyOperationWindowOpen = True
                Exit Function
            End If
        End If
    Next
    IsAnyOperationWindowOpen = False
End Function

Public Function IsOperationWindowOpen(ByVal OpId As Long) As frmSchOpProperties
    Dim frm
    For Each frm In Forms
        If TypeName(frm) = "frmSchOpProperties" Then
            If frm.SchId = Sid And frm.OpId = OpId Then
                Set IsOperationWindowOpen = frm
            End If
        End If
    Next
End Function
Public Function UnloadAllOperationWindows()
    Dim frm
    For Each frm In Forms
        If TypeName(frm) = "frmSchOpProperties" Then
            If frm.SchId = Sid Then
                Unload frm
            End If
        End If
    Next
End Function
Public Function UnloadOperationWindow(ByVal OpId As Long)
    Dim frm
    For Each frm In Forms
        If TypeName(frm) = "frmSchOpProperties" Then
            If frm.SchId = Sid And frm.OpId = OpId Then
                Unload frm
            End If
        End If
    Next
End Function
Public Sub UpdateStatus(ByVal OpId As String)
    Dim dbs
    Dim Item As ListItem
    For Each Item In ListView1.ListItems
        If CStr(Item.Tag) = OpId Then
            Dim dbSchedule As New clsDBSchedule
            Item.SmallIcon = LCase(dbSchedule.GetOperationValue(OpId, "Status"))
            Item.SubItems(2) = dbSchedule.GetOperationValue(OpId, "LastRun")
            Set dbSchedule = Nothing
        End If
    Next
End Sub

Public Sub SetTooltip()
    With Me
        If dbSettings.GetProfileSetting("ViewToolTips") Then
            .SchControls.Buttons(1).ToolTipText = "Adds a new schedule operation."
            .SchControls.Buttons(2).ToolTipText = "Edits the selected schedule operation."
            .SchControls.Buttons(3).ToolTipText = "Deletes the selected schedule operation."
            
            .SchControls.Buttons(5).ToolTipText = "Moves the selected operation up the list."
            .SchControls.Buttons(6).ToolTipText = "Moves the selected operation down the list."

            
            .ListView1.ToolTipText = "Displays all the schedules operations."
            
        Else
            .SchControls.Buttons(1).ToolTipText = ""
            .SchControls.Buttons(2).ToolTipText = ""
            .SchControls.Buttons(3).ToolTipText = ""
            
            .SchControls.Buttons(5).ToolTipText = ""
            .SchControls.Buttons(6).ToolTipText = ""

            
            .ListView1.ToolTipText = ""
            
        End If
    End With
End Sub
