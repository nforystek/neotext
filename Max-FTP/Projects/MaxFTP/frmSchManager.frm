


VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSchManager 
   Caption         =   "Manager"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7500
   Icon            =   "frmSchManager.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   7500
   Begin VB.PictureBox dContainer 
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   3
      Left            =   420
      ScaleHeight     =   735
      ScaleWidth      =   5190
      TabIndex        =   2
      Top             =   600
      Width           =   5190
      Begin MSComctlLib.Toolbar SchControls 
         Height          =   450
         Left            =   2355
         TabIndex        =   0
         Top             =   150
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   794
         ButtonWidth     =   826
         ButtonHeight    =   804
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
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
      Begin VB.Line Line3 
         BorderColor     =   &H80000014&
         Index           =   3
         Tag             =   "highlight"
         X1              =   2025
         X2              =   2025
         Y1              =   90
         Y2              =   660
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2205
      Left            =   150
      TabIndex        =   1
      Top             =   345
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3889
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Owner"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Public"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Type"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ID"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5550
      Top             =   1620
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchManager.frx":08CA
            Key             =   "schedule"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuMax 
      Caption         =   "&Max"
      Begin VB.Menu mnuNewClientWindow 
         Caption         =   "&New Client Window"
      End
      Begin VB.Menu mnuActiveAppCache 
         Caption         =   "&Active App Cache"
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuScriptIDE 
         Caption         =   "&Scripting IDE"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRunScript 
         Caption         =   "&Run Script"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDash43 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuSchedule 
      Caption         =   "S&chedule"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuDash7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddSchedule 
         Caption         =   "&Add"
      End
      Begin VB.Menu mnuRemoveSchedule 
         Caption         =   "&Remove"
      End
      Begin VB.Menu mnuDash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Re&fresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuServiceStatus 
         Caption         =   "&Stop Service"
      End
      Begin VB.Menu mnuDash4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "&Properties"
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Begin VB.Menu mnuRunSchedule 
         Caption         =   "&Run Selected"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuStopSchedule 
         Caption         =   "&Stop Selected"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "&Setup"
      Begin VB.Menu mnuOptions 
         Caption         =   "Setup &Options"
      End
      Begin VB.Menu mnuFileAssociations 
         Caption         =   "&File Associations"
      End
      Begin VB.Menu mnuNetworkDrives 
         Caption         =   "&Network Drives"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuContents 
         Caption         =   "&Documentation..."
      End
      Begin VB.Menu mnuNeoTextWebSite 
         Caption         =   "&Neotext.org..."
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmSchManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private WithEvents timServiceStatus As NTSchedule20.Timer
Attribute timServiceStatus.VB_VarHelpID = -1

Private Sub mnuActiveAppCache_Click()
    frmActiveCache.ShowForm
End Sub

Private Function GetScheduleTypeText(ByVal SchType As Long) As String
    Select Case SchType
        Case 0
            GetScheduleTypeText = "Manual Execute"
        Case 2
            GetScheduleTypeText = "Set Date/Time"
        Case 1
            GetScheduleTypeText = "Increment Run"
    End Select
End Function
Private Sub LoadGUI()
    
    Set SchControls.ImageList = frmMain.imgSchedule(0)
    Set SchControls.DisabledImageList = frmMain.imgSchedule(0)
    Set SchControls.HotImageList = frmMain.imgSchedule(1)
    
    Dim btnX As Button
    
    Set btnX = SchControls.Buttons.Add(1, "schedule_servicestop", , 0, "schedule_servicestopout")

    Set btnX = SchControls.Buttons.Add(2, "schedule_servicestart", , 0, "schedule_servicestartout")
    
    Set btnX = SchControls.Buttons.Add(3, , , 3)
    If GetCollectSkinValue("toolbarbutton_spacer") = "none" Then btnX.Visible = False
    
    Set btnX = SchControls.Buttons.Add(4, "schedule_add", , 0, "schedule_addout")
    Set btnX = SchControls.Buttons.Add(5, "schedule_edit", , 0, "schedule_editout")
    Set btnX = SchControls.Buttons.Add(6, "schedule_delete", , 0, "schedule_deleteout")
    Set btnX = SchControls.Buttons.Add(7, , , 3)
    If GetCollectSkinValue("toolbarbutton_spacer") = "none" Then btnX.Visible = False

    Set btnX = SchControls.Buttons.Add(8, "schedule_run", , 0, "schedule_runout")
    Set btnX = SchControls.Buttons.Add(9, "schedule_stop", , 0, "schedule_stopout")

    SetTooltip
    
End Sub
Private Sub UnloadGUI()
    
    SchControls.Buttons.Clear
    Set SchControls.ImageList = Nothing
    Set SchControls.DisabledImageList = Nothing
    Set SchControls.HotImageList = Nothing

End Sub

Public Sub RefreshSchedules()

    Dim dbConn As New clsDBConnection
    Dim rsSchedule As New ADODB.Recordset
    Dim rsUser As New ADODB.Recordset
    Dim dbSchedule As New clsDBSchedule
    Dim newItem As ListItem
       
    Dim selItem As Long
    If Not (ListView1.SelectedItem Is Nothing) Then
        selItem = ListView1.SelectedItem.Index
    End If
    
    Dim cnt As Long
    
    ListView1.ListItems.Clear
    
    dbConn.rsQuery rsSchedule, "SELECT * FROM Schedules WHERE ParentID=" & dbSettings.CurrentUserID & " OR IsPublic=Yes ORDER BY ScheduleName;"
    
    If Not rsSchedule.EOF And Not rsSchedule.BOF Then
    
        rsSchedule.MoveFirst
        Do
            If Len(CStr(rsSchedule("ID"))) > cnt Then cnt = Len(CStr(rsSchedule("ID")))
        
            If Not (rsSchedule("ScheduleName") = "") Then
            
                Set newItem = ListView1.ListItems.Add(, "s" & rsSchedule("ID"), rsSchedule("ScheduleName"), , "schedule")
            
                dbConn.rsQuery rsUser, "SELECT * FROM Users WHERE ID=" & rsSchedule("ParentID") & ";"
                newItem.Tag = rsSchedule("ID")
                newItem.SubItems(1) = rsUser("UserName")
                newItem.SubItems(2) = rsSchedule("IsPublic")
                newItem.SubItems(3) = GetScheduleTypeText(rsSchedule("ScheduleType"))
                newItem.SubItems(4) = rsSchedule("ID")
                
            End If

            rsSchedule.MoveNext
        Loop Until rsSchedule.EOF Or rsSchedule.BOF
    
        For Each newItem In ListView1.ListItems
            newItem.SubItems(4) = String(cnt - Len(Trim(newItem.SubItems(4))), "0") & newItem.SubItems(4)
            newItem.Selected = False
        Next
        
        If selItem > 0 Then ListView1.SelectedItem = ListView1.ListItems(selItem)
        
    End If
    
    Set dbSchedule = Nothing
    
    If rsUser.State <> 0 Then rsUser.Close
    Set rsUser = Nothing
    
    If rsSchedule.State <> 0 Then rsSchedule.Close
    Set rsSchedule = Nothing
    
    Set dbConn = Nothing

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

    SetIcecue Line2(3), "icecue_hilite"
    SetIcecue Line8(3), "icecue_hilite"
    SetIcecue Line3(3), "icecue_hilite"
    SetIcecue Line6(3), "icecue_hilite"
    
    SetIcecue Line1(3), "icecue_shadow"
    SetIcecue Line5(3), "icecue_shadow"
    SetIcecue Line4(3), "icecue_shadow"
    SetIcecue Line7(3), "icecue_shadow"
    
    mnuScriptIDE.Visible = IsScriptIDEInstalled
    mnuDash1.Visible = mnuScriptIDE.Visible

End Sub

Private Sub Form_Load()
    
    mnuScriptIDE.Visible = IsScriptIDEInstalled
    mnuRunScript.Visible = IsScriptIDEInstalled
    mnuDash1.Visible = mnuScriptIDE.Visible
    
    If dbSettings.GetScheduleSetting("mState") = -1 Then
        dbSettings.SetScheduleSetting "mState", 0
        dbSettings.SetScheduleSetting "mLeft", ((Screen.Width / 2) - (Me.Width / 2))
        dbSettings.SetScheduleSetting "mTop", ((Screen.Height / 2) - (Me.Height / 2))
        dbSettings.SetScheduleSetting "mWidth", Me.Width
        dbSettings.SetScheduleSetting "mHeight", Me.Height
    End If
    
    If ((dbSettings.GetScheduleSetting("mLeft") + dbSettings.GetScheduleSetting("mWidth")) > Screen.Width) Or _
        ((dbSettings.GetScheduleSetting("mTop") + dbSettings.GetScheduleSetting("mHeight")) > Screen.Height) Then
        dbSettings.SetScheduleSetting "mLeft", (32 * Screen.TwipsPerPixelX)
        dbSettings.SetScheduleSetting "mTop", (32 * Screen.TwipsPerPixelY)
    End If
    
    Me.Move dbSettings.GetScheduleSetting("mLeft"), dbSettings.GetScheduleSetting("mTop"), dbSettings.GetScheduleSetting("mWidth"), dbSettings.GetScheduleSetting("mHeight")
    Me.WindowState = dbSettings.GetScheduleSetting("mState")
    
    LoadGUI
    
    Set timServiceStatus = New NTSchedule20.Timer
    timServiceStatus.Interval = 2000
    timServiceStatus.enabled = True
    
    SetServiceStatus
    
    ListView1.ColumnHeaders(1).Width = dbSettings.GetScheduleSetting("mColumn1")
    ListView1.ColumnHeaders(2).Width = dbSettings.GetScheduleSetting("mColumn2")
    ListView1.ColumnHeaders(3).Width = dbSettings.GetScheduleSetting("mColumn3")
    ListView1.ColumnHeaders(4).Width = dbSettings.GetScheduleSetting("mColumn4")
    ListView1.ColumnHeaders(5).Width = dbSettings.GetScheduleSetting("mColumn5")
    
    ListView1.SortKey = dbSettings.GetScheduleSetting("mColumnKey")
    ListView1.SortOrder = dbSettings.GetScheduleSetting("mColumnSort")
    
    RefreshSchedules

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

Private Sub SetMenusEnabled()
    mnuOpen.enabled = (Not ListView1.SelectedItem Is Nothing)
    mnuRemoveSchedule.enabled = (Not ListView1.SelectedItem Is Nothing)
    mnuProperties.enabled = (Not ListView1.SelectedItem Is Nothing)
    
    SetServiceStatus
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    dbSettings.SetScheduleSetting "mState", Me.WindowState
    
    If Me.WindowState = 0 And Me.Visible Then
        dbSettings.SetScheduleSetting "mTop", IIf(Me.Top < 0, (Screen.TwipsPerPixelY * 32), Me.Top)
        dbSettings.SetScheduleSetting "mLeft", IIf(Me.Left < 0, (Screen.TwipsPerPixelX * 32), Me.Left)
        dbSettings.SetScheduleSetting "mHeight", IIf(Me.Height < 0, (Screen.TwipsPerPixelY * 157), Me.Height)
        dbSettings.SetScheduleSetting "mWidth", IIf(Me.Width < 0, (Screen.TwipsPerPixelX * 702), Me.Width)
    End If
    
    dbSettings.SetScheduleSetting "mColumn1", ListView1.ColumnHeaders(1).Width
    dbSettings.SetScheduleSetting "mColumn2", ListView1.ColumnHeaders(2).Width
    dbSettings.SetScheduleSetting "mColumn3", ListView1.ColumnHeaders(3).Width
    dbSettings.SetScheduleSetting "mColumn4", ListView1.ColumnHeaders(4).Width
    dbSettings.SetScheduleSetting "mColumn5", ListView1.ColumnHeaders(5).Width
    
    timServiceStatus.enabled = False
    Set timServiceStatus = Nothing
    
    UnloadGUI
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ColumnSortClick ListView1, ColumnHeader
    
    dbSettings.SetScheduleSetting "mColumnKey", ListView1.SortKey
    dbSettings.SetScheduleSetting "mColumnSort", ListView1.SortOrder
End Sub

Private Sub ListView1_DblClick()
    mnuOpen_Click
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set ListView1.SelectedItem = Item
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        SetMenusEnabled
        Me.PopupMenu mnuSchedule
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub


Private Sub mnuRunScript_Click()
    frmMain.RunScript
End Sub

Private Sub mnuSchedule_Click()
    SetMenusEnabled
End Sub

Private Sub mnuAddSchedule_Click()
    AddNewSchedule
End Sub

Private Sub OpenSchedule(ByVal ScheduleID As Long)
    
    Dim newSchedule As frmSchOperations
    Set newSchedule = IsScheduleWindowOpen(ScheduleID)
    If newSchedule Is Nothing Then
        Set newSchedule = New frmSchOperations
        newSchedule.ShowSchedule ScheduleID
    Else
        newSchedule.Show
    End If
    If (newSchedule.WindowState = 1) Then newSchedule.WindowState = 0
    
End Sub

Private Sub mnuClose_Click()
    Unload Me

End Sub

Private Sub mnuContents_Click()
    GotoHelp

End Sub

Private Sub mnuFileAssociations_Click()
    frmFileAssoc.Show
End Sub


Private Sub mnuNeoTextWebSite_Click()
    RunFile AppPath & "Neotext.org.url"
End Sub

Private Sub mnuNetworkDrives_Click()
        frmNetDrives.Show
End Sub

Private Sub mnuNewClientWindow_Click()
    Dim newWin As New frmFTPClientGUI
    newWin.LoadClient
    newWin.ShowClient
End Sub

Private Sub mnuOpen_Click()
    Dim cnt As Integer
    cnt = 1
    If ListView1.ListItems.Count >= 1 Then
        Do
        
            If ListView1.ListItems(cnt).Selected Then
            
                If Not IsSchedulePropertiesOpen(CLng(Mid(ListView1.ListItems(cnt).Key, 2))) Is Nothing Then
                    MsgBox "Please close schedule [" & ListView1.ListItems(cnt).Text & "] before editing its properties.", vbInformation, AppName
                Else
                    OpenSchedule CLng(Mid(ListView1.ListItems(cnt).Key, 2))
                End If
            End If
            cnt = cnt + 1

        Loop Until cnt > ListView1.ListItems.Count
    End If

End Sub

Private Sub mnuOptions_Click()
    frmSetup.Show
End Sub

Private Sub mnuProperties_Click()

    Dim cnt As Integer
    cnt = 1
    If ListView1.ListItems.Count >= 1 Then
        Do
    
            If ListView1.ListItems(cnt).Selected Then
                If Not IsScheduleWindowOpen(CLng(Mid(ListView1.ListItems(cnt).Key, 2))) Is Nothing Then
                    MsgBox "Please close schedule [" & ListView1.ListItems(cnt).Text & "] before editing its properties.", vbInformation, AppName
                Else
                    Dim newProp As New frmSchProperties
                    newProp.ShowProperties CLng(Mid(ListView1.ListItems(cnt).Key, 2)), False
                End If
            End If
            cnt = cnt + 1

        Loop Until cnt > ListView1.ListItems.Count
    End If
    
End Sub

Private Sub mnuRefresh_Click()
    RefreshSchedules
End Sub

Private Sub AddNewSchedule()
    Dim newSchedule As ScheduleType

    newSchedule.ParentID = dbSettings.CurrentUserID
    newSchedule.IsPublic = False

    Dim ID As Long
    ID = AddSchedule(newSchedule)
    newSchedule.ScheduleName = ""

    Dim newProp As New frmSchProperties
    newProp.ShowProperties ID, True
End Sub

Private Sub mnuRemoveSchedule_Click()
    
    If ListView1.SelectedItem Is Nothing Then
        MsgBox "You must select the schedule you want to remove.", vbInformation, AppName
    Else
    
        If MsgBox("Are you sure you want to remove the " & GetSelectedCount(ListView1) & " selected schedule(s)?", vbQuestion + vbYesNo, AppName) = vbYes Then

            Dim cnt As Integer
            Dim Sid As Long
            
            cnt = 1
            If ListView1.ListItems.Count >= 1 Then
                Do
                
                    If ListView1.ListItems(cnt).Selected Then
                    
                        Sid = CLng(Mid(ListView1.ListItems(cnt).Key, 2))
                        RemoveSchedule Sid
                        UnloadScheduleWindow Sid
                        UnloadPropertiesWindow Sid
                        ListView1.ListItems.Remove cnt
    
                        If Not (ProcessRunning(ServiceFileName) = 0) Then
                            MessageQueueAdd ServiceFileName, "/unloadschedule " & Sid
                        End If
                        
                    Else
                        cnt = cnt + 1
                    End If

                Loop Until cnt > ListView1.ListItems.Count
            End If
        
        End If

    End If

End Sub

Private Sub mnuScriptIDE_Click()
    RunProcess AppPath & MaxIDEFileName, "", vbNormalFocus, False
End Sub

Public Function IsSchedulePropertiesOpen(ByVal SchId As Long) As frmSchProperties
    Dim frm
    For Each frm In Forms
        If TypeName(frm) = "frmSchProperties" Then
            If frm.ID = SchId Then
                Set IsSchedulePropertiesOpen = frm
            End If
        End If
    Next
End Function

Public Function IsScheduleWindowOpen(ByVal SchId As Long) As frmSchOperations
    Dim frm
    For Each frm In Forms
        If TypeName(frm) = "frmSchOperations" Then
            If frm.Sid = SchId Then
                Set IsScheduleWindowOpen = frm
            End If
        End If
    Next
End Function

Public Function UnloadScheduleWindow(ByVal SchId As Long)
    Dim frm
    For Each frm In Forms
        If TypeName(frm) = "frmSchOperations" Then
            If frm.Sid = SchId Then
                Unload frm
            End If
        End If
    Next
End Function
Public Function UnloadPropertiesWindow(ByVal SchId As Long)
    Dim frm
    For Each frm In Forms
        If TypeName(frm) = "frmSchProperties" Then
            If frm.ID = SchId Then
                Unload frm
            End If
        End If
    Next
End Function

Private Sub SchControls_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Key
        Case "schedule_servicestart"
            If (dbSettings.CurrentUserAccessRights = ar_Administrator) Or ((Not (dbSettings.CurrentUserAccessRights = ar_Administrator)) And dbSettings.GetPublicSetting("ServiceAllowAny")) Then
                RunProcess AppPath & MaxUtilityFileName, "start " & MaxServiceName, vbHide, False
            End If
        Case "schedule_servicestop"
            If (dbSettings.CurrentUserAccessRights = ar_Administrator) Or ((Not (dbSettings.CurrentUserAccessRights = ar_Administrator)) And dbSettings.GetPublicSetting("ServiceAllowAny")) Then
                RunProcess AppPath & MaxUtilityFileName, "stop " & MaxServiceName, vbHide, False
            End If
        Case "schedule_add"
            mnuAddSchedule_Click
        Case "schedule_edit"
            mnuOpen_Click
        Case "schedule_delete"
            mnuRemoveSchedule_Click
        Case "schedule_run"
            mnuRunSchedule_Click
            
        Case "schedule_stop"
            mnuStopSchedule_Click
            
    End Select

End Sub

Private Sub SetServiceStatus()
    If (dbSettings.CurrentUserAccessRights = ar_Administrator) Or ((Not (dbSettings.CurrentUserAccessRights = ar_Administrator)) And dbSettings.GetPublicSetting("ServiceAllowAny")) Then
    
        If (Not (ProcessRunning(ServiceFileName) = 0)) Or IsDebugger("MaxService") Then
    
            SchControls.Buttons("schedule_servicestart").Visible = False
            SchControls.Buttons("schedule_servicestop").Visible = True
            Me.Caption = "Manager - (Service Running)"
        Else
    
            SchControls.Buttons("schedule_servicestart").Visible = True
            SchControls.Buttons("schedule_servicestop").Visible = False
            Me.Caption = "Manager - (Service Stopped)"
        End If
    Else
        If (Not (ProcessRunning(ServiceFileName) = 0)) Or IsDebugger("MaxService") Then
    
            SchControls.Buttons("schedule_servicestart").Visible = False
            SchControls.Buttons("schedule_servicestop").Visible = False
            Me.Caption = "Manager - (Service Running)"
        Else
    
            SchControls.Buttons("schedule_servicestart").Visible = False
            SchControls.Buttons("schedule_servicestop").Visible = False
            Me.Caption = "Manager - (Service Stopped)"
        End If
    End If
End Sub

Private Sub timServiceStatus_OnTicking()
    SetServiceStatus
End Sub

Private Sub mnuRunSchedule_Click()

    If (ProcessRunning(ServiceFileName) = 0) And Not IsDebugger Then
        MsgBox "The schedule service is not running, please start it from the schedule manager.", vbInformation, AppName
    Else

        If MsgBox("Are you sure you want to run the selected schedules?", vbQuestion + vbYesNo, AppName) = vbYes Then
            Dim cnt As Integer
            Dim cmds As String
            For cnt = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(cnt).Selected Then cmds = cmds & "/runschedule " & ListView1.ListItems(cnt).Tag
            Next
            MessageQueueAdd ServiceFileName, cmds
        End If

    End If

End Sub

Private Sub mnuStopSchedule_Click()

    If (ProcessRunning(ServiceFileName) = 0) And Not IsDebugger("MaxService") Then
        MsgBox "The schedule service is not running, please start it from the schedule manager.", vbInformation, AppName
    Else

        If PromptAbortClose("Are you sure you want to stop selected schedules?") Then
            Dim cnt As Integer
            Dim cmds As String
            For cnt = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(cnt).Selected Then cmds = cmds & "/stopschedule " & ListView1.ListItems(cnt).Tag
            Next
            MessageQueueAdd ServiceFileName, cmds
        End If

    End If

End Sub

Public Sub SetTooltip()
    With Me
        If dbSettings.GetProfileSetting("ViewToolTips") Then
            .SchControls.Buttons(1).ToolTipText = "Stops the schedule service."
            .SchControls.Buttons(2).ToolTipText = "Starts the schedule service."
            
            .SchControls.Buttons(4).ToolTipText = "Adds a new schedule."
            .SchControls.Buttons(5).ToolTipText = "Opens the selected schedule."
            .SchControls.Buttons(6).ToolTipText = "Deletes the selected schedule."
            
            .SchControls.Buttons(8).ToolTipText = "Runs selected operations in order."
            .SchControls.Buttons(9).ToolTipText = "Stops the selected operation."
            
            .ListView1.ToolTipText = "Displays all the schedules."
            
        Else
            .SchControls.Buttons(1).ToolTipText = ""
            .SchControls.Buttons(2).ToolTipText = ""
            
            .SchControls.Buttons(4).ToolTipText = ""
            .SchControls.Buttons(5).ToolTipText = ""
            .SchControls.Buttons(6).ToolTipText = ""
            
            .SchControls.Buttons(8).ToolTipText = ""
            .SchControls.Buttons(9).ToolTipText = ""
                    
            .ListView1.ToolTipText = ""
            
        End If
    End With
End Sub