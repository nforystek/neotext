VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmGlobalEvents 
   Caption         =   "Global Events"
   ClientHeight    =   3645
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8625
   Icon            =   "frmGlobalEvents.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   8625
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3120
      Top             =   2910
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ListView lstErrors 
      Height          =   2250
      Left            =   240
      TabIndex        =   0
      Top             =   390
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   3969
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Source"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Location"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Action"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Message"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Date/Time"
         Object.Width           =   2999
      EndProperty
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   3315
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   15
      X2              =   4275
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuEvents 
      Caption         =   "&Events"
      Begin VB.Menu mnuExport 
         Caption         =   "&Export"
         Begin VB.Menu mnuTextFile 
            Caption         =   "&Text"
         End
         Begin VB.Menu mnuHTMLFile 
            Caption         =   "&HTML"
         End
      End
      Begin VB.Menu mnudash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "C&lear"
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmGlobalEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private Sub Form_Load()
    SetIcecue Line1, "icecue_shadow"
    SetIcecue Line2, "icecue_hilite"
    
    Set lstErrors.SmallIcons = frmMain.imgOperations(0)

    Me.Move dbSettings.GetProfileSetting("errLeft"), dbSettings.GetProfileSetting("errTop"), dbSettings.GetProfileSetting("errWidth"), dbSettings.GetProfileSetting("errHeight")
    Me.WindowState = dbSettings.GetProfileSetting("errState")
    lstErrors.ColumnHeaders(1).Width = dbSettings.GetProfileSetting("errColumn1")
    lstErrors.ColumnHeaders(2).Width = dbSettings.GetProfileSetting("errColumn2")
    lstErrors.ColumnHeaders(3).Width = dbSettings.GetProfileSetting("errColumn3")
    lstErrors.ColumnHeaders(4).Width = dbSettings.GetProfileSetting("errColumn4")
    lstErrors.ColumnHeaders(5).Width = dbSettings.GetProfileSetting("errColumn5")

    RefreshEvents

End Sub

Private Sub RefreshEvents()

    lstErrors.ListItems.Clear
    
    Dim dbConn As New clsDBConnection
    Dim rs As New ADODB.Recordset
    Dim newItem
    
    dbConn.rsQuery rs, "SELECT * FROM GlobalEvents WHERE ParentID=" & dbSettings.CurrentUserID & " ORDER BY EventTime;"
    Do Until rsEnd(rs)
            
        Set newItem = lstErrors.ListItems.Add(, , rs("Source"), , Trim(rs("IconKey")))
        newItem.SubItems(1) = rs("Location")
        newItem.SubItems(2) = rs("Action")
        newItem.SubItems(3) = rs("Message")
        newItem.SubItems(4) = rs("EventTime")
    
        rs.MoveNext
    Loop
    
    If lstErrors.ListItems.Count > 0 Then lstErrors.ListItems(lstErrors.ListItems.Count).EnsureVisible
    
    rsClose rs
    Set dbConn = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    dbSettings.SetProfileSetting "errState", Me.WindowState
    
    If Me.WindowState = 0 Then
        dbSettings.SetProfileSetting "errTop", Me.Top
        dbSettings.SetProfileSetting "errLeft", Me.Left
        dbSettings.SetProfileSetting "errHeight", Me.Height
        dbSettings.SetProfileSetting "errWidth", Me.Width
    End If
    
    dbSettings.SetProfileSetting "errColumn1", lstErrors.ColumnHeaders(1).Width
    dbSettings.SetProfileSetting "errColumn2", lstErrors.ColumnHeaders(2).Width
    dbSettings.SetProfileSetting "errColumn3", lstErrors.ColumnHeaders(3).Width
    dbSettings.SetProfileSetting "errColumn4", lstErrors.ColumnHeaders(4).Width
    dbSettings.SetProfileSetting "errColumn5", lstErrors.ColumnHeaders(5).Width

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    Const Border = 4
    
    Line1.X1 = 0
    Line1.X2 = ScaleWidth
    
    Line2.X1 = 0
    Line2.X2 = ScaleWidth
    
    lstErrors.Move (Border * Screen.TwipsPerPixelX), (Border * Screen.TwipsPerPixelY) * 2, ScaleWidth - ((Border * Screen.TwipsPerPixelX) * 2), ScaleHeight - ((Border * Screen.TwipsPerPixelY) * 3)

    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Private Sub lstErrors_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mnuEvents
    End If
End Sub

Private Sub mnuClear_Click()
    
    If MsgBox("Are you sure you want to clear all the error data?", vbQuestion + vbYesNo, AppName) = vbYes Then
                
        GlobalEvents.ClearEvents dbSettings

        lstErrors.ListItems.Clear

    End If
    
End Sub

Private Sub mnuClose_Click()
    Me.Hide
End Sub

Public Sub ExportEvents(ByVal FileName As String, ByVal FileFormat As String)

    Dim lstItem
    Dim filenum As Integer

    Select Case Trim(LCase(FileFormat))
        Case "html"
            filenum = FreeFile
            Open FileName For Output As #filenum
            
            Print #filenum, "<HTML><HEAD><TITLE>Max-FTP Global Events - [" & dbSettings.CurrentUserProfile & "]</TITLE></HEAD>"
            Print #filenum, "<BODY><CENTER>"
                
            Print #filenum, "<BR><BR><H3>Max-FTP Global Events - [" & dbSettings.CurrentUserProfile & "]</H3><BR>"
            Print #filenum, "Exported on " & Now() & "<BR><BR><BR>"
            
            Print #filenum, "<TABLE BORDER=2 CELLPADDING=5>"
            Print #filenum, "<TR><TD BGCOLOR=#c0c0c0><B>Source</B></TD><TD BGCOLOR=#c0c0c0><B>Location</B></TD><TD BGCOLOR=#c0c0c0><B>Action</B></TD><TD BGCOLOR=#c0c0c0><B>Message</B></TD><TD BGCOLOR=#c0c0c0><B>Date/Time</B></TD></TR>"
                
            For Each lstItem In lstErrors.ListItems
            
                Print #filenum, "<TR><TD>" & lstItem.Text & "</TD><TD>" & lstItem.SubItems(1) & "</TD><TD>" & lstItem.SubItems(2) & "</TD><TD>" & lstItem.SubItems(3) & "</TD><TD>" & lstItem.SubItems(4) & "</TD></TR>"
            
            Next
            
            Print #filenum, "<TR><TD BGCOLOR=#c0c0c0><B>Total Events</B></TD><TD BGCOLOR=#c0c0c0>&nbsp;</TD><TD BGCOLOR=#c0c0c0>&nbsp;</TD><TD BGCOLOR=#c0c0c0>&nbsp;</TD><TD BGCOLOR=#c0c0c0><B>" & lstErrors.ListItems.Count & "</B></TD></TR>"
            
            Print #filenum, "</TABLE>"
            
            Print #filenum, "</CENTER></BODY>"
            Print #filenum, "</HTML>"
            
            Close #filenum
        
        Case "text"
            
            filenum = FreeFile
            Open FileName For Output As #filenum
                
            For Each lstItem In lstErrors.ListItems
                        
                Print #filenum, lstItem.Text & String(2, vbKeyTab) & lstItem.SubItems(1) & String(2, vbKeyTab) & lstItem.SubItems(2) & String(2, vbKeyTab) & lstItem.SubItems(3) & String(2, vbKeyTab) & lstItem.SubItems(4)
            
            Next
        
            Close #filenum
                    
        Case Else
            Err.Raise 8, App.EXEName, "File Format not recognized. (Use 'HTML' or 'TEXT')"
    End Select

End Sub

Private Sub mnuHTMLFile_Click()
    On Error Resume Next
    CommonDialog1.CancelError = True
    CommonDialog1.DialogTitle = "Export to HTML"
    CommonDialog1.Filter = "HTML Files|*.htm;*.html|All Files|*.*"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.Action = 2
    If Err = 0 Then
        ExportEvents CommonDialog1.FileName, "html"
        OpenWebsite CommonDialog1.FileName, False
    Else
        Err.Clear
    End If
    On Error GoTo 0

End Sub

Private Sub mnuRefresh_Click()
    RefreshEvents
End Sub

Private Sub mnuTextFile_Click()
    On Error Resume Next
    CommonDialog1.CancelError = True
    CommonDialog1.DialogTitle = "Export to Text"
    CommonDialog1.Filter = "Text Files|*.txt|All Files|*.*"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.Action = 2
    If Err = 0 Then
        ExportEvents CommonDialog1.FileName, "text"
        OpenWebsite CommonDialog1.FileName, False
    Else
        Err.Clear
    End If
    On Error GoTo 0
End Sub
