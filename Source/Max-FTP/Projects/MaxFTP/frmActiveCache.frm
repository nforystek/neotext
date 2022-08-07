VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmActiveCache 
   Caption         =   "Active App Cache"
   ClientHeight    =   4200
   ClientLeft      =   6135
   ClientTop       =   345
   ClientWidth     =   9960
   Icon            =   "frmActiveCache.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   4368
      TabIndex        =   7
      Top             =   3804
      Visible         =   0   'False
      Width           =   2016
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      Height          =   312
      Index           =   4
      Left            =   3408
      TabIndex        =   6
      Top             =   3780
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.PictureBox pFileIcon 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   8925
      ScaleHeight     =   240
      ScaleMode       =   0  'User
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   312
      Index           =   3
      Left            =   7248
      TabIndex        =   4
      Top             =   3792
      Width           =   852
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Remove"
      Height          =   312
      Index           =   2
      Left            =   2328
      TabIndex        =   3
      Top             =   3792
      Width           =   852
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Upload"
      Height          =   312
      Index           =   1
      Left            =   1248
      TabIndex        =   2
      Top             =   3792
      Width           =   852
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Open"
      Height          =   312
      Index           =   0
      Left            =   156
      TabIndex        =   1
      Top             =   3792
      Width           =   852
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3264
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   8196
      _ExtentX        =   14446
      _ExtentY        =   5768
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "FileName"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Location"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "AlteredDate"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Uploaded"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Cached File"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList imgFiles 
      Left            =   9315
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActiveCache.frx":08CA
            Key             =   "folder"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuUpload 
         Caption         =   "&Upload"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "&Remove"
      End
   End
End
Attribute VB_Name = "frmActiveCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private CurrentAnalysis As Long

Private Const Border = 4

Private WithEvents cacheWatch As NTSchedule20.Timer
Attribute cacheWatch.VB_VarHelpID = -1

Private dbConn As New clsDBConnection

Public Function ExistsInCache(ByVal FileName As String, ByVal Location As String, ByVal Username As String, ByVal Password As String) As String
    Dim enc As New NTCipher10.ncode
    Dim rs As New ADODB.Recordset
    Dim ret As String

    If Not PathExists(GetTemporaryFolder & "\" & Left(ActiveAppFolder, Len(ActiveAppFolder) - 1)) Then
        MkDir GetTemporaryFolder & "\" & Left(ActiveAppFolder, Len(ActiveAppFolder) - 1)
    End If
    
    dbConn.rsQuery rs, "SELECT * FROM ActiveApp WHERE ParentID=" & dbSettings.CurrentUserID & " AND FileName='" & FileName & "' AND Location='" & _
            Replace(Location, "'", "''") & "' AND Username='" & enc.EncryptString(Replace(Username, "'", "''"), dbSettings.CryptKey, True) & _
            "' AND Password='" & enc.EncryptString(Replace(Password, "'", "''"), dbSettings.CryptKey(Replace(Username, "'", "''")), True) & "';"
    
    If Not rsEnd(rs) Then
        ret = rs("SubFolder")
    Else

        Do
            If Err Then Err.Clear
            ret = Replace(modGuid.GUID, "-", "")
            MkDir GetTemporaryFolder & "\" & ActiveAppFolder & ret
        Loop While Not (Err.Number = 0)

    End If

    ExistsInCache = ret
    
    rsClose rs
    
    Set rs = Nothing
    Set enc = Nothing
End Function
Public Sub AddToCache(ByVal FileName As String, ByVal Location As String, ByVal Username As String, ByVal Password As String, ByVal gID As String)
    Dim enc As New NTCipher10.ncode
    Dim rs As New ADODB.Recordset

    If Not PathExists(GetTemporaryFolder & "\" & Left(ActiveAppFolder, Len(ActiveAppFolder) - 1)) Then
        MkDir GetTemporaryFolder & "\" & Left(ActiveAppFolder, Len(ActiveAppFolder) - 1)
    End If

    
    dbConn.rsQuery rs, "SELECT * FROM ActiveApp WHERE ParentID=" & dbSettings.CurrentUserID & " AND FileName='" & FileName & "' AND Location='" & _
            Replace(Location, "'", "''") & "' AND SubFolder='" & Replace(gID, "'", "''") & "';"
    
    If Not rsEnd(rs) Then
        
        dbConn.rsQuery rs, "UPDATE ActiveApp SET Username='" & enc.EncryptString(Replace(Username, "'", "''"), _
                dbSettings.CryptKey, True) & "', Password='" & enc.EncryptString(Replace(Password, "'", "''"), _
                dbSettings.CryptKey(Replace(Username, "'", "''")), True) & "', AlteredDate='" & _
                GetFileDate(Replace(GetTemporaryFolder & "\" & ActiveAppFolder & gID & "\" & FileName, "\\", "\")) & _
                "', Uploaded=No WHERE ParentID=" & dbSettings.CurrentUserID & " AND FileName='" & _
                Replace(FileName, "'", "''") & "' AND Location='" & Replace(Location, "'", "''") & _
                "' AND SubFolder='" & Replace(gID, "'", "''") & "';"
    
    Else
    
        dbConn.rsQuery rs, "INSERT INTO ActiveApp (ParentID, FileName, Location, Username, Password, AlteredDate, Uploaded, SubFolder) VALUES (" & _
                dbSettings.CurrentUserID & ",'" & Replace(FileName, "'", "''") & "', '" & Replace(Location, "'", "''") & "', '" & _
                enc.EncryptString(Replace(Username, "'", "''"), dbSettings.CryptKey, True) & "', '" & enc.EncryptString(Replace(Password, "'", "''"), _
                dbSettings.CryptKey(Replace(Username, "'", "''")), True) & "', '" & GetFileDate(Replace(GetTemporaryFolder & "\" & ActiveAppFolder & gID & "\" & _
                FileName, "\\", "\")) & "', No, '" & Replace(gID, "'", "''") & "');"
    
    End If

    RefreshActiveAppList

    rsClose rs
    
    Set rs = Nothing
    Set enc = Nothing
End Sub

Public Sub RemoveFromCache(ByVal FileName As String, ByVal Location As String, ByVal gID As String)

    cacheWatch.enabled = False
    
    Dim rs As New ADODB.Recordset

    dbConn.rsQuery rs, "SELECT * FROM ActiveApp WHERE ParentID=" & dbSettings.CurrentUserID & " AND Location='" & Replace(Location, "'", "''") & _
                "' AND FileName='" & Replace(FileName, "'", "''") & "' AND SubFolder='" & Replace(gID, "'", "''") & "';"
    
    If Not rsEnd(rs) Then
    
        dbConn.rsQuery rs, "DELETE FROM ActiveApp WHERE ParentID=" & dbSettings.CurrentUserID & " AND Location='" & Replace(Location, "'", "''") & _
                    "' AND FileName='" & Replace(FileName, "'", "''") & "' AND SubFolder='" & Replace(gID, "'", "''") & "';"

    End If
    
    rsClose rs
    
    If PathExists(Replace(GetTemporaryFolder & "\" & ActiveAppFolder & gID & "\" & FileName, "\\", "\"), True) Then
        SetAttr Replace(GetTemporaryFolder & "\" & ActiveAppFolder & gID & "\" & FileName, "\\", "\"), VbFileAttribute.vbNormal
        Kill Replace(GetTemporaryFolder & "\" & ActiveAppFolder & gID & "\" & FileName, "\\", "\")
    End If
    If PathExists(Replace(GetTemporaryFolder & "\" & ActiveAppFolder & gID, "\\", "\"), False) Then
        RmDir Replace(GetTemporaryFolder & "\" & ActiveAppFolder & gID, "\\", "\")
    End If
        
    RefreshActiveAppList
    
    cacheWatch.enabled = True
        
End Sub

Private Sub RefreshActiveAppList()
    Dim rs As New ADODB.Recordset
    Dim fileType As String
    Dim lstItem As MSComctlLib.ListItem
    
    For Each lstItem In ListView1.ListItems
        lstItem.Tag = True
    Next
    
    If ListView1.ListItems.count = 0 Then ListView1.SmallIcons = frmMain.imgFiles
    dbConn.rsQuery rs, "SELECT * FROM ActiveApp;"
    Do Until rsEnd(rs)
                
        fileType = GetFileExt(rs("FileName"))
        
        GetAssociation fileType, fileType, pFileIcon(1)
        LoadAssociation pFileIcon(1)
        If (ListView1.ListItems.count > 0) Then
            ListView1.Tag = ListView1.ListItems.count
            If (lstItem Is Nothing) Then Set lstItem = ListView1.ListItems(1)
            
            Do While (Not (lstItem.Key = "FILE" & rs("ID"))) And (ListView1.Tag > 0)
                ListView1.Tag = ListView1.Tag - 1
                If lstItem.Index = ListView1.ListItems.count Then
                    Set lstItem = ListView1.ListItems(1)
                Else
                    Set lstItem = ListView1.ListItems(lstItem.Index + 1)
                End If
            Loop
            If (Not (lstItem.Key = "FILE" & rs("ID"))) Or (ListView1.ListItems.count = 0) Then
                ListView1.Tag = -1
            End If
        Else
            ListView1.Tag = -1
        End If
        
        If ListView1.Tag = -1 Then
            Set lstItem = ListView1.ListItems.Add(, "FILE" & rs("ID"), rs("FileName"), , pFileIcon(1).Tag)
        End If

        lstItem.SubItems(1) = rs("Location")
        lstItem.SubItems(2) = rs("AlteredDate")
        lstItem.SubItems(3) = rs("Uploaded")
        lstItem.SubItems(4) = rs("SubFolder")
        lstItem.Tag = False
        
        rs.MoveNext
    Loop
    
    rsClose rs
    Set rs = Nothing

    Dim cnt As Long
    cnt = 1
    Do While cnt <= ListView1.ListItems.count

        If ListView1.ListItems(cnt).Tag = True Then
            ListView1.ListItems.Remove ListView1.ListItems(cnt).Index
        Else
            cnt = cnt + 1
        End If

    Loop

End Sub

Public Sub ShowForm()
    
    RefreshActiveAppList
   
    If Me.WindowState = 1 Then Me.WindowState = 0
    Me.Show
    
End Sub
Private Sub cacheWatch_OnTicking()
On Error GoTo fileissue
    
    Dim FilePath As String
    If CurrentAnalysis > ListView1.ListItems.count Or CurrentAnalysis = 0 Then CurrentAnalysis = 1
    If CurrentAnalysis <= ListView1.ListItems.count Then

        FilePath = Replace(GetTemporaryFolder & "\" & ActiveAppFolder & ListView1.ListItems(CurrentAnalysis).SubItems(4) & "\" & ListView1.ListItems(CurrentAnalysis).Text, "\\", "\")
        If PathExists(FilePath, True) Then
        
            If Not CDate(GetFileDate(FilePath)) = CDate(ListView1.ListItems(CurrentAnalysis).SubItems(2)) Then
                
                If dbSettings.GetClientSetting("ActiveAppUpload") Then
                    If Not CBool(ListView1.ListItems(CurrentAnalysis).SubItems(3)) Then
                        ListView1.ListItems(CurrentAnalysis).Selected = True
                        
                        cacheWatch.enabled = False
                        
                        Command1_Click 1
                    
                    Else
                        If ListView1.ListItems(CurrentAnalysis).SubItems(3) <> "False" Then
                        
                            ListView1.ListItems(CurrentAnalysis).SubItems(3) = "False"
                        End If
                    
                    End If
                Else
                    If ListView1.ListItems(CurrentAnalysis).SubItems(3) <> "False" Then
                        ListView1.ListItems(CurrentAnalysis).SubItems(3) = "False"
                    End If
                End If
            End If
        End If
    End If

    Exit Sub
fileissue:
    Err.Clear
    Resume Next
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            If Not ListView1.SelectedItem Is Nothing Then
            
                OpenAssociatedFile Replace(GetTemporaryFolder & "\" & ActiveAppFolder & ListView1.SelectedItem.SubItems(4) & "\" & ListView1.SelectedItem.Text, "\\", "\"), False
            
            Else
                MsgBox "Please select the cached file you want to open.", vbInformation, AppName
            End If
        
        Case 1
            If Not ListView1.SelectedItem Is Nothing Then
                
                cacheWatch.enabled = False
                
                Dim useForm As frmFTPClientGUI
                Dim useIndex As Integer
                
                Set useForm = New frmFTPClientGUI
                useForm.LoadClient "Active App Cache "
                useIndex = 0
                useForm.setLocation(0).Text = ListView1.SelectedItem.SubItems(1)
                useForm.FTPConnect 0
                    
                Do While useForm.GetState(0) = st_Processing
                    DoTasks
                Loop
                
                If useForm.GetState(0) = st_ProcessSuccess Then
                    
                    Dim itemIndex As Integer
                
                    useForm.setLocation(1).Text = Replace(GetTemporaryFolder & "\" & ActiveAppFolder & ListView1.SelectedItem.SubItems(4) & "\", "\\", "\")
                    useForm.FTPConnect 1
                    Do While useForm.GetState(1) = st_Processing
                        DoEvents
                    Loop
                
                    itemIndex = IsOnListItems(useForm.pView(1), ListView1.SelectedItem.Text)
                
                    If itemIndex > 0 Then
                        Dim cnt As Integer
                        For cnt = 1 To useForm.pView(1).ListItems.count
                            useForm.pView(1).ListItems(cnt).Selected = (itemIndex = cnt)
                        Next
                        CopyFiles useForm, 1, "Copy"
                        PasteFromClipboard useForm, 0
                    End If

                    If dbSettings.GetClientSetting("ActiveAppRemove") Then
                        RemoveFromCache ListView1.SelectedItem.Text, ListView1.SelectedItem.SubItems(1), ListView1.SelectedItem.SubItems(4)
                    Else
                        Dim enc As New NTCipher10.ncode
                        dbConn.dbQuery "UPDATE ActiveApp SET AlteredDate='" & GetFileDate(Replace(GetTemporaryFolder & "\" & ActiveAppFolder & ListView1.SelectedItem.SubItems(4) & "\" & ListView1.SelectedItem.Text, "\\", "\")) & "', Uploaded=Yes WHERE ParentID=" & dbSettings.CurrentUserID & " AND FileName='" & Replace(ListView1.SelectedItem.Text, "'", "''") & "' AND Location='" & Replace(ListView1.SelectedItem.SubItems(1), "'", "''") & "' AND SubFolder='" & Replace(ListView1.SelectedItem.SubItems(4), "'", "''") & "';"
                        Set enc = Nothing
                        RefreshActiveAppList
                    End If

                End If
                Unload useForm
                Set useForm = Nothing
                
                cacheWatch.enabled = True
            
            Else
                MsgBox "Please select the cached file you want to re-upload.", vbInformation, AppName
            End If
            
        Case 2
            If Not ListView1.SelectedItem Is Nothing Then
            
                RemoveFromCache ListView1.SelectedItem.Text, ListView1.SelectedItem.SubItems(1), ListView1.SelectedItem.SubItems(4)
                
            Else
                MsgBox "Please select the cached file you want to remove.", vbInformation, AppName
            End If
            
        Case 3
            Unload Me
    End Select

End Sub



Private Sub Form_Load()
    Me.Tag = True

    Set cacheWatch = New NTSchedule20.Timer
    cacheWatch.Interval = 100
    cacheWatch.enabled = True

End Sub

Public Sub FormResize(ByRef myForm)
    With myForm
        On Error Resume Next
        .ListView1.Top = (Border * Screen.TwipsPerPixelY)
        .ListView1.Left = (Border * Screen.TwipsPerPixelX)
        .ListView1.Width = .ScaleWidth - ((Border * Screen.TwipsPerPixelX) * 2)
        .ListView1.Height = .ScaleHeight - ((Border * Screen.TwipsPerPixelY) * 2) - .Command1(0).Height


        .Command1(0).Left = (Border * Screen.TwipsPerPixelX)
        .Command1(1).Left = .Command1(0).Left + .Command1(0).Width + (Border * Screen.TwipsPerPixelX)
        .Command1(2).Left = .Command1(1).Left + .Command1(1).Width + (Border * Screen.TwipsPerPixelX)
        .Command1(4).Left = .Command1(2).Left + .Command1(2).Width + (Border * Screen.TwipsPerPixelX)
        
        .Command1(3).Left = .ScaleWidth - (.Command1(3).Width + (Border * Screen.TwipsPerPixelX))

        .Command1(0).Top = .ListView1.Height + (Border * Screen.TwipsPerPixelY)
        .Command1(1).Top = .ListView1.Height + (Border * Screen.TwipsPerPixelY)
        .Command1(2).Top = .ListView1.Height + (Border * Screen.TwipsPerPixelY)
        .Command1(3).Top = .ListView1.Height + (Border * Screen.TwipsPerPixelY)
        .Command1(4).Top = .ListView1.Height + (Border * Screen.TwipsPerPixelY)
        
        .BrowseButton1.Left = Command1(3).Left - (Border * Screen.TwipsPerPixelX) - .BrowseButton1.Width
        .BrowseButton1.Top = ((Border * Screen.TwipsPerPixelY) * 2) + .ListView1.Height
        .Text1.Top = (Border * Screen.TwipsPerPixelY) + .ListView1.Height
        .Text1.Left = .Command1(4).Left + .Command1(3).Width + (Border * Screen.TwipsPerPixelX)
        .Text1.Width = .ScaleWidth - .Text1.Left - .Command1(3).Width - ((Border * Screen.TwipsPerPixelX) * 3) - .BrowseButton1.Width


        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        
    End With
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    cacheWatch.enabled = False
End Sub

Private Sub Form_Resize()
    FormResize Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    cacheWatch.enabled = False

    Dim lstItem As MSComctlLib.ListItem
    Set ListView1.SmallIcons = Nothing
    For Each lstItem In ListView1.ListItems
        RemoveAssociation lstItem.SmallIcon
    Next
    ListView1.ListItems.Clear
    
        
    Set dbConn = Nothing

    cacheWatch.enabled = False

    
    Set cacheWatch = Nothing
End Sub



Private Sub ListView1_DblClick()

    If Not ListView1.SelectedItem Is Nothing Then

        OpenAssociatedFile Replace(GetTemporaryFolder & "\" & ActiveAppFolder & ListView1.SelectedItem.SubItems(4) & "\" & ListView1.SelectedItem.Text, "\\", "\"), False
    End If

End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo itemdeleted
    
    Static lastSel As MSComctlLib.ListItem
    If lastSel Is Nothing Then
        Set lastSel = Item
    ElseIf lastSel.SubItems(4) = Item.SubItems(4) Then
        
        If lastSel.Selected Then
            PopUp Me.hwnd, 0
        End If
        
    End If
    
itemdeleted:
    If Err.Number = 35605 Then Err.Clear
    
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button Then
        PopUp Me.hwnd, 0
    End If
End Sub


Private Sub mnuOpen_Click()
    Command1_Click 0
End Sub

Private Sub mnuRemove_Click()
    Command1_Click 2
End Sub

Private Sub mnuUpload_Click()
    Command1_Click 1
End Sub
