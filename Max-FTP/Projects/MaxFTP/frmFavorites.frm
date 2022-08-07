VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFavorites 
   Caption         =   "Favorites"
   ClientHeight    =   3780
   ClientLeft      =   2472
   ClientTop       =   5160
   ClientWidth     =   7704
   HelpContextID   =   6
   Icon            =   "frmFavorites.frx":0000
   LinkTopic       =   "frmSiteSetup"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3780
   ScaleWidth      =   7704
   Tag             =   "sitesetup"
   Begin VB.PictureBox FavBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   4830
      ScaleHeight     =   648
      ScaleWidth      =   2568
      TabIndex        =   1
      Top             =   255
      Width           =   2565
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3435
      HelpContextID   =   6
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   6054
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   423
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6360
      Top             =   1425
      _ExtentX        =   1101
      _ExtentY        =   1101
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFavorites.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFavorites.frx":0BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFavorites.frx":117E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFavorites.frx":1718
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFavorites.frx":1FF2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   -75
      X2              =   3240
      Y1              =   30
      Y2              =   30
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   -60
      X2              =   4200
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Menu mnuFavorites 
      Caption         =   "&Favorites"
      Begin VB.Menu mnuMaxFav 
         Caption         =   "&Max Favorites"
      End
      Begin VB.Menu mnuWinFav 
         Caption         =   "&Windows Favorites"
      End
      Begin VB.Menu mnuDash32 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuAdvanced 
         Caption         =   "&Explore"
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Begin VB.Menu mnuEditSite 
         Caption         =   "&Edit Site"
      End
      Begin VB.Menu mnuOpenSite 
         Caption         =   "&Open Site"
      End
      Begin VB.Menu mnuDash4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewSite 
         Caption         =   "&New Site"
      End
      Begin VB.Menu mnuNewFolder 
         Caption         =   "New &Folder"
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRenameItem 
         Caption         =   "Raname &Item"
      End
      Begin VB.Menu mnuDash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemoveItem 
         Caption         =   "&Remove Item"
      End
   End
End
Attribute VB_Name = "frmFavorites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private RenameOld As String

Public Sub SetMenusEnabled()

    If TreeView1.SelectedItem.Text = PluginRootText Then
        mnuOpenSite.enabled = False
        mnuNewSite.enabled = True
        mnuNewFolder.enabled = True
        mnuEditSite.enabled = False
        mnuRemoveItem.enabled = False
        mnuRenameItem.enabled = False
    Else
        Select Case TreeView1.SelectedItem.Tag
        Case "1"
            mnuOpenSite.enabled = False
            mnuNewSite.enabled = True
            mnuNewFolder.enabled = True
            mnuEditSite.enabled = False
        Case FTPSiteExt
            mnuOpenSite.enabled = (TreeView1.SelectedItem.Tag = FTPSiteExt)
            mnuNewSite.enabled = False
            mnuNewFolder.enabled = False
            mnuEditSite.enabled = True
        Case Else
            mnuOpenSite.enabled = (TreeView1.SelectedItem.Tag = FTPSiteExt)
            mnuNewSite.enabled = False
            mnuNewFolder.enabled = False
            mnuEditSite.enabled = False
        End Select
        mnuRemoveItem.enabled = True
        mnuRenameItem.enabled = True
    End If

End Sub

Function FormulatePluginPath(ByVal tPath As String) As String
    Dim tmpPath As String
    tmpPath = GetMaxFavoritesDir(mnuWinFav.Checked)
    If InStr(tPath, "\") > 0 Then
        tmpPath = tmpPath & Mid(tPath, InStr(tPath, "\") + 1)
    End If
    If Right(tmpPath, 1) = "\" Then tmpPath = Left(tmpPath, Len(tmpPath) - 1)
    FormulatePluginPath = tmpPath
End Function

Private Function PluginRootText() As String

    If mnuMaxFav.Checked = True Then
        PluginRootText = "Max-FTP Favorites"
    Else
        PluginRootText = "Window Favorites"
    End If

End Function

Private Sub Form_Load()
    
    SetIcecue Line1, "icecue_shadow"
    SetIcecue Line2, "icecue_hilite"
    
    Me.Move dbSettings.GetProfileSetting("favLeft"), dbSettings.GetProfileSetting("favTop"), dbSettings.GetProfileSetting("favWidth"), dbSettings.GetProfileSetting("favHeight")
    Me.WindowState = dbSettings.GetProfileSetting("favState")

    InitializeFavorites

    SetPicture "favorites_graphic", FavBack

End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    Const Border = 4

    Line1.X1 = 0
    Line1.X2 = ScaleWidth

    Line2.X1 = 0
    Line2.X2 = ScaleWidth

    TreeView1.Move (Border * Screen.TwipsPerPixelX), (Border * Screen.TwipsPerPixelY) * 2, ScaleWidth - ((Border * Screen.TwipsPerPixelX) * 2), ScaleHeight - ((Border * Screen.TwipsPerPixelY) * 3)

    If TreeView1.Width - ((Border * Screen.TwipsPerPixelX) * 4) < FavBack.Width Then
        FavBack.Visible = False
    Else
        If TreeView1.Height - ((Border * Screen.TwipsPerPixelY) * 5) < FavBack.Height Then
            FavBack.Visible = False
        Else
            FavBack.Visible = True
            FavBack.Move TreeView1.Width - ((Border * Screen.TwipsPerPixelX) * 4) - FavBack.Width, ((Border * Screen.TwipsPerPixelY) * 5)
        End If
    End If

    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    
End Sub

Public Sub InitializeFavorites()

    Dim Favorite As Boolean
    Favorite = dbSettings.GetClientSetting("WinFavorites")
    
    mnuMaxFav.Checked = Not Favorite
    mnuWinFav.Checked = Favorite
    
    If Favorite Then
        Me.Caption = "Windows Favorites"
    Else
        Me.Caption = "Max-FTP Favorites"
    End If
    
    Dim Directory As String
    Directory = GetMaxFavoritesDir(mnuWinFav.Checked)

    If PathExists(Directory) Then
        TreeView1.Nodes.Clear
        Dim nodX As MSComctlLib.Node
        Dim fso As New Scripting.FileSystemObject
        Set nodX = TreeView1.Nodes.Add(, 4, , PluginRootText, 1)
    
        LoadFavoritesFolder fso, Directory, 1
    
        nodX.Selected = True
        nodX.Expanded = True
        nodX.Tag = "0"
    
        Set TreeView1.SelectedItem = TreeView1.Nodes(1)
    
        SetMenusEnabled

        Set fso = Nothing
    Else
        MsgBox "The directory which contains information about your websites was not able to be retrieved or was not able to be created.  You will not be able to save you sites information.  Accessing directory - '" + Directory + "' errored.  " + (Error), vbCritical, AppName
    End If
    
    Err.Clear
    On Error GoTo 0

End Sub
Private Sub LoadFavoritesFolder(ByRef fso As Scripting.FileSystemObject, ByVal Directory As String, ByVal Relative As Integer)
    Dim nodX As MSComctlLib.Node
    Dim theName As String
    
    Dim f2 As Folder
    Dim f3 As File

    For Each f2 In fso.GetFolder(Directory).SubFolders
        Set nodX = TreeView1.Nodes.Add(Relative, 4, , f2.Name, 2)
        nodX.Tag = "1"
        nodX.ExpandedImage = 3
        LoadFavoritesFolder fso, Directory + "\" + f2.Name, TreeView1.Nodes.Count
    Next

    For Each f3 In fso.GetFolder(Directory).Files
        If InStr(LCase(f3.Name), LCase(FTPSiteExt)) > 0 Then
            Set nodX = TreeView1.Nodes.Add(Relative, 4, , Left(f3.Name, InStrRev(f3.Name, ".") - 1), 4)
            nodX.Tag = FTPSiteExt
        ElseIf InStr(LCase(f3.Name), LCase(WebSiteExt)) > 0 Then
            Set nodX = TreeView1.Nodes.Add(Relative, 4, , Left(f3.Name, InStrRev(f3.Name, ".") - 1), 5)
            nodX.Tag = WebSiteExt
        ElseIf InStr(LCase(f3.Name), LCase(URLSiteExt)) > 0 Then
            Set nodX = TreeView1.Nodes.Add(Relative, 4, , Left(f3.Name, InStrRev(f3.Name, ".") - 1), 5)
            nodX.Tag = URLSiteExt
        End If
    Next

End Sub

Private Sub Form_Unload(Cancel As Integer)

    dbSettings.SetProfileSetting "favState", Me.WindowState
    
    If Me.WindowState = 0 Then
        dbSettings.SetProfileSetting "favTop", Me.Top
        dbSettings.SetProfileSetting "favLeft", Me.Left
        dbSettings.SetProfileSetting "favHeight", Me.Height
        dbSettings.SetProfileSetting "favWidth", Me.Width
    End If

End Sub

Private Sub mnuAction_Click()
    SetMenusEnabled
End Sub

Private Sub mnuAdvanced_Click()
    Dim nStr As String
    nStr = GetMaxFavoritesDir(mnuWinFav.Checked)
    If Right(nStr, 1) = "\" Then nStr = Left(nStr, Len(nStr) - 1)
    OpenWebsite nStr, False

End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub mnuEditSite_Click()

    EditSite

End Sub

Private Function EditSite()
    Dim fname As String
    fname = FormulatePluginPath(TreeView1.SelectedItem.FullPath + FTPSiteExt)
    If PathExists(fname, True) Then
    
        Dim frm As Form
        For Each frm In Forms
            If TypeName(frm) = "frmFavoriteSite" Then
                If LCase(frm.FileName) = LCase(fname) Then
                    frm.Show
                    Exit Function
                End If
            End If
        Next
    
        Dim eSite As New frmFavoriteSite
        eSite.LoadSite fname
        eSite.Caption = TreeView1.SelectedItem.Text
        eSite.Show
    End If

End Function

Function MakeNewWebSiteTitle(ByVal Directory As String, ByVal IsFolder As Integer, ByVal BaseName As String) As String
    
    Dim NewTitle As String
    Dim NewNum As Integer
    NewNum = 0
    NewTitle = BaseName + " "
    
    If IsFolder Then
        Do
            NewNum = NewNum + 1
        Loop Until Not PathExists(Directory & "\" & NewTitle + Trim(str(NewNum)))
    Else
        Do
            NewNum = NewNum + 1
        Loop Until Not PathExists(Directory & "\" & NewTitle + Trim(str(NewNum)) & FTPSiteExt)
    End If
    
    MakeNewWebSiteTitle = NewTitle + Trim(str(NewNum))

End Function

Private Sub mnuNewFolder_Click()

    Dim nodX As MSComctlLib.Node
    Dim FolderName As String
    FolderName = MakeNewWebSiteTitle(FormulatePluginPath(TreeView1.SelectedItem.FullPath), True, "New Folder")
    MkDir FormulatePluginPath(TreeView1.SelectedItem.FullPath) & "\" & FolderName
        
    TreeView1.SelectedItem.Expanded = True
    TreeView1.SelectedItem.Selected = False
    If TreeView1.SelectedItem.Index > 1 Then TreeView1.SelectedItem.Image = 3
    Set nodX = TreeView1.Nodes.Add(TreeView1.SelectedItem.Index, 4, , FolderName, 2)
    nodX.Tag = "1"
    nodX.Selected = True
    Set TreeView1.SelectedItem = nodX
    SetMenusEnabled
        
    nodX.EnsureVisible
    TreeView1.StartLabelEdit

End Sub

Public Sub RemoveFTPSitePlugins(ByVal dataLocation As String)

    Dim cnt As Integer
    If InStr(dataLocation, FTPSiteExt) > 0 Then
        Kill dataLocation
    Else
        Dim fso As New Scripting.FileSystemObject
        Dim f2 As Folder
        Dim f3 As File
        
        For Each f3 In fso.GetFolder(dataLocation).Files
            Kill dataLocation & "\" & f3.Name
        Next
        
        For Each f2 In fso.GetFolder(dataLocation).SubFolders
            RemoveFTPSitePlugins dataLocation & "\" & f2.Name
        Next
        
        RmDir dataLocation
        
        Set fso = Nothing
        Set f2 = Nothing
        Set f3 = Nothing
    End If

End Sub

Private Sub mnuNewSite_Click()

    Dim nodX As MSComctlLib.Node
    Dim FileName As String
    Dim newSite As New frmFavoriteSite
    
    newSite.Caption = MakeNewWebSiteTitle(FormulatePluginPath(TreeView1.SelectedItem.FullPath), False, "FTP Site Information")
    FileName = FormulatePluginPath(TreeView1.SelectedItem.FullPath + "\" + newSite.Caption + FTPSiteExt)
    
    newSite.SaveSite FileName
    
    TreeView1.SelectedItem.Expanded = True
    TreeView1.SelectedItem.Selected = False
    If TreeView1.SelectedItem.Index > 1 Then TreeView1.SelectedItem.Image = 3
    Set nodX = TreeView1.Nodes.Add(TreeView1.SelectedItem.Index, 4, , newSite.Caption, 4)
    nodX.Tag = FTPSiteExt
    nodX.Selected = True
    Set TreeView1.SelectedItem = nodX
    Unload newSite
    SetMenusEnabled
    nodX.EnsureVisible
    TreeView1.StartLabelEdit
    
    frmMain.RefreshFavorites
    
End Sub

Private Sub OpenSite(ByVal InNew As Boolean)
        Dim FileName As String
        Dim cnt As Integer
        
        FileName = FormulatePluginPath(TreeView1.SelectedItem.FullPath + TreeView1.SelectedItem.Tag)
        Select Case TreeView1.SelectedItem.Tag
            Case FTPSiteExt
                Dim OpenSite As New frmFavoriteSite
                OpenSite.LoadSite FileName
                
                Dim topClient As Form
                Set topClient = GetTopMostClientGUI()
                
                If topClient Is Nothing Or InNew Then
                    Dim newFTPClient As New frmFTPClientGUI
                    newFTPClient.LoadClient
                    newFTPClient.ShowClient
                    newFTPClient.FTPOpenSite OpenSite
                Else
                    topClient.Show
                    topClient.FTPOpenSite OpenSite
                End If
                
                Unload OpenSite
                
            Case Else
                OpenAssociatedFile FileName, False
        End Select
End Sub

Private Sub mnuOpenInNew_Click()
    OpenSite True
End Sub

Private Sub mnuOpenSite_Click()
    OpenSite True
End Sub

Private Sub mnuRefresh_Click()
    InitializeFavorites
    
End Sub

Private Sub mnuRemoveItem_Click()

        If TreeView1.SelectedItem.Tag <> "1" Then
            If MsgBox("Are you sure you want to delete the FTP site - '" + TreeView1.SelectedItem.Text + "'?", vbQuestion + vbYesNo, AppName) = vbYes Then
                RemoveFTPSitePlugins FormulatePluginPath(TreeView1.SelectedItem.FullPath + FTPSiteExt)
                
                Dim TIndex1 As Integer
                TIndex1 = TreeView1.SelectedItem.Index
                TreeView1.Nodes.Remove TIndex1
                TreeView1.Nodes(TIndex1 - 1).Selected = True
                Set TreeView1.SelectedItem = TreeView1.Nodes(TIndex1 - 1)
            
            End If
        Else
            If MsgBox("Are you sure you want to delete the folder - '" + TreeView1.SelectedItem.Text + "'?  This will delete all the FTP sites it contains.", vbInformation + vbYesNo, AppName) = vbYes Then
                RemoveFTPSitePlugins FormulatePluginPath(TreeView1.SelectedItem.FullPath)
            
                Dim TIndex As Integer
                TIndex = TreeView1.SelectedItem.Index
                TreeView1.Nodes.Remove TIndex
                TreeView1.Nodes(TIndex - 1).Selected = True
                Set TreeView1.SelectedItem = TreeView1.Nodes(TIndex - 1)
            
            End If
        End If
       
        SetMenusEnabled
        
        frmMain.RefreshFavorites

End Sub

Private Sub mnuRenameItem_Click()
        
    TreeView1.StartLabelEdit
    
End Sub

Private Sub mnuMaxFav_Click()
    
    dbSettings.SetClientSetting "WinFavorites", False
    
    InitializeFavorites
    
    frmMain.RefreshFavorites
    
End Sub

Private Sub mnuWinFav_Click()
    dbSettings.SetClientSetting "WinFavorites", True
    
    InitializeFavorites
    
    frmMain.RefreshFavorites

End Sub

Private Sub TreeView1_AfterLabelEdit(Cancel As Integer, NewString As String)

    Dim tName As String
    tName = FormulatePluginPath(TreeView1.SelectedItem.FullPath)
    tName = Left(tName, InStrRev(tName, "\") - 1)
    
    If PathExists(tName & "\" & NewString) Or PathExists(tName & "\" & NewString & FTPSiteExt) Then
        MsgBox "There is already an item in that folder with the same name.  Please choose a different one.", vbInformation, AppName
        Cancel = True
    Else
        If Not IsFileNameValid(NewString) Then
            MsgBox "The name can not contain the characters  \ / : * ? "" < > | ", vbInformation, AppName
            Cancel = True
        Else
            If TreeView1.SelectedItem.Tag = "1" Then
                Name tName + "\" + RenameOld As tName + "\" + NewString
            Else
                Name tName + "\" + RenameOld + FTPSiteExt As tName + "\" + NewString + FTPSiteExt
            End If
        End If
    End If
    
    frmMain.RefreshFavorites
    
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

    If TreeView1.SelectedItem.Text = PluginRootText Then
        Cancel = True
    Else
        RenameOld = TreeView1.SelectedItem.Text
    End If

End Sub

Private Sub TreeView1_DblClick()
    If Not TreeView1.SelectedItem Is Nothing Then
        EditSite

    End If
End Sub

Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        SetMenusEnabled
        Me.PopupMenu mnuAction
    End If

End Sub
