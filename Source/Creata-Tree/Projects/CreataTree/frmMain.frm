VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   Caption         =   "Creata-Tree"
   ClientHeight    =   5475
   ClientLeft      =   7275
   ClientTop       =   4965
   ClientWidth     =   6555
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHead 
      Height          =   2055
      Left            =   225
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "frmMain.frx":08CA
      Top             =   3270
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.TextBox txtBody 
      Height          =   1875
      Left            =   3240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Text            =   "frmMain.frx":1DAC
      Top             =   3225
      Visible         =   0   'False
      Width           =   2145
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      CausesValidation=   0   'False
      Height          =   2805
      Left            =   3180
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   225
      Width           =   2970
      ExtentX         =   5239
      ExtentY         =   4948
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin CreataTree.ctlNav nTree 
      Height          =   2895
      Left            =   105
      TabIndex        =   0
      Top             =   225
      Visible         =   0   'False
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   5106
   End
   Begin VB.Menu mnuTree 
      Caption         =   "&Tree"
      Begin VB.Menu mnuNew 
         Caption         =   "&New.."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As.."
      End
      Begin VB.Menu mnuDash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Propert&ies.."
      End
      Begin VB.Menu mnuDash14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export.."
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecent 
         Caption         =   "(No Recent Files)"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuRecent 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecent 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecent 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDash1345 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuItem 
      Caption         =   "&Item"
      Begin VB.Menu mnuNewSubItem 
         Caption         =   "&Add"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuDash4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCutItem 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopyItem 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPasteItem 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuDash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMove 
         Caption         =   "Move &Up"
         Index           =   0
      End
      Begin VB.Menu mnuMove 
         Caption         =   "Move &Down"
         Index           =   1
      End
      Begin VB.Menu mnuDash13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemoveItem 
         Caption         =   "&Remove"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuDash6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuItemProperties 
         Caption         =   "Propert&ies.."
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuGraphicalEditing 
         Caption         =   "&Graphical Editing"
      End
      Begin VB.Menu mnuPreviewPane 
         Caption         =   "&Preview Pane"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuDash9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuForceDimensions 
         Caption         =   "&Force Dimensions"
      End
      Begin VB.Menu mnuDash12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPromptForTemplate 
         Caption         =   "&Use Templates"
      End
      Begin VB.Menu mnuAlwaysUseBlank 
         Caption         =   "Use &Blank Tree"
      End
      Begin VB.Menu mnuDash11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMedia 
         Caption         =   "&Media Library.."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuExamples 
         Caption         =   "&Examples"
         Begin VB.Menu mnuExplorerStyleTree 
            Caption         =   "&Explorer Style Tree"
         End
         Begin VB.Menu mnuHelpStyleTree 
            Caption         =   "&Help Style Tree"
         End
         Begin VB.Menu mnuMacintoshStyleTree 
            Caption         =   "Ma&cintosh Style Tree"
         End
         Begin VB.Menu mnuMenuStyleTree 
            Caption         =   "&Menu Style Tree"
         End
      End
      Begin VB.Menu mnuDash10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDocumentation 
         Caption         =   "&Documentation..."
      End
      Begin VB.Menu mnuNeotext 
         Caption         =   "&Neotext.org..."
      End
      Begin VB.Menu mnuDash7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About.."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public WithEvents nForm As clsFileForm
Attribute nForm.VB_VarHelpID = -1
Public nBase As clsItem

Private Sub Form_Resize()
    EnableForm
End Sub

Private Sub mnuClose_Click()
    nForm.CloseFile
End Sub

Private Sub mnuDocumentation_Click()

    If PathExists(AppPath & "Help\index.htm", True) Then
        RunFile AppPath & "help\index.htm"
    Else
        RunFile AppWebsite & "/ipub/help/creata-tree"
    End If
    
End Sub

Private Sub mnuExplorerStyleTree_Click()
    
    nForm.OpenFile ExampleFolder & "Explorer Style.tree"
End Sub

Private Sub mnuExport_Click()
    
    frmExport.ShowForm nBase
    frmExport.Show 1
    
End Sub

Private Sub mnuHelpStyleTree_Click()
    nForm.OpenFile ExampleFolder & "Help Style.tree"
End Sub

Private Sub mnuItem_Click()
    EnableItemMenu
End Sub

Private Sub mnuMacintoshStyleTree_Click()
    nForm.OpenFile ExampleFolder & "Macintosh Style.tree"
End Sub

Private Sub mnuMedia_Click()
    frmMedia.Show 1, Me
    Unload frmMedia
End Sub

Private Sub mnuMenuStyleTree_Click()
    nForm.OpenFile ExampleFolder & "Menu Style.tree"
End Sub

Private Sub mnuMove_Click(Index As Integer)
    If Not (nTree.SelectedNode Is Nothing) Then
        If Not nTree.SelectedItem.IsBase Then
        
            With nTree.SelectedNode
                Select Case Index
                    Case 0
                        
                        If Not (.Key = .FirstSibling.Key) Then
                            
                            If Not (.Previous Is Nothing) Then
                            
                                SwapItems .Tag, .Previous.Tag
                                
                            End If
                        
                        End If
                        
                    Case 1
                        If Not (.Key = .LastSibling.Key) Then
                            
                            If Not (.Next Is Nothing) Then
                                SwapItems .Tag, .Next.Tag
                            End If
                        
                        End If
                    
                End Select
            End With
        End If
    End If
End Sub
Private Sub SwapItems(ByRef nItem1 As clsItem, ByRef nItem2 As clsItem)
    Dim nXML1 As String
    Dim nXML2 As String

    nXML1 = nItem1.XMLText
    nXML2 = nItem2.XMLText
    nItem1.XMLText = nXML2
    nItem2.XMLText = nXML1

    RefreshForm nItem2.Key
End Sub
Private Sub mnuCopyItem_Click()
    Clipboard.SetText "Creata-Tree:" & HesEncodeData(nTree.SelectedItem.XMLText)
End Sub

Private Sub mnuCutItem_Click()
    
    Clipboard.SetText "Creata-Tree:" & HesEncodeData(nTree.SelectedItem.XMLText)
    nTree.SelectedItem(True).RemoveItem nTree.SelectedItem.Key
    
    nForm.Changed = True
    RefreshForm
End Sub

Private Sub mnuPasteItem_Click()
    nForm.Changed = True
    RefreshForm nTree.SelectedItem.AddItem(HesDecodeData(Mid(Clipboard.GetText, 13))).Key
End Sub

Private Sub mnuItemProperties_Click()
    If nTree.SelectedItem.IsBase Then
        mnuProperties_Click
    Else
        frmItem.XMLText = nTree.SelectedItem.XMLText
        frmItem.Show 1, Me
        If frmItem.IsOk Then
            nTree.SelectedItem.XMLText = frmItem.XMLText
            nForm.Changed = True
            RefreshForm nTree.SelectedItem.Key
        End If
        Unload frmItem
    End If
End Sub

Private Sub mnuForceDimensions_Click()
    Ini.Setting("ForceDimensions") = Not Ini.Setting("ForceDimensions")
    EnableForm
End Sub

Private Sub mnuGraphicalEditing_Click()
    Ini.Setting("ViewGraphicalEdit") = Not Ini.Setting("ViewGraphicalEdit")
    RefreshForm
End Sub

Private Sub mnuPreviewPane_Click()
    Ini.Setting("ViewPreviewPane") = Not Ini.Setting("ViewPreviewPane")
    RefreshBrowser
End Sub

Private Sub mnuPromptForTemplate_Click()
    Ini.Setting("PromptForTemplate") = True
    EnableForm
End Sub
Private Sub mnuAlwaysUseBlank_Click()
    Ini.Setting("PromptForTemplate") = False
    EnableForm
End Sub

Private Sub mnuProperties_Click()
    frmTree.XMLText = nBase.XMLText
    frmTree.Show 1, Me
    If frmTree.IsOk Then
        nBase.XMLText = frmTree.XMLText
        nForm.Changed = True
        RefreshForm nBase.Key
    End If
    Unload frmTree
End Sub

Private Sub mnuRecent_Click(Index As Integer)
    If PathExists(mnuRecent(Index).Tag, True) Then
        nForm.OpenFile mnuRecent(Index).Tag
    End If
End Sub

Private Sub mnuRemoveItem_Click()
    
    If MsgBox(IIf((nTree.SelectedNode.Children > 0), "WARNING!!! All sub items under this item will be removed if you continue." & vbCrLf & vbCrLf, vbNullString) & "Are you sure you want to remove the selected item?", vbYesNo + vbQuestion) = vbYes Then
    
        nTree.SelectedItem(True).RemoveItem nTree.SelectedItem.Key
        
        nForm.Changed = True
        
        RefreshForm
    End If
End Sub

Private Sub mnuTree_Click()
    EnableTreeMenu
End Sub

Private Sub nForm_CloseFile(Cancel As Boolean)
    Set nBase = Nothing
End Sub

Private Sub nForm_Refresh()
    RefreshForm
End Sub

Private Sub nTree_DblClick()
    mnuItemProperties_Click
End Sub

Private Sub nTree_KeyDown(KeyCode As Integer, Shift As Integer)
    If nForm.Loaded And (Shift = 1) Then
        Select Case KeyCode
            Case 38
                mnuMove_Click 0
                KeyCode = 0
            Case 40
                mnuMove_Click 1
                KeyCode = 0
        End Select
    End If
    
End Sub

Private Sub nTree_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        EnableItemMenu
        Me.PopupMenu mnuItem
    End If
End Sub

Private Sub Form_Load()
    Set Ini = New clsIniFile
    Ini.LoadIniFile AppPath & "CreataTree.ini", DefaultIni
    
    Set nForm = New clsFileForm
    With nForm
        Set .Form = Me

        .FileExt = TreeFileExt
        .Filter = "Creata-Tree (*" & TreeFileExt & ")" & Chr(0) & "*" & TreeFileExt & Chr(0) & .Filter
        .FilePath = MyTreeFolder
        .FileExt = TreeFileExt
        .Caption = "Tree"

    End With
    
    mnuAbout.Tag = CBool(False)
    RefreshForm
    
    RefreshRecent
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Ini.SaveIniFile AppPath & "CreataTree.ini"
    Set Ini = Nothing
    
    Set nForm = Nothing
    Set nBase = Nothing
    
    End
End Sub

Private Sub mnuAbout_Click()
    
    mnuAbout.Tag = CBool(True)
    RefreshBrowser

End Sub

Private Sub mnuNeotext_Click()
    RunFile AppPath & "Neotext.org.url"
End Sub

Private Sub mnuNew_Click()
    nForm.NewFile
End Sub

Private Sub mnuNewSubItem_Click()
    If Ini.Setting("PromptForTemplate") Then
    
        Select Case frmNew.SetupItem(nBase)
            Case 0
                LoadItem ReadFile(BlankItemFile)
            Case 1
                LoadItem nBase.SubItem(frmNew.Template).XMLText
            Case Else
                frmNew.Show 1, Me
                If frmNew.IsOk Then
                    LoadItem nBase.SubItem(frmNew.Template).XMLText
                End If
        End Select
        
        Unload frmNew
    Else
        LoadItem ReadFile(BlankItemFile)
    End If
End Sub

Private Sub mnuOpen_Click()
    nForm.OpenFile
    AddRecent nForm.FileName
End Sub

Private Sub mnuSave_Click()
    nForm.SaveFile False
    AddRecent nForm.FileName
End Sub

Private Sub mnuSaveAs_Click()
    nForm.SaveFile True
    AddRecent nForm.FileName
End Sub

Private Sub nForm_NewFile(Cancel As Boolean)
    If Ini.Setting("PromptForTemplate") Then
        Select Case frmNew.SetupTree
            Case 0
                LoadTree BlankTreeFile
            Case 1
                LoadTree TemplateFolder & frmNew.Template & TreeFileExt
            Case Else
                frmNew.Show 1, Me
                If frmNew.IsOk Then
                    LoadTree TemplateFolder & frmNew.Template & TreeFileExt
                Else
                    Cancel = True
                End If
        End Select
        Unload frmNew
    Else
        LoadTree BlankTreeFile
    End If
End Sub

Private Function LoadTree(ByVal FileName As String)
    Set nBase = New clsItem
    OpenXMLFile nBase, FileName
End Function
Private Function LoadItem(ByVal XMLText As String) As Boolean
    
    Dim nItem As clsItem
    Set nItem = nTree.SelectedItem.AddItem(XMLText)
    nItem.Value("Opened") = True
    
    nForm.Changed = True
    RefreshForm nItem.Key
    Set nItem = Nothing
End Function

Private Sub nForm_OpenFile()
    LoadTree nForm.FileName
    AddRecent nForm.FileName
End Sub

Private Sub nForm_SaveFile()
    SaveXMLFile nBase, nForm.FileName
    AddRecent nForm.FileName
End Sub

Private Function RefreshBrowser()
    If CBool(mnuAbout.Tag) Or (Not nForm.Loaded) Then
        WebBrowser1.Navigate EngineFolder & "About\About.html"
    ElseIf nForm.Loaded Then
        If Ini.Setting("ViewPreviewPane") Then
            
            WriteHTMLText nBase, EngineFolder, False, False
            WebBrowser1.Navigate EngineFolder & "tree.html"
        End If
    End If

    EnableForm
End Function

Public Function RefreshForm(Optional ByVal nSelectedKey As String = vbNullString)
    If Not (nBase Is Nothing) Then
        If nSelectedKey = vbNullString Then nSelectedKey = nTree.SelectedKey
        
        nBase.Value("Label") = Replace(GetFileName(nForm.FileName), TreeFileExt, vbNullString)
        Me.Caption = AppName & " - [" & nBase.Value("Label") & "]" & IIf(nForm.Changed, "*", vbNullString)
        nTree.Refresh nBase, False, Ini.Setting("ViewGraphicalEdit")
        
        nTree.SelectedKey = nSelectedKey
    Else
        Me.Caption = AppName
    End If
    RefreshBrowser
End Function

Public Sub RefreshRecent()

    mnuRecent(0).Enabled = Not (Ini.Setting("Recent0") = "")
    mnuRecent(0).Caption = IIf(mnuRecent(0).Enabled, GetFileName(Ini.Setting("Recent0")), "(No Recent Files)")
    mnuRecent(0).Tag = Ini.Setting("Recent0")

    mnuRecent(1).Visible = Not (Ini.Setting("Recent1") = "")
    mnuRecent(1).Caption = IIf(mnuRecent(1).Visible, GetFileName(Ini.Setting("Recent1")), "")
    mnuRecent(1).Tag = Ini.Setting("Recent1")
    
    mnuRecent(2).Visible = Not (Ini.Setting("Recent2") = "")
    mnuRecent(2).Caption = IIf(mnuRecent(2).Visible, GetFileName(Ini.Setting("Recent2")), "")
    mnuRecent(2).Tag = Ini.Setting("Recent2")
    
    mnuRecent(3).Visible = Not (Ini.Setting("Recent3") = "")
    mnuRecent(3).Caption = IIf(mnuRecent(3).Visible, GetFileName(Ini.Setting("Recent3")), "")
    mnuRecent(3).Tag = Ini.Setting("Recent3")

End Sub

Public Sub AddRecent(ByVal fName As String)

    If PathExists(fName, True) Then
    
        Dim mnu
        For Each mnu In mnuRecent
            If LCase(Trim(mnu.Tag)) = LCase(Trim(fName)) Then
                Exit Sub
            End If
        Next
       
        Ini.Remove "Recent3"
        Ini.Add "Recent3", Ini.Setting("Recent2")
            
        Ini.Remove "Recent2"
        Ini.Add "Recent2", Ini.Setting("Recent1")

        Ini.Remove "Recent1"
        Ini.Add "Recent1", Ini.Setting("Recent0")
        
        Ini.Remove "Recent0"
        Ini.Add "Recent0", fName
        
    End If
    
    RefreshRecent
End Sub

Public Function EnableTreeMenu()
    mnuClose.Enabled = nForm.Loaded
    mnuSave.Enabled = nForm.Loaded
    mnuSaveAs.Enabled = nForm.Loaded
    mnuProperties.Enabled = nForm.Loaded
    mnuExport.Enabled = nForm.Loaded
End Function
Public Function EnableItemMenu()
    Dim IsEnabled As Boolean
    Dim IsBase As Boolean
    Dim IsFirst As Boolean
    Dim IsLast As Boolean
    IsEnabled = (nForm.Loaded And (Not (nTree.SelectedItem Is Nothing)))
    If nTree.SelectedItem Is Nothing Then
        IsEnabled = False
        IsBase = False
    Else
        IsEnabled = nForm.Loaded
        IsBase = (IsEnabled And (Not nTree.SelectedItem.IsBase))
        IsFirst = (IsBase And (Not nTree.SelectedNode.Key = nTree.SelectedNode.FirstSibling.Key))
        IsLast = (IsBase And (Not nTree.SelectedNode.Key = nTree.SelectedNode.LastSibling.Key))
    End If
    
    mnuNewSubItem.Enabled = IsEnabled
    mnuCopyItem.Enabled = IsBase
    mnuCutItem.Enabled = IsBase
    mnuPasteItem.Enabled = (IsEnabled And (Left(Clipboard.GetText, 12) = "Creata-Tree:"))
    mnuMove(0).Enabled = IsFirst
    mnuMove(1).Enabled = IsLast
    mnuRemoveItem.Enabled = IsBase
    mnuItemProperties.Enabled = IsEnabled

End Function
Public Function EnableOptionMenu()
    mnuPreviewPane.Checked = Ini.Setting("ViewPreviewPane")
    mnuGraphicalEditing.Checked = Ini.Setting("ViewGraphicalEdit")
    mnuForceDimensions.Checked = Ini.Setting("ForceDimensions")
    mnuPromptForTemplate.Checked = Ini.Setting("PromptForTemplate")
    mnuAlwaysUseBlank.Checked = Not Ini.Setting("PromptForTemplate")

End Function
Public Function EnableForm()
    EnableOptionMenu
    
    On Error Resume Next
    
    nTree.Top = 0
    nTree.Left = 0
    
    If (Not CBool(mnuAbout.Tag)) And nForm.Loaded Then
        If Not nTree.Visible Then nTree.Visible = True
        If Not WebBrowser1.Visible = Ini.Setting("ViewPreviewPane") Then WebBrowser1.Visible = Ini.Setting("ViewPreviewPane")
        If Ini.Setting("ViewPreviewPane") Then
            nTree.Width = (Me.ScaleWidth / 2)
            WebBrowser1.Top = 0
            WebBrowser1.Left = nTree.Width
            WebBrowser1.Width = nTree.Width
        Else
            nTree.Width = Me.ScaleWidth
        End If
        nTree.Height = Me.ScaleHeight
        WebBrowser1.Height = Me.ScaleHeight
    Else
        If nTree.Visible Then nTree.Visible = False
        If Not WebBrowser1.Visible Then WebBrowser1.Visible = True
        WebBrowser1.Top = 0
        WebBrowser1.Left = 0
        WebBrowser1.Width = Me.ScaleWidth
        WebBrowser1.Height = Me.ScaleHeight
    End If
    
    If nForm.Loaded And CBool(mnuAbout.Tag) Then
        mnuAbout.Tag = CBool(False)
    End If
    
    If Err Then Err.Clear
    On Error GoTo 0
End Function

Private Sub WebBrowser1_TitleChange(ByVal Text As String)
    If LCase(Trim(Text)) = "hide" Then
        mnuAbout.Tag = CBool(False)
        RefreshBrowser
    End If
End Sub
