VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProjectExplorer 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4770
   ControlBox      =   0   'False
   Icon            =   "frmProjectExplorer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MaxIDE.ctlDragger ctlDragger1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      Top             =   0
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   688
      Dockable        =   0   'False
      Movable         =   0   'False
      Docked          =   0   'False
      Caption         =   "Project Explorer"
      RepositionForm  =   0   'False
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   1740
      Left            =   360
      TabIndex        =   0
      Top             =   540
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   3069
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   530
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "frmProjectExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public SelectedNode As MSComctlLib.Node
Public Function NodeExists(ByVal nKey As String)
    On Error Resume Next
    Dim Node
    Set Node = TreeView1.Nodes(nKey)
    If Err = 0 Then
        NodeExists = True
    Else
        Err.Clear
        NodeExists = False
    End If
End Function
Public Function SelectNode(ByRef nNode As MSComctlLib.Node)
    If Not nNode.Selected Then nNode.Selected = True
    Set SelectedNode = nNode
End Function
Public Function IsFolder(ByRef nNode As MSComctlLib.Node)
    IsFolder = (nNode.Key = "Objects") Or (nNode.Key = "Project") Or (nNode.Key = "Root")
End Function
Public Function IsSelected() As Boolean
    If Not TreeView1.SelectedItem Is Nothing Then
        IsSelected = True
    Else
        IsSelected = False
    End If
End Function

Public Function IsObject(ByVal nClass As String) As Boolean
    Select Case nClass
        Case "MaxIDE.Project", "MaxIDE.Module"
            IsObject = False
        Case Else
            IsObject = True
    End Select
End Function

Public Function GetItemIconKey(ByVal nClass As String) As String
    Select Case nClass
        Case "MaxIDE.Project"
            GetItemIconKey = Project.Language
        Case "MaxIDE.Module", "MaxIDE.Generic"
            GetItemIconKey = ModuleFile
            
        Case Else
            
            GetItemIconKey = ObjectFile

    End Select
End Function

Public Function AllowOpen() As Boolean
    If IsSelected Then
        AllowOpen = (Not (SelectedNode.Key = "Root")) And (Not (SelectedNode.Key = "Objects")) And (Not (SelectedNode.Key = "Template"))
    Else
        AllowOpen = False
    End If
End Function

Public Function AllowDelete() As Boolean
    If IsSelected Then
        
        If IsFolder(SelectedNode) Then
            AllowDelete = False
        Else
            AllowDelete = (Not (SelectedNode.Key = "Template")) And (Not (SelectedNode.Key = "Name")) And (Not (SelectedNode.Key = "Source"))
        End If
    Else
        AllowDelete = False
    End If
End Function

Public Function AllowRename() As Boolean
    If IsSelected Then
        
        If IsFolder(SelectedNode) Then
            AllowRename = False
        Else
            AllowRename = Not (SelectedNode.Key = "Source")
        End If
    Else
        AllowRename = False
    End If
End Function

Public Function ChildExists(ByVal ParentIndex As Integer, ByVal ChildName As String) As Boolean
    Dim Node, exists As Boolean
    exists = False
    If frmProjectExplorer.TreeView1.Nodes(ParentIndex).children > 0 Then
        Set Node = frmProjectExplorer.TreeView1.Nodes(ParentIndex).child.FirstSibling
        Do Until Node.Index = frmProjectExplorer.TreeView1.Nodes(ParentIndex).LastSibling.Index Or Node.Next Is Nothing
            If ChildName = Node.Text Then
                exists = True
                Exit Do
            End If
            Set Node = Node.Next
        Loop
        If ChildName = Node.Text Then
            exists = True
        End If
    End If
    ChildExists = exists
End Function
Public Function GetNextName1(ByVal Prefix As String)
    
    If ChildExists(TreeView1.Nodes("Root").Index, Prefix) Then
        Dim Node, cnt As Integer, str As String
        cnt = 1
        Do Until (Not ChildExists(TreeView1.Nodes("Root").Index, Prefix & cnt))
            cnt = cnt + 1
        Loop
        GetNextName1 = Prefix & cnt
    Else
        GetNextName1 = Prefix
    End If
End Function
Public Function GetNextName2(ByVal Prefix As String)
    
    If ChildExists(TreeView1.Nodes("Project").Index, Prefix) Or ChildExists(TreeView1.Nodes("Objects").Index, Prefix) Then
        Dim cnt As Integer
        cnt = 1
        Do Until (Not ChildExists(TreeView1.Nodes("Project").Index, Prefix & cnt)) And (Not ChildExists(TreeView1.Nodes("Objects").Index, Prefix & cnt))
            cnt = cnt + 1
        Loop
        GetNextName2 = Prefix & cnt
    Else
        GetNextName2 = Prefix
    End If
End Function
Public Function RefreshProject()
    TreeView1.Nodes.Clear
    
    If Project.Loaded Then
        
        Dim cNode As MSComctlLib.Node
        Dim cItem As clsItem
        
        If Project.IsTemplate Then
        
            Set cNode = TreeView1.Nodes.Add(, , "Template", Project.Template.ItemClass, "FolderOpen")
            Set cNode.Tag = Project.Template
            
            Set cNode = TreeView1.Nodes.Add("Template", tvwChild, Project.Template.ItemName, Project.Template.ItemName, ObjectFile)
            Set cNode.Tag = Project.Template

            TreeView1.Nodes("Template").Expanded = True
            
            If frmMainIDE.Visible Then ShowScriptPage Project.Template.ItemName
            
        Else
        
            Set cNode = TreeView1.Nodes.Add(, , "Root", GetFileTitle(Project.FileName), "Root")
            AddItem Project.Items("Project")
            Set cNode = TreeView1.Nodes.Add("Root", tvwChild, "Objects", "Objects", "FolderClose")
            
            For Each cItem In Project.Items
                If Not cItem.ItemName = "Project" Then AddItem cItem
            Next
            
            TreeView1.Nodes("Root").Expanded = True
            TreeView1.Nodes("Project").Expanded = True
            TreeView1.Nodes("Objects").Expanded = True
            
            If frmMainIDE.Visible Then ShowScriptPage "Project"
        End If
        
        frmDebug.ResetDialect Project.Language, Project.Compiler
    End If
End Function

Public Function RefreshName()
    If NodeExists("Root") Then
        TreeView1.Nodes("Root").Text = GetFileTitle(Project.FileName)
    End If
End Function
Public Function AddItem(ByRef cItem As clsItem) As MSComctlLib.Node
    Dim cNode As MSComctlLib.Node
    Select Case cItem.ItemClass
        Case "MaxIDE.Project"
            cItem.ItemName = GetNextName1(cItem.ItemName)
            Set cNode = TreeView1.Nodes.Add("Root", tvwChild, cItem.ItemName, cItem.ItemName, GetItemIconKey(cItem.ItemClass))
        
        Case "MaxIDE.Module"
            cItem.ItemName = GetNextName2(cItem.ItemName)
            Set cNode = TreeView1.Nodes.Add("Project", tvwChild, cItem.ItemName, cItem.ItemName, GetItemIconKey(cItem.ItemClass))
        
        Case Else
            cItem.ItemName = GetNextName2(cItem.ItemName)
            Set cNode = TreeView1.Nodes.Add("Objects", tvwChild, cItem.ItemName, cItem.ItemName, GetItemIconKey(cItem.ItemClass))
            
    End Select

    Set cNode.Tag = cItem
    Set AddItem = cNode
    Set cNode = Nothing
End Function

Public Function DeleteItem(ByVal nKey As String)
    If TypeName(TreeView1.Nodes(nKey).Tag) = "clsItem" Then
        Set TreeView1.Nodes(nKey).Tag = Nothing
    End If
    TreeView1.Nodes.Remove nKey
End Function

Private Sub ctlDragger1_Resize(Left As Long, Top As Long, Width As Long, ByVal Height As Long)
    On Error Resume Next
    TreeView1.Move Left, Top, Width, Height
    Err.Clear
End Sub

Private Sub Form_Load()
    Set TreeView1.ImageList = frmMainIDE.CommonIcons
    
    With ctlDragger1
        .Docked = dbSettings.GetScriptingSetting("pDocked")
        
        .DockedWidth = dbSettings.GetScriptingSetting("pDockedWidth")
        .DockedHeight = dbSettings.GetScriptingSetting("pDockedHeight")
    
        .FloatingTop = dbSettings.GetScriptingSetting("pFloatTop")
        .FloatingLeft = dbSettings.GetScriptingSetting("pFloatLeft")
        .FloatingWidth = dbSettings.GetScriptingSetting("pFloatWidth")
        .FloatingHeight = dbSettings.GetScriptingSetting("pFloatHeight")
    
        .SetupDockedForm frmMainIDE, Me, 3
    End With
End Sub

Public Function ShowScriptPage(ByVal pName As String, Optional ByVal LineNum As Long = 0, Optional ByVal LineColNum As Long = 0)
    Dim IVX As Long
    
    If frmMainIDE.Visible Then
        If Not (GetTextObj Is Nothing) Then
            IVX = GetTextObj.Parent.WindowState
            IVX = IVX + 1
        End If
    End If
        
    SelectNode TreeView1.Nodes(pName)
    
    Dim frm As frmScriptPage
    Set frm = GetScriptPage(pName)
    If frm Is Nothing Then
    
        Set frm = New frmScriptPage
        Load frm
        
        frm.Language = Project.Language
        Set frm.Item = SelectedNode.Tag

    End If
    
    If frmMainIDE.Visible Then
    
        If Not (GetTextObj Is Nothing) And (IVX > 0) Then
            frm.WindowState = (IVX - 1)
        End If
        
        frm.ZOrder 0

    End If
    
    If LineNum > 0 Then

        frm.CodeEdit1.SelectRow LineNum, False
        If (LineColNum > 0) Then frm.CodeEdit1.SelStart = frm.CodeEdit1.SelStart + LineColNum
        If Not (frm.CodeEdit1.ErrorLine = LineNum) Then frm.CodeEdit1.ErrorLine = LineNum
'        frm.CodeEdit1.colorerror = frm.CodeEdit1.ConvertColor(GetCollectSkinValue("script_errorcolor"))
'        If Not (frm.CodeEdit1.colorerror = frm.CodeEdit1.ConvertColor(GetCollectSkinValue("script_errorcolor"))) Then
'            frm.CodeEdit1.colorerror = frm.CodeEdit1.ConvertColor(GetCollectSkinValue("script_errorcolor"))
'        End If
    End If

    Set frm = Nothing
End Function
Public Function PageIsOpen(ByVal pName As String) As Boolean
    PageIsOpen = False
    Dim frm
    For Each frm In Forms
       
        If TypeName(frm) = "frmScriptPage" Then
            If frm.Item.ItemName = pName Then
                PageIsOpen = True
            End If
        ElseIf TypeName(frm) = "frmIEPage" Then
            If frm.Caption = pName Then
                PageIsOpen = True
            End If
        End If
    Next
End Function
Public Function GetScriptPage(ByVal pName As String) As frmScriptPage
    Dim frm
    For Each frm In Forms
        If TypeName(frm) = "frmScriptPage" Then
            If frm.Item.ItemName = pName Then
                Set GetScriptPage = frm
                Exit For
            End If
        End If
    Next
End Function

Public Function GetInternetPage(ByVal pName As String) As frmIEPage
    Dim frm
    For Each frm In Forms
        If TypeName(frm) = "frmIEPage" Then
            If frm.Caption = pName Then
                Set GetInternetPage = frm
                Exit For
            End If
        End If
    Next
End Function
Private Sub UpdateScriptedItem(ByVal NewString As String)
    Dim pItem As New clsItem
    If Not Project.IsTemplate Then
        Set pItem = Project.Items(SelectedNode.Tag.ItemName)
        Project.Items.Remove SelectedNode.Tag.ItemName
        Project.Items.Add pItem, NewString
    End If
    
    Dim frm As frmScriptPage
    Set frm = GetScriptPage(SelectedNode.Tag.ItemName)
    If Not frm Is Nothing Then
        frm.Caption = NewString
        frm.Item.ItemName = NewString
        RefreshWindowMenu
    End If
    SelectedNode.Tag.ItemName = NewString
    SelectedNode.Key = NewString
    
    frmMainIDE.ProjectSet True
End Sub
Private Sub TreeView1_AfterLabelEdit(Cancel As Integer, NewString As String)
    If Not IsAlphaNumeric(NewString) Then
        MsgBox "You must specify a name that is alpha numeric with no spaces.", vbInformation, AppName
        Cancel = True
    Else
        If IsSelected Then
            If Project.IsTemplate And (SelectedNode.Parent Is Nothing) Then
                SelectedNode.Tag.ItemClass = NewString
                frmMainIDE.ProjectSet True
                
            ElseIf Project.IsTemplate Then
                UpdateScriptedItem NewString
            Else
                If NodeExists("Project") Then
                    If ChildExists(TreeView1.Nodes("Project").Index, NewString) Or LCase(Trim(NewString)) = "project" Or ChildExists(TreeView1.Nodes("Objects").Index, NewString) Then
                        MsgBox "That name already exists, please choose another.", vbInformation, AppName
                        Cancel = True
                    Else
                        UpdateScriptedItem NewString
                    End If
                Else
                    UpdateScriptedItem NewString
                End If
            End If
        End If
    End If
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)
    Cancel = BeginRename(True)
End Sub
Public Function BeginRename(ByVal LabelEdited As Boolean) As Boolean
    Dim Cancel As Boolean
    Cancel = Not AllowRename
    If Not Cancel And Not LabelEdited Then
        TreeView1.StartLabelEdit
    End If
    If Cancel Then
        MsgBox "This item can not be renamed.", vbInformation, AppName
    End If
    BeginRename = Cancel
End Function

Public Sub OpenItem()
    If IsSelected Then

        If (Not ((SelectedNode.Key = "Objects") Or (SelectedNode.Key = "Root"))) And (Not (SelectedNode.Key = "Template")) Then
            ShowScriptPage SelectedNode.Key
        End If
        
    End If
End Sub
Private Sub TestNode(ByVal Key As String)
    If Not (SelectedNode Is Nothing) Then
        If SelectedNode.Key = Key Then
            TreeView1.Nodes(Key).Expanded = Not TreeView1.Nodes(Key).Expanded
        End If
    End If
End Sub
Private Sub TreeView1_DblClick()
    TestNode "Project"
    TestNode "Objects"
    
    OpenItem
End Sub

Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)

    If NodeExists("Root") Then
        If Not TreeView1.Nodes("Root").Expanded Then TreeView1.Nodes("Root").Expanded = True
    ElseIf NodeExists("Template") Then
        If Not TreeView1.Nodes("Template").Expanded Then TreeView1.Nodes("Template").Expanded = True
    End If

    Select Case Node.Image
        Case "FolderOpen"
            Node.Image = IIf(Node.Expanded, "FolderOpen", "FolderClose")
    End Select

End Sub

Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)

    Select Case Node.Image
        Case "FolderClose"
            Node.Image = IIf(Node.Expanded, "FolderOpen", "FolderClose")
    End Select
End Sub

Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        RefreshScriptMenu
        If Not (SelectedNode Is Nothing) Then
            If SelectedNode.Key = "Root" Then
                Me.PopupMenu frmMainIDE.mnuProject
            Else
                Me.PopupMenu frmMainIDE.mnuItem
            End If
        End If
    End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    SelectNode Node
End Sub


