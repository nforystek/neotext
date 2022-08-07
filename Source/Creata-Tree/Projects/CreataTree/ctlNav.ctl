VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl ctlNav 
   ClientHeight    =   3735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5145
   ScaleHeight     =   3735
   ScaleWidth      =   5145
   ToolboxBitmap   =   "ctlNav.ctx":0000
   Begin MSComctlLib.TreeView nView 
      Height          =   2025
      Left            =   345
      TabIndex        =   0
      Top             =   195
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3572
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   557
      LabelEdit       =   1
      Style           =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList nImages 
      Left            =   405
      Top             =   2565
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   16711935
      _Version        =   393216
   End
End
Attribute VB_Name = "ctlNav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event DblClick()

Public Property Get Graphical() As Boolean
    Graphical = CBool(nImages.Tag)
End Property
Public Property Let Graphical(ByVal NewValue As Boolean)
    nImages.Tag = CBool(NewValue)
End Property

Public Property Get SelectedItem(Optional ByVal GetParent As Boolean = False) As clsItem
    If SelectedNode(GetParent) Is Nothing Then
        Set SelectedItem = Nothing
    Else
        Set SelectedItem = SelectedNode(GetParent).Tag
    End If
End Property

Public Property Get SelectedNode(Optional ByVal GetParent As Boolean = False) As INode
    If nView.SelectedItem Is Nothing Then
        Set SelectedNode = Nothing
    Else
        If GetParent Then
            If nView.SelectedItem.Parent Is Nothing Then
                Set SelectedNode = Nothing
            Else
                Set SelectedNode = nView.SelectedItem.Parent
            End If
        Else
            Set SelectedNode = nView.SelectedItem
        End If
    End If
End Property

Public Property Get SelectedKey() As String
    If SelectedItem Is Nothing Then
        SelectedKey = vbNullString
    Else
        SelectedKey = SelectedItem.Key
    End If
End Property

Public Property Let SelectedKey(ByVal NewValue As String)
    If NodeExists(nView.Nodes, NewValue) Then
        Set nView.SelectedItem = nView.Nodes(NewValue)
        nView.SelectedItem.EnsureVisible
    ElseIf nView.Nodes.Count > 0 Then
        Set nView.SelectedItem = nView.Nodes(1)
        nView.SelectedItem.EnsureVisible
    End If
End Property

Public Function Refresh(ByRef nBase As clsItem, Optional ByVal ShowTemplates As Boolean = False, Optional IsGraphical As Boolean = True)

    Graphical = IsGraphical

    nView.Nodes.Clear
    Set nView.ImageList = Nothing
    
    If Graphical Then
    
        nImages.ListImages.Clear
        nImages.ImageHeight = 16
        nImages.ImageWidth = 16
        
        nImages.ListImages.Add , BlankImageKey, LoadPicture(EngineFolder & "Media\202.gif")
        nImages.ListImages.Add , ErrorImageKey, LoadPicture(EngineFolder & "Media\404.gif")
        
        If nBase.IsBase Then
            If nBase.Value("UseTreelines") And nBase.Value("UsePlusMinus") Then
                nView.Style = 7
            ElseIf nBase.Value("UseTreelines") Then
                nView.Style = 5
            ElseIf nBase.Value("UsePlusMinus") Then
                nView.Style = 3
            Else
                nView.Style = 1
            End If
        End If
        Set nView.ImageList = nImages
    Else
        nView.Style = 6
    End If
    
    AddNode nBase
    
    RefreshItem nBase, ShowTemplates
    
End Function
Private Function RefreshItem(ByRef Parent As clsItem, Optional ByVal ShowTemplates As Boolean = False)
    
    If Parent.SubItems.Count > 0 Then
        Dim nItem As clsItem
        For Each nItem In Parent.SubItems
            If ((ShowTemplates And nItem.IsTemplate) Or (Not nItem.IsTemplate)) And (Not nItem.IsBlank) Then
                AddNode nItem, Parent
                If Not ShowTemplates Then RefreshItem nItem
            End If
        Next
    End If
    
End Function

Private Function CacheBullet(ByRef nItem As clsItem, ByVal Expanded As Boolean)
    With nItem
        
        If Not NodeExists(nImages.ListImages, .BulletKey(Expanded)) Then
            nImages.ListImages.Add , nItem.BulletKey(Expanded), LoadPicture(MediaFolder & nItem.Bullet(Expanded))
        End If
    End With
    
End Function

Private Function AddNode(ByVal nItem As clsItem, Optional ByRef nParent As clsItem = Nothing)
    
    Dim nNode As INode
    If Graphical Then
        CacheBullet nItem, False
        CacheBullet nItem, True
        If nParent Is Nothing Then
            Set nNode = nView.Nodes.Add(, , nItem.Key, nItem.Display, nItem.BulletKey(nItem.Value("Opened")))
        Else
            Set nNode = nView.Nodes.Add(nParent.Key, tvwChild, nItem.Key, nItem.Display, nItem.BulletKey(nItem.Value("Opened")))
        End If
    
    Else
    
        If nParent Is Nothing Then
            Set nNode = nView.Nodes.Add(, , nItem.Key, nItem.Display)
        Else
            Set nNode = nView.Nodes.Add(nParent.Key, tvwChild, nItem.Key, nItem.Display)
        End If
    
    End If
    
    nNode.Expanded = nItem.Value("Opened")
    
    Set nNode.Tag = nItem
    Set nNode = Nothing
    
End Function

Private Sub nView_DblClick()
    If Not SelectedNode Is Nothing Then
        If SelectedNode.Children = 0 Then
            
            RaiseEvent DblClick
            
            SetBullet SelectedItem, Not SelectedItem.Value("Opened")

        End If
        
    End If
    
End Sub
Private Sub nView_Collapse(ByVal node As MSComctlLib.node)
    SetBullet node.Tag, False
End Sub

Private Sub nView_Expand(ByVal node As MSComctlLib.node)
    SetBullet node.Tag, True
End Sub
Private Sub SetBullet(ByRef nItem As clsItem, ByVal Expanded As Boolean)
    If Graphical Then
        If nItem.BulletUse(Expanded) Then
            
            nView.Nodes(nItem.Key).Image = nItem.BulletKey(Expanded)
        End If
    End If
    nItem.Value("Opened") = Expanded

End Sub

Private Sub nView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub nView_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub nView_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub nView_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    nView.Left = 0
    nView.Top = 0
    nView.Width = UserControl.ScaleWidth
    nView.Height = UserControl.ScaleHeight
   
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub UserControl_Terminate()
    Reset
End Sub

