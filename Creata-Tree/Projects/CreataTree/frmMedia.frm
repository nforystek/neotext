VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMedia 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Media Library"
   ClientHeight    =   4575
   ClientLeft      =   8445
   ClientTop       =   3840
   ClientWidth     =   5760
   Icon            =   "frmMedia.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5760
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2520
      ScaleHeight     =   270
      ScaleWidth      =   1290
      TabIndex        =   9
      Top             =   4185
      Width           =   1350
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Delete"
      Height          =   315
      Index           =   2
      Left            =   4860
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   60
      Width           =   810
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      Height          =   315
      Index           =   4
      Left            =   4860
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4185
      Width           =   810
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Select"
      Default         =   -1  'True
      Height          =   315
      Index           =   3
      Left            =   3990
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4185
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   2280
      TabIndex        =   7
      Top             =   135
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   75
      Width           =   2100
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3645
      Left            =   75
      TabIndex        =   4
      Top             =   450
      Width           =   5610
      _ExtentX        =   9895
      _ExtentY        =   6429
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   6244
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Format"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&View"
      Height          =   315
      Index           =   1
      Left            =   3990
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Width           =   810
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Import"
      Height          =   315
      Index           =   0
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      Width           =   810
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   -165
      TabIndex        =   8
      Top             =   4185
      Width           =   2565
   End
   Begin VB.Menu mnuMedia 
      Caption         =   "Media"
      Visible         =   0   'False
      Begin VB.Menu mnuView 
         Caption         =   "View"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImport 
         Caption         =   "Import"
      End
   End
End
Attribute VB_Name = "frmMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public IsOk As Boolean

Public Property Get Dimensions() As String
    Dimensions = ListView1.SelectedItem.ListSubItems(1).Tag
End Property

Public Property Get Value() As String
    Value = Combo1.Text & "\" & ListView1.SelectedItem.Tag
End Property
Public Property Let Library(ByVal NewValue As String)
    Combo1.ListIndex = IsOnList(Combo1, NewValue)
End Property
Public Property Let Selected(ByVal NewValue As String)
    Combo1.Enabled = False
    Command1(3).Visible = True
    Command1(4).Caption = "&Cancel"
    ListView1.Tag = SafeKey(GetFileTitle(NewValue))
        
    Dim nInfo As ImageInfoType
    nInfo = GetImageInfo(MediaFolder & NewValue)
    Label1.Caption = nInfo.Title & " " & nInfo.Desc
End Property

Private Property Get Selecting() As Boolean
    Selecting = Command1(3).Visible
End Property

Public Property Get ForcedDimensions() As String
    ForcedDimensions = Command1(3).Tag
End Property
Public Property Let ForcedDimensions(ByVal NewValue As String)
    Command1(3).Tag = NewValue
End Property

Private Sub RefreshMedia()
    
    ListView1.ListItems.Clear
    File1.Path = MediaFolder & Combo1.Text
    
    AddMedia ".gif"
    AddMedia ".jpg"
    AddMedia ".bmp"
    AddMedia ".png"
    
    If NodeExists(ListView1.ListItems, ListView1.Tag) Then
        ListView1.ListItems(ListView1.Tag).Selected = True
    Else
        If ListView1.ListItems.Count > 0 Then
            ListView1.ListItems(1).Selected = True
        End If
    End If
    
    RefreshPreview
    
    EnableForm
End Sub
Private Sub RefreshPreview()
    If Not ListView1.SelectedItem Is Nothing Then
        Picture1.Picture = LoadPicture(MediaFolder & Value)
        
    End If
End Sub
Private Sub AddMedia(ByVal Ext As String)
    Dim cnt As Long
    Dim nInfo As ImageInfoType
    Dim nList As ListItem
    Dim nSubList As ListSubItem
    File1.Pattern = "*" & Ext
    If File1.ListCount > 0 Then
        For cnt = 0 To File1.ListCount - 1
            nInfo = GetImageInfo(File1.Path & "\" & File1.List(cnt))
            
            With nInfo
                Set nList = ListView1.ListItems.Add(, SafeKey(.Title), .Title)
                nList.Tag = .name
            
                Set nSubList = nList.ListSubItems.Add(, , .Desc)
                nSubList.Tag = .Dims
            
                Set nSubList = Nothing
                Set nList = Nothing
            End With
        Next
    End If

End Sub

Private Sub Combo1_Click()
    RefreshMedia
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            mnuImport_Click
        Case 1
            mnuView_Click
        Case 2
            mnuDelete_Click
        Case 3
            If Not ListView1.SelectedItem Is Nothing Then
                If (GetHeight(ForcedDimensions) = 0 Or (GetHeight(ForcedDimensions) = GetHeight(Dimensions))) Or (Not Ini.Setting("ForceDimensions")) Then
                    If (GetWidth(ForcedDimensions) = 0 Or (GetWidth(ForcedDimensions) = GetWidth(Dimensions))) Or (Not Ini.Setting("ForceDimensions")) Then
                        IsOk = True
                        Me.Hide
                    Else
                        MsgBox "The width of the image you whish to select must be " & GetWidth(ForcedDimensions) & " pixels for the current tree." & vbCrLf & "This is defined by the 'Height of Each Item' value in the main tree properties.", vbInformation
                    End If
                Else
                    MsgBox "The height of the image you whish to select must be " & GetHeight(ForcedDimensions) & " pixels for the current tree." & vbCrLf & "This is defined by the 'Height of Each Item' value in the main tree properties.", vbInformation
                End If
            Else
                MsgBox "You must highlight an image from the list above to select it." & vbCrLf & "If not images are displayed, you must import them using the import button above.", vbInformation
            End If
        Case 4
            If Selecting Then
                IsOk = False
                Me.Hide
            Else
                Unload Me
            End If
    End Select
End Sub

Private Sub Form_Activate()
    RefreshMedia

End Sub

Private Sub Form_Load()
    Combo1.AddItem "Bullet Icons"
    Combo1.AddItem "Menu Pictures"
    Combo1.AddItem "Backgrounds"
    Combo1.ListIndex = 0
    
    ForcedDimensions = "0x0"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        If Selecting Then
            IsOk = False
            Cancel = True
            Me.Hide
        End If
    End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ColumnSortClick ListView1, ColumnHeader
End Sub

Private Sub ListView1_DblClick()
    If Selecting Then
        Command1_Click 3
    Else
        mnuView_Click
    End If
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Picture1.Picture = LoadPicture(MediaFolder & Value)
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        EnableForm
        Me.PopupMenu mnuMedia
    End If
End Sub

Private Sub mnuDelete_Click()
    If MsgBox("WARNING: Deleting this file will cause missing file references in trees that use this file." & vbCrLf & vbCrLf & "Are you sure you want to delete this file?", vbYesNo + vbQuestion) = vbYes Then
        On Error Resume Next
        Kill MediaFolder & Value
        If Err Then Err.Clear
        On Error GoTo 0
        RefreshMedia
    End If
End Sub

Private Sub mnuImport_Click()
    Dim sOpen As SelectedFile

    FileDialog.sFilter = "BMP Image Format (*.BMP)" & Chr(0) & "*.bmp" & Chr(0) & "GIF Image Format (*.GIF)" & Chr(0) & "*.gif" & Chr(0) & "JPG Image Format (*.JPG)" & Chr(0) & "*.jpg" & Chr(0) & "PNG Image Format (*.PNG)" & Chr(0) & "*.png" & Chr(0) & "All Files (*.*)" & Chr(0) & "*.*"
    FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
    FileDialog.sDlgTitle = "Import Images"
    FileDialog.sInitDir = AppPath
    
    On Error Resume Next
    sOpen = ShowOpen(Me.hWnd)
    If (Not (Err.Number = 32755)) And (sOpen.bCanceled = False) Then
        Dim errMsg As String
        Dim cnt As Integer
        For cnt = 1 To sOpen.nFilesSelected
        
            errMsg = errMsg & ImportImageFile(sOpen.sLastDirectory & "\" & sOpen.sFiles(cnt))
        Next
        
        RefreshMedia
        If Not errMsg = vbNullString Then
            MsgBox IIf((sOpen.nFilesSelected = 1), "There was a problem importing the image.", "There was problems importing the images.") & vbCrLf & errMsg, vbInformation
        End If
    End If

    If Err Then Err.Clear
    On Error GoTo 0

End Sub
Private Function ImportImageFile(ByVal FileName As String) As String
    Dim nInfo As ImageInfoType
    nInfo = GetImageInfo(FileName)
    If nInfo.Valid Then
        If Not PathExists(MediaFolder & Combo1.Text & "\" & GetFileName(FileName), True) Then
            ImportImageFile = vbNullString
            On Error Resume Next
            FileCopy FileName, MediaFolder & Combo1.Text & "\" & GetFileName(FileName)
            If Err Then
                ImportImageFile = GetFileName(FileName) & " - Error copying image: " & Err.Description & vbCrLf
                Err.Clear
            End If
            On Error GoTo 0
        Else
            ImportImageFile = GetFileName(FileName) & " - Image already exists in media library." & vbCrLf
        End If
    Else
        ImportImageFile = GetFileName(FileName) & " - Unrecognized or invalid image file." & vbCrLf
    End If
End Function

Private Sub mnuView_Click()
    frmView.LoadImage MediaFolder & Value
    frmView.Show 1, Me
End Sub

Private Sub mnuMedia_Click()
    EnableForm
End Sub

Private Function EnableForm()
    mnuView.Enabled = (Not (ListView1.SelectedItem Is Nothing))
    mnuDelete.Enabled = (Not (ListView1.SelectedItem Is Nothing))
    Command1(1).Enabled = (Not (ListView1.SelectedItem Is Nothing))
    Command1(2).Enabled = (Not (ListView1.SelectedItem Is Nothing))
End Function


