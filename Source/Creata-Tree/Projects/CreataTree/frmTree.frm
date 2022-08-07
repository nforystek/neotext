VERSION 5.00
Begin VB.Form frmTree 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tree Properties"
   ClientHeight    =   6030
   ClientLeft      =   8310
   ClientTop       =   3390
   ClientWidth     =   6465
   Icon            =   "frmTree.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   6465
   Visible         =   0   'False
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Templates"
      Enabled         =   0   'False
      Height          =   2640
      Left            =   6345
      TabIndex        =   21
      Top             =   45
      Visible         =   0   'False
      Width           =   3090
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delete"
         Height          =   315
         Index           =   2
         Left            =   2175
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2175
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Edit"
         Height          =   315
         Index           =   1
         Left            =   1410
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2175
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "New"
         Height          =   315
         Index           =   0
         Left            =   645
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2175
         Width           =   735
      End
      Begin CreataTree.ctlNav ctlNav1 
         Height          =   1845
         Left            =   135
         TabIndex        =   9
         Top             =   255
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   3254
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Defaults"
      Height          =   2820
      Left            =   3285
      TabIndex        =   19
      Top             =   2730
      Width           =   3090
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   4
         Left            =   705
         TabIndex        =   15
         Top             =   1860
         Width           =   2190
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   3
         Left            =   705
         TabIndex        =   14
         Top             =   1455
         Width           =   2190
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         Left            =   705
         TabIndex        =   16
         Top             =   2265
         Width           =   2190
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   2790
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Default Font"
         Height          =   225
         Left            =   135
         TabIndex        =   31
         Top             =   1155
         Width           =   945
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Default Link Target"
         Height          =   225
         Left            =   135
         TabIndex        =   30
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Color:"
         Height          =   270
         Index           =   1
         Left            =   210
         TabIndex        =   24
         Top             =   1905
         Width           =   450
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Size:"
         Height          =   270
         Index           =   0
         Left            =   285
         TabIndex        =   23
         Top             =   2310
         Width           =   390
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Family:"
         Height          =   285
         Index           =   2
         Left            =   135
         TabIndex        =   22
         Top             =   1500
         Width           =   540
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   315
      Index           =   0
      Left            =   5430
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5625
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      Height          =   315
      Index           =   1
      Left            =   4470
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5625
      Width           =   915
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Base"
      Height          =   5505
      Left            =   90
      TabIndex        =   20
      Top             =   45
      Width           =   3090
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto Size Bullet Images"
         Height          =   240
         Left            =   120
         TabIndex        =   34
         Top             =   975
         Width           =   2505
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Save Images in Tree File"
         Height          =   240
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   2505
      End
      Begin CreataTree.cltMedia ctlMedia1 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   4935
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   503
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1755
         TabIndex        =   2
         Text            =   "Combo2"
         Top             =   2355
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   2205
         TabIndex        =   0
         Text            =   "5"
         Top             =   1365
         Width           =   600
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   2205
         TabIndex        =   1
         Text            =   "5"
         Top             =   1725
         Width           =   600
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   4275
         Width           =   2790
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2970
         Width           =   1110
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3390
         Width           =   1110
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Plus/Minus"
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   3420
         Value           =   1  'Checked
         Width           =   1590
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Treelines"
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   3015
         Value           =   1  'Checked
         Width           =   1590
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Background Image"
         Height          =   225
         Left            =   135
         TabIndex        =   32
         Top             =   4680
         Width           =   1500
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Background Color"
         Height          =   225
         Left            =   135
         TabIndex        =   29
         Top             =   4020
         Width           =   1500
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Note: The first tree item (the base, or root) is not visible in the exported tree."
         Height          =   465
         Left            =   210
         TabIndex        =   28
         Top             =   240
         Width           =   2745
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Height of each item:"
         Height          =   270
         Left            =   210
         TabIndex        =   27
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Top Left X Pixel Position:"
         Height          =   285
         Index           =   1
         Left            =   300
         TabIndex        =   26
         Top             =   1410
         Width           =   1770
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Top Left Y Pixel Position:"
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   25
         Top             =   1755
         Width           =   1800
      End
   End
   Begin VB.Image Image1 
      Height          =   2430
      Left            =   3405
      Picture         =   "frmTree.frx":08CA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2955
   End
End
Attribute VB_Name = "frmTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public IsOk As Boolean
Private nItem As New clsItem

Public Property Let XMLText(ByVal NewValue As String)
    nItem.XMLText = NewValue
    
    With nItem
        Text1(1).Text = .Value("Left")
        Text1(0).Text = .Value("Top")
        
        Combo2.ListIndex = IsOnList(Combo2, .Value("Height"))
        
        Check1.Value = -CInt(.Value("SaveImages"))
        Check3.Value = -CInt(.Value("StretchBullets"))
    
        Check2(0).Value = -CInt(.Value("UseTreelines"))
        Check2(1).Value = -CInt(.Value("UsePlusMinus"))
        
        Combo1(0).ListIndex = IsOnList(Combo1(0), .Value("TreelineColor"))
        Combo1(1).ListIndex = IsOnList(Combo1(1), .Value("PlusMinusColor"))
        
        Text2(4).Text = .Value("BackColor")
        
        ctlMedia1.Value = .Value("BackImage")
        
        Text2(3).Text = .Value("LinkTarget")
        
        Combo1(3).Text = .Value("FontFamily")
        Combo1(4).Text = .Value("FontColor")
        Combo1(2).Text = .Value("FontSize")
    
    End With

    EnableForm

End Property
Public Property Get XMLText() As String
    
    With nItem
    
        .Value("Left") = Text1(1).Text
        .Value("Top") = Text1(0).Text
        
        .Value("Height") = Combo2.Text
    
        .Value("UseTreelines") = CBool(Check2(0).Value)
        .Value("UsePlusMinus") = CBool(Check2(1).Value)
        
        .Value("SaveImages") = CBool(Check1.Value)
        .Value("StretchBullets") = CBool(Check3.Value)
        
        .Value("TreelineColor") = Combo1(0).Text
        .Value("PlusMinusColor") = Combo1(1).Text
        
        .Value("BackColor") = Text2(4).Text
        
        .Value("BackImage") = ctlMedia1.Value
        
        .Value("LinkTarget") = Text2(3).Text
        
        .Value("FontFamily") = Combo1(3).Text
        .Value("FontColor") = Combo1(4).Text
        .Value("FontSize") = Combo1(2).Text
    
    End With
    
    
    XMLText = nItem.XMLText
    
End Property

Private Sub Check2_Click(Index As Integer)
    EnableForm
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 1
            IsOk = True
            Me.Hide
        Case 0
            IsOk = False
            Me.Hide
    End Select
End Sub

Private Sub Form_Load()
    InitFolderColor "Treelines", Combo1(0)
    InitFolderColor "PlusMinus", Combo1(1)
    InitItemHeight Combo2
    InitFontFamily Combo1(3)
    InitFontColor Combo1(4)
    InitFontSize Combo1(2)
    
    ctlMedia1.Inititialize "Backgrounds", "0x0"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = True
        IsOk = False
        Me.Hide
    End If
End Sub

Public Function EnableForm()
    Combo1(0).Enabled = CBool(Check2(0).Value)
    Combo1(0).BackColor = IIf(Combo1(0).Enabled, &HFFFFFF, &HE0E0E0)

    Combo1(1).Enabled = CBool(Check2(1).Value)
    Combo1(1).BackColor = IIf(Combo1(1).Enabled, &HFFFFFF, &HE0E0E0)

End Function

