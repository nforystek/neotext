VERSION 5.00
Begin VB.Form frmItem 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Properties"
   ClientHeight    =   6780
   ClientLeft      =   6510
   ClientTop       =   3015
   ClientWidth     =   6795
   Icon            =   "frmItem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bullet"
      Height          =   2805
      Left            =   90
      TabIndex        =   29
      Top             =   450
      Width           =   3225
      Begin CreataTree.cltMedia ctlMedia2 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1380
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   503
      End
      Begin CreataTree.cltMedia ctlMedia1 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   555
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   503
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Collapsed Bullet"
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Width           =   1905
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Expanded Bullet"
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   1095
         Width           =   1890
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Note: Bullet height and width should be the same as the 'Height of Each Item' value in the 'Tree Properties'."
         Height          =   615
         Left            =   225
         TabIndex        =   30
         Top             =   1965
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pictures"
      Height          =   2805
      Left            =   3435
      TabIndex        =   27
      Top             =   450
      Width           =   3285
      Begin CreataTree.cltMedia ctlMedia4 
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   1380
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   503
      End
      Begin CreataTree.cltMedia ctlMedia3 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   555
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   503
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mouse Out Picture"
         Height          =   270
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   270
         Width           =   2925
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mouse Over Picture"
         Height          =   270
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   1095
         Width           =   2925
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Note: The height should be the same as the 'Height of Each Item' value in the 'Tree Properties'."
         Height          =   615
         Left            =   225
         TabIndex        =   28
         Top             =   1965
         Width           =   2655
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Text"
      Height          =   3390
      Left            =   90
      TabIndex        =   23
      Top             =   3300
      Width           =   3225
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         Left            =   705
         TabIndex        =   11
         Top             =   2745
         Width           =   2190
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Text"
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   315
         Width           =   720
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Tag             =   "False"
         Top             =   600
         Width           =   2790
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Font"
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   1575
         Width           =   810
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   705
         TabIndex        =   9
         Top             =   1935
         Width           =   2190
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   705
         TabIndex        =   10
         Top             =   2340
         Width           =   2190
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Default"
         Height          =   240
         Index           =   2
         Left            =   2070
         TabIndex        =   8
         Top             =   1605
         Width           =   840
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Family:"
         Height          =   285
         Index           =   2
         Left            =   135
         TabIndex        =   33
         Top             =   1980
         Width           =   540
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "(Can include some HTML tags)"
         Height          =   285
         Left            =   660
         TabIndex        =   26
         Top             =   930
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Size:"
         Height          =   270
         Index           =   0
         Left            =   285
         TabIndex        =   25
         Top             =   2790
         Width           =   390
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Color:"
         Height          =   270
         Index           =   1
         Left            =   210
         TabIndex        =   24
         Top             =   2385
         Width           =   450
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Anchor"
      Height          =   2955
      Left            =   3435
      TabIndex        =   22
      Top             =   3300
      Width           =   3285
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   2790
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   2370
         Width           =   2790
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   6
         Left            =   120
         TabIndex        =   17
         Top             =   1515
         Width           =   2790
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Default"
         Height          =   240
         Index           =   3
         Left            =   2070
         TabIndex        =   18
         Top             =   2100
         Width           =   825
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Link Target"
         Height          =   225
         Left            =   135
         TabIndex        =   36
         Top             =   2130
         Width           =   915
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Link URL"
         Height          =   225
         Left            =   135
         TabIndex        =   35
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Link ToolTip"
         Height          =   225
         Left            =   135
         TabIndex        =   34
         Top             =   1275
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   315
      Index           =   0
      Left            =   5790
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6360
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Index           =   1
      Left            =   4830
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6360
      Width           =   915
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   615
      TabIndex        =   0
      Text            =   "(None Specifyed)"
      Top             =   75
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label:"
      Height          =   225
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   450
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "(Not used in actual tree)"
      Height          =   270
      Left            =   2865
      TabIndex        =   31
      Top             =   105
      Width           =   1785
   End
End
Attribute VB_Name = "frmItem"
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
        Text3.Text = .Value("Label")
        Check1(0).Value = -CInt(.Value("UseCollapsed"))
        ctlMedia1.Value = .Value("Collapsed")
        Check1(1).Value = -CInt(.Value("UseExpanded"))
        ctlMedia2.Value = .Value("Expanded")
        Check1(2).Value = -CInt(.Value("UseMouseOut"))
        ctlMedia3.Value = .Value("MouseOut")
        Check1(3).Value = -CInt(.Value("UseMouseOver"))
        ctlMedia4.Value = .Value("MouseOver")
        Check2.Value = -CInt(.Value("UseText"))
        Text1.Text = .Value("Text")
        
        If (.Cast("UseFont") = "d") Then
            Check4(2).Value = 1
            Check3.Value = 0
        Else
            Check4(2).Value = 0
            Check3.Value = -CInt(.Value("UseFont"))
        End If
        
        Combo1(0).Text = .Value("FontFamily")
        Combo1(1).Text = .Value("FontColor")
        Combo1(2).Text = .Value("FontSize")
        
        Text2(4).Text = .Value("LinkURL")
        Text2(6).Text = .Value("LinkToolTip")
        
        If (.Cast("UseLinkTarget") = "d") Then
            Check4(3).Value = 1
        Else
            Check4(3).Value = 0
        End If
        Text2(5).Text = .Value("LinkTarget")
    
    End With

    EnableForm

    Text1.Tag = True

End Property
Public Property Get XMLText() As String
    
    With nItem
        .Value("Label") = Text3.Text
        .Value("UseCollapsed") = CBool(Check1(0).Value)
        .Value("Collapsed") = ctlMedia1.Value
        .Value("UseExpanded") = CBool(Check1(1).Value)
        .Value("Expanded") = ctlMedia2.Value
        .Value("UseMouseOut") = CBool(Check1(2).Value)
        .Value("MouseOut") = ctlMedia3.Value
        .Value("UseMouseOver") = CBool(Check1(3).Value)
        .Value("MouseOver") = ctlMedia4.Value
        .Value("UseText") = CBool(Check2.Value)
        .Value("Text") = Text1.Text
        
        If CBool(Check4(2).Value) Then
            .Value("UseFont") = "Default"
        Else
            .Value("UseFont") = CBool(Check3.Value)
        End If
        .Value("FontFamily") = Combo1(0).Text
        .Value("FontColor") = Combo1(1).Text
        .Value("FontSize") = Combo1(2).Text
        
        .Value("LinkURL") = Text2(4).Text
        .Value("LinkToolTip") = Text2(6).Text
        
        If CBool(Check4(3).Value) Then
            .Value("UseLinkTarget") = "Default"
        Else
            .Value("UseLinkTarget") = Not (Trim(Text2(5).Text) = vbNullString)
        End If
        .Value("LinkTarget") = Text2(5).Text
    
    End With
    
    XMLText = nItem.XMLText
    
End Property

Private Sub Check1_Click(Index As Integer)
    EnableForm
End Sub

Private Sub Check2_Click()
    EnableForm
End Sub

Private Sub Check3_Click()
    EnableForm
End Sub

Private Sub Check4_Click(Index As Integer)
    EnableForm
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 1
            If Trim(Text3.Text) = "" Then
                If Not Trim(Text1.Text) = "" Then
                    Text3.Text = Text1.Text
                End If
            End If
            
            If Trim(Text3.Text) = "" Then
                MsgBox "You must enter a label for this item." & vbCrLf & vbCrLf & _
                        "The label value is not visible in the final tree" & vbCrLf & _
                        "and would generally be the same as the text value." & vbCrLf & _
                        "However, since the text value can be disabled," & vbCrLf & _
                        "the label value is used to identify the item.", vbInformation
            Else
                IsOk = True
                Me.Hide
            End If
        Case 0
            IsOk = False
            Me.Hide
    End Select
End Sub

Private Sub Form_Load()
    InitFontFamily Combo1(0)
    InitFontColor Combo1(1)
    InitFontSize Combo1(2)
    
    Dim ItemHeight As String
    ItemHeight = Trim(frmMain.nBase.Value("Height"))
    
    
    ctlMedia1.Inititialize "Bullet Icons", ItemHeight & "x" & ItemHeight
    ctlMedia2.Inititialize "Bullet Icons", ItemHeight & "x" & ItemHeight
    
    ctlMedia3.Inititialize "Menu Pictures", ItemHeight & "x0"
    ctlMedia4.Inititialize "Menu Pictures", ItemHeight & "x0"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = True
        IsOk = False
        Me.Hide
    End If
End Sub

Public Function EnableForm()
    
    ctlMedia1.Enabled = CBool(Check1(0).Value)
    ctlMedia1.BackColor = IIf(ctlMedia1.Enabled, &HFFFFFF, &HE0E0E0)
    
    ctlMedia2.Enabled = CBool(Check1(1).Value)
    ctlMedia2.BackColor = IIf(ctlMedia2.Enabled, &HFFFFFF, &HE0E0E0)
    
    ctlMedia3.Enabled = CBool(Check1(2).Value)
    ctlMedia3.BackColor = IIf(ctlMedia3.Enabled, &HFFFFFF, &HE0E0E0)
    
    ctlMedia4.Enabled = CBool(Check1(3).Value)
    ctlMedia4.BackColor = IIf(ctlMedia4.Enabled, &HFFFFFF, &HE0E0E0)
    
    Text1.Enabled = CBool(Check2.Value)
    Text1.BackColor = IIf(Text1.Enabled, &HFFFFFF, &HE0E0E0)
    
    Check3.Enabled = Not CBool(Check4(2).Value)
    Combo1(0).Enabled = CBool(Check3.Value) And Check3.Enabled
    Combo1(0).BackColor = IIf(Combo1(0).Enabled, &HFFFFFF, &HE0E0E0)
    
    Combo1(1).Enabled = CBool(Check3.Value) And Check3.Enabled
    Combo1(1).BackColor = IIf(Combo1(1).Enabled, &HFFFFFF, &HE0E0E0)
    
    Combo1(2).Enabled = CBool(Check3.Value) And Check3.Enabled
    Combo1(2).BackColor = IIf(Combo1(2).Enabled, &HFFFFFF, &HE0E0E0)
    
    Text2(5).Enabled = Not CBool(Check4(3).Value)
    Text2(5).BackColor = IIf(Text2(5).Enabled, &HFFFFFF, &HE0E0E0)
End Function

Private Sub Text1_Change()
    If Text1.Tag = True Then Text3.Text = Text1.Text
End Sub
