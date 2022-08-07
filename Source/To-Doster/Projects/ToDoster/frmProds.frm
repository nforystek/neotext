
VERSION 5.00
Begin VB.Form frmProds 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Edit"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   Icon            =   "frmProds.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Delete"
      Height          =   315
      Index           =   2
      Left            =   4530
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   825
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      Height          =   315
      Index           =   0
      Left            =   225
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
      Height          =   315
      Index           =   1
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1380
      MaxLength       =   65534
      TabIndex        =   1
      Top             =   1590
      Width           =   2280
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   1380
      TabIndex        =   0
      Top             =   90
      Width           =   3990
   End
   Begin VB.Image Image1 
      Height          =   1305
      Left            =   0
      Picture         =   "frmProds.frx":2CFA
      Top             =   0
      Width           =   1275
   End
End
Attribute VB_Name = "frmProds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN
Option Compare Text

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
        
            If Trim(Text1.Text) = "" Then
                MsgBox "Please enter a product name in the text box below the list.", vbOKOnly + vbInformation, App.Title
            Else
                If List1.ListIndex = 0 Then
                    List1.AddItem Text1.Text
                    colProducts.Add Text1.Text
                    List1.ListIndex = (List1.ListCount - 1)
                Else
                    List1.List(List1.ListIndex) = Text1.Text
                    Do Until colProducts.Count = 0
                        colProducts.Remove 1
                    Loop
                    If List1.ListCount > 1 Then
                        Dim cnt As Long
                        For cnt = 1 To List1.ListCount
                            colProducts.Add List1.List(cnt)
                        Next
                    End If
                End If
            End If
        Case 2
        
            If List1.ListIndex > 0 Then
                If MsgBox("Really delete this product - [" & List1.List(List1.ListIndex) & "]?", vbYesNo + vbQuestion, App.Title) = vbYes Then
                    colProducts.Remove List1.ListIndex
                    List1.RemoveItem List1.ListIndex
                    Text1.Text = ""
                    List1.ListIndex = 0
                End If
            Else
                MsgBox "This selection can not be deleted.", vbInformation, App.Title
            End If
    End Select
End Sub

Private Sub Form_Load()
    LoadSettings
    
    List1.AddItem "(Add New Product)"
    
    Dim nProduct
    For Each nProduct In colProducts
        List1.AddItem nProduct
    Next

    List1.ListIndex = 0
    
    SetForm
    
End Sub
Private Sub SetForm()
    
    If List1.ListIndex = 0 Then
        Command1(2).Enabled = False
        Text1.Text = ""
    Else
    
        Text1.Text = List1.List(List1.ListIndex)
        Command1(2).Enabled = True
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSettings
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub List1_Click()
    SetForm
End Sub
