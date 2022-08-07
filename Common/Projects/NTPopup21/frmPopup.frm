VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPopup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   Icon            =   "frmPopup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   145
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   389
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtMessage 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   1380
      Left            =   990
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   315
      Width           =   4770
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":000C
            Key             =   "k64"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":045E
            Key             =   "k32"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":08B0
            Key             =   "k48"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":0D02
            Key             =   "k16"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   315
      Index           =   1
      Left            =   4875
      TabIndex        =   0
      Top             =   1800
      Width           =   900
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   990
      TabIndex        =   1
      Top             =   30
      Width           =   4740
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Top             =   600
      Width           =   480
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private pMyParent As Window

Friend Property Get MyParent() As Window
    Set MyParent = pMyParent
End Property

Friend Property Set MyParent(ByRef newVal As Window)
    Set pMyParent = newVal
End Property

Private Sub Command1_Click(Index As Integer)
    Me.Visible = False
    If Not (pMyParent Is Nothing) Then
        pMyParent.Visible = False
        Unhook pMyParent
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = True
    End If
    If Not (pMyParent Is Nothing) Then
        pMyParent.Visible = False
        Unhook pMyParent
    End If
    Me.Visible = False
End Sub

Private Sub Label2_Click()
    If Not (CStr(Label2.Tag) = "") Then RunFile CStr(Label2.Tag)
End Sub
