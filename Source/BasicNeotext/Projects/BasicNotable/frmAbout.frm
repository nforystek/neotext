VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Notable"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3180
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   3180
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "&Close"
      DownPicture     =   "frmAbout.frx":000C
      Height          =   345
      Left            =   150
      MaskColor       =   &H00000000&
      Picture         =   "frmAbout.frx":0153
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   135
      Width           =   840
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1755
      Left            =   15
      ScaleHeight     =   1695
      ScaleWidth      =   3090
      TabIndex        =   0
      Top             =   0
      Width           =   3150
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private Sub Command1_Click()
    Unload Me
End Sub
