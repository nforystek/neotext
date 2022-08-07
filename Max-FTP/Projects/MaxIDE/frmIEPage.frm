VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmIEPage 
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   Icon            =   "frmIEPage.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4020
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1260
      Left            =   930
      TabIndex        =   0
      Top             =   540
      Width           =   3090
      ExtentX         =   5450
      ExtentY         =   2222
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
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
End
Attribute VB_Name = "frmIEPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public Sub ShowForm(ByVal Title As String, ByVal URL As String)
    Me.Caption = Title
   
    RefreshWindowMenu
    SelectWindowMenu Title
    
    If frmMainIDE.Visible Then
        Me.ZOrder 0
    End If
    
    WebBrowser1.Navigate URL
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
        WebBrowser1.Top = 0
        WebBrowser1.Left = 0
        WebBrowser1.Width = Me.ScaleWidth
        WebBrowser1.Height = Me.ScaleHeight
    Err.Clear
End Sub

Private Sub Form_Terminate()
    RefreshWindowMenu
End Sub

Private Sub Form_Activate()
    SelectWindowMenu Me.Caption
End Sub

Attribute 