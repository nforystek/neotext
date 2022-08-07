VERSION 5.00
Begin VB.UserControl NameInformation 
   ClientHeight    =   1344
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3072
   LockControls    =   -1  'True
   ScaleHeight     =   1344
   ScaleWidth      =   3072
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   288
      Left            =   0
      TabIndex        =   0
      Tag             =   "(Company or Artist Name)"
      Top             =   0
      Width           =   3060
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   0
      TabIndex        =   1
      Tag             =   "(Line 1 of Address)"
      Top             =   348
      Width           =   3060
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   0
      TabIndex        =   2
      Tag             =   "(Line 2 of Address)"
      Top             =   696
      Width           =   3060
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   0
      TabIndex        =   3
      Tag             =   "(City or Township)"
      Top             =   1044
      Width           =   1692
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1740
      TabIndex        =   4
      Tag             =   "(State)"
      Top             =   1044
      Width           =   528
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   2340
      TabIndex        =   5
      Tag             =   "(Zip-Code)"
      Top             =   1044
      Width           =   720
   End
End
Attribute VB_Name = "NameInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetFocus Lib "user32" () As Long

Public Property Get Indetification() As String
    If Not Text1.Tag = Text1.Text Then
        Indetification = Text1.Text
    End If
End Property
Public Property Let Indetification(ByVal RHS As String)
    Text1.Text = RHS
    TextChange Text1
End Property

Public Property Get Address1() As String
    If Not Text2.Tag = Text2.Text Then
        Address1 = Text2.Text
    End If
End Property
Public Property Let Address1(ByVal RHS As String)
    Text2.Text = RHS
    TextChange Text2
End Property

Public Property Get Address2() As String
    If Not Text3.Tag = Text3.Text Then
        Address2 = Text3.Text
    End If
End Property
Public Property Let Address2(ByVal RHS As String)
    Text3.Text = RHS
    TextChange Text3
End Property

Public Property Get City() As String
    If Not Text4.Tag = Text4.Text Then
        City = Text4.Text
    End If
End Property
Public Property Let City(ByVal RHS As String)
    Text4.Text = RHS
    TextChange Text4
End Property

Public Property Get State() As String
    If Not Text5.Tag = Text5.Text Then
        State = Text5.Text
    End If
End Property
Public Property Let State(ByVal RHS As String)
    Text5.Text = RHS
    TextChange Text5
End Property

Public Property Get Zip() As String
    If Not Text6.Tag = Text6.Text Then
        Zip = Text6.Text
    End If
End Property
Public Property Let Zip(ByVal RHS As String)
    Text6.Text = RHS
    TextChange Text6
End Property

Private Sub TextChange(ByRef TxtBox As TextBox)
    If GetFocus = TxtBox.hwnd Then
        If TxtBox.Text = TxtBox.Tag Then
            TxtBox.Text = ""
            TxtBox.ForeColor = &H80000008
        End If
    Else
        If TxtBox.Text = "" Or TxtBox.Text = TxtBox.Tag Then
            TxtBox.ForeColor = &H8000000B
            TxtBox.Text = TxtBox.Tag
        Else
            TxtBox.ForeColor = &H80000008
        End If
    End If
End Sub

Private Sub Text1_Change()
    TextChange Text1
End Sub

Private Sub Text1_Click()
    TextChange Text1
End Sub

Private Sub Text1_GotFocus()
    TextChange Text1
End Sub

Private Sub Text1_LostFocus()
    TextChange Text1
End Sub

Private Sub Text2_Change()
    TextChange Text2
End Sub

Private Sub Text2_Click()
    TextChange Text2
End Sub

Private Sub Text2_GotFocus()
    TextChange Text2
End Sub

Private Sub Text2_LostFocus()
    TextChange Text2
End Sub

Private Sub Text3_Change()
    TextChange Text3
End Sub

Private Sub Text3_Click()
    TextChange Text3
End Sub

Private Sub Text3_GotFocus()
    TextChange Text3
End Sub

Private Sub Text3_LostFocus()
    TextChange Text3
End Sub

Private Sub Text4_Change()
    TextChange Text4
End Sub

Private Sub Text4_Click()
    TextChange Text4
End Sub

Private Sub Text4_GotFocus()
    TextChange Text4
End Sub

Private Sub Text4_LostFocus()
    TextChange Text4
End Sub

Private Sub Text5_Change()
    TextChange Text5
End Sub

Private Sub Text5_Click()
    TextChange Text5
End Sub

Private Sub Text5_GotFocus()
    TextChange Text5
End Sub

Private Sub Text5_LostFocus()
    TextChange Text5
End Sub

Private Sub Text6_Change()
    TextChange Text6
End Sub

Private Sub Text6_Click()
    TextChange Text6
End Sub

Private Sub Text6_GotFocus()
    TextChange Text6
End Sub

Private Sub Text6_LostFocus()
    TextChange Text6
End Sub

Private Sub UserControl_Initialize()
    Text1.Text = Text1.Tag
    Text2.Text = Text2.Tag
    Text3.Text = Text3.Tag
    Text4.Text = Text4.Tag
    Text5.Text = Text5.Tag
    Text6.Text = Text6.Tag
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 3072
    UserControl.Height = 1344
End Sub
