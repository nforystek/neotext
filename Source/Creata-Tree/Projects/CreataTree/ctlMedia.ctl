VERSION 5.00
Begin VB.UserControl cltMedia 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4095
   LockControls    =   -1  'True
   ScaleHeight     =   1200
   ScaleWidth      =   4095
   ToolboxBitmap   =   "ctlMedia.ctx":0000
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   405
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   2145
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2850
      Picture         =   "ctlMedia.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   300
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2550
      Picture         =   "ctlMedia.ctx":068B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   300
   End
End
Attribute VB_Name = "cltMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private nLibrary As String
Private nForced As String
Private nInfo As ImageInfoType
Public Property Get Enabled() As Boolean
    Enabled = Text1.Enabled
End Property
Public Property Let Enabled(ByVal NewValue As Boolean)
    Text1.Enabled = NewValue
End Property
Public Sub Inititialize(ByVal Library As String, Optional ByVal ForcedDimensions As String = "0x0")
    nLibrary = Library
    nForced = ForcedDimensions
End Sub

Public Property Get Value() As String
    If Text1.Text = BlankImageText Then
        Value = vbNullString
    Else
        Value = nLibrary & "\" & nInfo.name
    End If
End Property
Public Property Let Value(ByVal NewValue As String)
    If NewValue = vbNullString Then
        Text1.Text = BlankImageText
    Else
        With nInfo
            nInfo = GetImageInfo(MediaFolder & NewValue)
            Text1.Text = .Title & " " & .Desc
        End With
    End If
End Property
Public Property Let BackColor(ByVal NewValue As Long)
    UserControl.BackColor = NewValue
    Text1.BackColor = NewValue
End Property

Private Sub Image1_Click()
    If Enabled Then
        With frmMedia
            .Library = nLibrary
            .ForcedDimensions = nForced
            .Selected = Value
            .Show 1, Me
            If .IsOk Then
                Value = .Value
            End If
        End With
        Unload frmMedia
    End If
End Sub

Private Sub Image2_Click()
    If Enabled Then
        If MsgBox("Are you sure you want to clear this image field?", vbYesNo + vbQuestion) = vbYes Then
            Text1.Text = BlankImageText
        End If
    End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Enabled Then
    
        If KeyCode = 46 Then
            Image2_Click
        End If
        
        If KeyCode = 32 Then
            Image1_Click
        End If
    
    End If
    
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Text1.Left = 0
    Text1.Top = 0
    UserControl.Height = Text1.Height
    Image1.Top = 0
    Image2.Top = 0
    Image2.Left = UserControl.ScaleWidth - Image1.Width
    Image1.Left = Image2.Left - Image1.Width
    Text1.Width = Image1.Left
    If Err Then Err.Clear
    On Error GoTo 0
End Sub
