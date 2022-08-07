VERSION 5.00
Begin VB.UserControl URLBox 
   ClientHeight    =   1530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4395
   ScaleHeight     =   1530
   ScaleWidth      =   4395
   ToolboxBitmap   =   "URLBox.ctx":0000
   Begin NTControls22.BrowseButton BrowseButton1 
      Height          =   315
      Left            =   3375
      TabIndex        =   1
      Top             =   45
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
   End
   Begin NTControls22.AutoType Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   556
      ReadOnly        =   0   'False
   End
End
Attribute VB_Name = "URLBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary


Public Event Change()
Public Event KeyDown(ByVal KeyCode As Integer)

Public Property Get AutoTypeCombo() As AutoType
    Set AutoTypeCombo = Combo1
End Property

Public Property Get BrowseButton() As BrowseButton
    Set BrowseButton = BrowseButton1
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.Parent.hwnd
End Property
Public Property Get Text() As String
    Text = Combo1.Text
End Property
Public Property Let Text(ByVal newText As String)
    Combo1.Text = newText
    RaiseEvent Change
End Property
Private Sub BrowseButton1_ButtonClick(ByVal BrowseReturn As String)
    If BrowseReturn <> "" Then
        Combo1.Text = BrowseReturn
        RaiseEvent Change
    End If
End Sub

Private Sub Combo1_Change()
    RaiseEvent Change
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyDown(KeyCode)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Combo1.Top = 0
    Combo1.Left = 0
    BrowseButton1.Top = 0
    BrowseButton1.Left = UserControl.ScaleWidth - BrowseButton1.Width
    Combo1.Width = BrowseButton1.Left
    UserControl.Height = Combo1.Height
    If Err Then Err.Clear
    On Error GoTo 0
End Sub



