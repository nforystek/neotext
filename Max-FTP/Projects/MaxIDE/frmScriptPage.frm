VERSION 5.00
Object = "{7FF3E6C6-1DB9-4FC3-8810-D6E83C52513D}#1518.1#0"; "NTControls22.ocx"
Begin VB.Form frmScriptPage 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmScriptPage.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin NTControls22.CodeEdit CodeEdit1 
      Height          =   2445
      Left            =   405
      TabIndex        =   0
      Top             =   240
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   4313
      FontSize        =   9
      BackColor       =   16777215
      ColorText       =   0
      ColorDream1     =   8388736
      ColorDream2     =   8388608
      ColorDream3     =   8421376
      ColorDream4     =   32768
      ColorDream5     =   32896
      ColorDream6     =   16512
   End
End
Attribute VB_Name = "frmScriptPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private ReadOnlyCode As Boolean

Private cItem As clsItem

Public Property Get Locked() As Boolean
    Locked = CodeEdit1.Locked
End Property
Public Property Let Locked(ByVal newVal As Boolean)
    CodeEdit1.Locked = newVal

End Property

Public Property Get Language() As String
    Language = CodeEdit1.Language
End Property
Public Property Let Language(ByVal newVal As String)
    CodeEdit1.Language = newVal
End Property
Public Property Get Item() As clsItem
    Set Item = cItem
End Property
Public Property Set Item(ByVal newVal As Object)

    Set cItem = newVal

    Me.Caption = newVal.ItemName

    CodeEdit1.Text = cItem.ItemSource
        
    RefreshWindowMenu
    SelectWindowMenu Me.Caption
    
End Property

Private Sub CodeEdit1_Change()

'    If CodeEdit1.ErrorLine > 0 Then
'        CodeEdit1.colorerror = CodeEdit1.ConvertColor(GetCollectSkinValue("script_textcolor"))
'    End If
    
    cItem.ItemSource = CodeEdit1.Text

    frmMainIDE.ProjectSet True
    
End Sub

Private Sub CodeEdit1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu frmMainIDE.mnuEdit
    End If
End Sub

Private Sub Form_Activate()
    SelectWindowMenu Me.Caption
    
End Sub

Private Sub Form_Load()
    
    CodeEdit1.ColorComment = CodeEdit1.ConvertColor(GetCollectSkinValue("script_commentcolor"))
    CodeEdit1.colorerror = CodeEdit1.ConvertColor(GetCollectSkinValue("script_errorcolor"))
    CodeEdit1.ColorOperator = CodeEdit1.ConvertColor(GetCollectSkinValue("script_operatorcolor"))
    CodeEdit1.ColorStatement = CodeEdit1.ConvertColor(GetCollectSkinValue("script_statementcolor"))
    CodeEdit1.ColorText = CodeEdit1.ConvertColor(GetCollectSkinValue("script_textcolor"))
    CodeEdit1.ColorVariable = CodeEdit1.ConvertColor(GetCollectSkinValue("script_userdefinedcolor"))
    CodeEdit1.ColorValue = CodeEdit1.ConvertColor(GetCollectSkinValue("script_valuecolor"))

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    CodeEdit1.Top = 0
    CodeEdit1.Left = 0
    CodeEdit1.Width = Me.ScaleWidth
    CodeEdit1.Height = Me.ScaleHeight
    If Err Then Err.Clear
    On Error GoTo 0
End Sub

Private Sub Form_Terminate()
    RefreshWindowMenu
End Sub

Attribute 