VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl usrWizOperation 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   LockControls    =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "usrWizOperation.ctx":0000
   Begin VB.CommandButton Command1 
      Caption         =   "&Edit"
      Height          =   330
      Left            =   4035
      TabIndex        =   3
      Top             =   1335
      Width           =   630
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2220
      Left            =   120
      TabIndex        =   2
      Top             =   1275
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   3916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Param Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Param Value"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   915
      Width           =   3780
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   315
      Width           =   3780
   End
   Begin VB.Label Label1 
      Caption         =   "What should this operation do?"
      Height          =   210
      Left            =   180
      TabIndex        =   5
      Top             =   675
      Width           =   3825
   End
   Begin VB.Label Label5 
      Caption         =   "Enter a name or description for this operation."
      Height          =   225
      Left            =   165
      TabIndex        =   4
      Top             =   60
      Width           =   3285
   End
End
Attribute VB_Name = "usrWizOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public Property Get Parameters() As ListView
    Set Parameters = ListView1
End Property

Public Property Get Procedure() As String
    If Combo1.ListIndex > -1 Then
        Procedure = DisplayToScript(Combo1.List(Combo1.ListIndex))
    Else
        Procedure = ""
    End If
End Property
Public Property Let Procedure(ByVal newVal As String)
    Dim procIndex As Integer
    procIndex = IsOnList(Combo1, ScriptToDisplay(newVal))
    If procIndex > -1 Then
        Combo1.ListIndex = procIndex
    End If
End Property
Public Property Get Description() As String
    Description = Text1.Text
End Property
Public Property Let Description(ByVal newVal As String)
    Text1.Text = newVal
End Property

Private Sub Combo1_Click()
    
    InitializeParameters
    
End Sub

Public Sub InitializeParameters()
    ListView1.ListItems.Clear
    Dim proc
    
    Dim ProcName As String
    ProcName = DisplayToScript(Combo1.List(Combo1.ListIndex))
    
    For Each proc In Procedures(ProcName).Parameters
        ListView1.ListItems.Add , , Replace(proc.ParamName, "_", " ")
    Next

End Sub
Public Sub InitializeProcedures()
    If frmMain.ScriptControl1.Procedures.Count > 0 Then
        Dim proc
        For Each proc In Procedures
            Combo1.AddItem ScriptToDisplay(proc.ProcedureName)
        Next
    End If
End Sub
Private Sub InvokeEdit()
    Load frmEditParam
    frmEditParam.EditParameter DisplayToScript(Combo1.List(Combo1.ListIndex)), DisplayToScript(ListView1.SelectedItem.Text)
    frmEditParam.Show 1
End Sub
Private Sub Command1_Click()
    If Not ListView1.SelectedItem Is Nothing Then
        InvokeEdit
    Else
        MsgBox "Please select a parameter to edit.", vbInformation, "Edit Parameter"
    End If
End Sub

Private Sub ListView1_DblClick()
    If Not ListView1.SelectedItem Is Nothing Then
        InvokeEdit
    End If
End Sub

Private Sub UserControl_Resize()

    UserControl.Height = 3600
    UserControl.Width = 4800

End Sub
