VERSION 5.00
Object = "{6C527299-AA3C-4211-A6A6-494DE2E03DF5}#461.0#0"; "NTControls22.ocx"
Begin VB.Form frmEditParam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Parameter Value"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   Icon            =   "frmEditParam.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin NTControls22.BrowseButton BrowseButton1 
      Height          =   315
      Left            =   3975
      TabIndex        =   1
      Top             =   615
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      bBrowseTitle    =   "Browse for File"
      bBrowseAction   =   0
      bFileFilter     =   ""
      bFileFilterIndex=   0
      pButtonDimension=   315
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Index           =   1
      Left            =   2550
      TabIndex        =   3
      Top             =   1605
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   315
      Index           =   0
      Left            =   3510
      TabIndex        =   4
      Top             =   1605
      Width           =   870
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   285
      TabIndex        =   0
      Top             =   630
      Width           =   3630
   End
   Begin VB.ComboBox cmbParamValue 
      Height          =   315
      Left            =   300
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   615
      Width           =   3630
   End
   Begin VB.TextBox txtParamValue 
      Height          =   810
      Left            =   285
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   615
      Visible         =   0   'False
      Width           =   3630
   End
   Begin VB.Label lblParamName 
      Height          =   195
      Left            =   300
      TabIndex        =   5
      Top             =   330
      Width           =   3630
   End
End
Attribute VB_Name = "frmEditParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private ParamType As String
Private ValidateNum As Boolean

Public Function EditParameter(ByVal ProcName As String, ByVal ParamName As String)
    
    ValidateNum = False
    txtFileName.Visible = False
    txtParamValue.Visible = False
    cmbParamValue.Clear
    cmbParamValue.Visible = False
    BrowseButton1.Visible = False
    
    lblParamName.Caption = ScriptToDisplay(ParamName)
    
    ParamType = Procedures(ProcName).Parameters(ParamName).ParamType
    
    Select Case LCase(ParamType)
        Case "boolean"
            cmbParamValue.Visible = True
            cmbParamValue.AddItem "Yes (Enabled)"
            cmbParamValue.AddItem "No (Disabled)"
            
        Case "numeric", "integer", "long", "byte", "num"
            ValidateNum = True
            txtParamValue.Visible = True
        Case "string"
            txtParamValue.Visible = True
        Case "browse"
            txtFileName.Visible = True
            BrowseButton1.Visible = True
        Case Else
            If EnumeratorExists(ParamType) Then
                cmbParamValue.Visible = True
                Dim enumVal
                For Each enumVal In Enumerators(ParamType).EnumValues
                    cmbParamValue.AddItem ScriptToDisplay(enumVal.EnumName)
                Next
            
            Else
                MsgBox "Enumerator not found: " & ParamType, vbInformation, "Edit Parameter"
            End If
    End Select
    
    GetParameter
        
End Function

Private Sub GetParameter()
    Dim lstView
    Set lstView = frmWizard.usrWizOperation1.Parameters
    
    Select Case LCase(ParamType)
        Case "numeric", "integer", "long", "byte", "num", "string"
            txtParamValue.Text = lstView.SelectedItem.SubItems(1)
        Case "browse"
            txtFileName.Text = lstView.SelectedItem.SubItems(1)
        Case Else
            If IsOnList(cmbParamValue, lstView.SelectedItem.SubItems(1)) > -1 Then
                cmbParamValue.ListIndex = IsOnList(cmbParamValue, lstView.SelectedItem.SubItems(1))
            End If
    End Select
    
End Sub
Private Sub SetParameter()
    Dim lstView
    Set lstView = frmWizard.usrWizOperation1.Parameters
    
    Select Case LCase(ParamType)
        Case "numeric", "integer", "long", "byte", "num", "string"
            lstView.SelectedItem.SubItems(1) = txtParamValue.Text
        Case "browse"
            lstView.SelectedItem.SubItems(1) = txtFileName.Text
        Case "boolean"
            lstView.SelectedItem.SubItems(1) = cmbParamValue.List(cmbParamValue.ListIndex)
        Case Else
            lstView.SelectedItem.SubItems(1) = cmbParamValue.List(cmbParamValue.ListIndex)
    End Select
    
End Sub

Private Sub BrowseButton1_ButtonClick(ByVal BrowseReturn As String)
    If BrowseReturn <> "" Then
        txtFileName.Text = BrowseReturn
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 1
            If ValidateNum Then
                If Not IsNumeric(txtParamValue.Text) Then
                    MsgBox "You must enter a numeric value.", vbInformation, "Edit Parameter"
                Else
                    SetParameter
                    Unload Me
                End If
            Else
                SetParameter
                Unload Me
            End If
        Case 0
            Unload Me
    End Select
End Sub

