VERSION 5.00
Begin VB.Form frmWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operation"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   Icon            =   "frmWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   195
      Picture         =   "frmWizard.frx":0442
      ScaleHeight     =   3555
      ScaleWidth      =   1275
      TabIndex        =   4
      Top             =   165
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   350
      Index           =   2
      Left            =   210
      TabIndex        =   1
      Top             =   4110
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "< &Back"
      Enabled         =   0   'False
      Height          =   350
      Index           =   1
      Left            =   4245
      TabIndex        =   2
      Top             =   4110
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Next >"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   5550
      TabIndex        =   3
      Top             =   4110
      Width           =   1200
   End
   Begin RemindMe.usrWizOperation usrWizOperation1 
      Height          =   3600
      Left            =   1845
      TabIndex        =   5
      Top             =   135
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   6350
   End
   Begin RemindMe.usrWizTimer usrWizTimer1 
      Height          =   3600
      Left            =   1860
      TabIndex        =   0
      Top             =   135
      Visible         =   0   'False
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   6350
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      X1              =   1680
      X2              =   1680
      Y1              =   90
      Y2              =   3795
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   1695
      X2              =   1695
      Y1              =   90
      Y2              =   3795
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   195
      X2              =   6750
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   195
      X2              =   6735
      Y1              =   3945
      Y2              =   3945
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public EditID As Long

Private Sub MakePanelVisible(ByVal panelName As String)
    usrWizOperation1.Visible = (panelName = "usrwizoperation")
    usrWizTimer1.Visible = (panelName = "usrwiztimer")
End Sub
Private Function GetVisiblePanel() As String
    If usrWizOperation1.Visible Then
        GetVisiblePanel = LCase(Trim(TypeName(usrWizOperation1)))
    End If
    If usrWizTimer1.Visible Then
        GetVisiblePanel = LCase(Trim(TypeName(usrWizTimer1)))
    End If
End Function

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            If Command1(0).Caption = "&Finish" Then
                If usrWizOperation1.Description = "" Then
                    MsgBox "You must enter a name for this operation.", vbInformation, "Operation"
                Else
                    If usrWizOperation1.Procedure = "" Then
                        MsgBox "You must specify what this operation should do.", vbInformation, "Operation"
                    Else
                        SaveOperation
                    End If
                End If
            Else
                Select Case GetVisiblePanel
                    Case "usrwizoperation"
                        MakePanelVisible "usrwiztimer"
                        Command1(0).Enabled = True
                        Command1(1).Enabled = True
                        Command1(0).Caption = "&Finish"
                End Select
            End If
        Case 1
            Select Case GetVisiblePanel
                Case "usrwiztimer"
                    MakePanelVisible "usrwizoperation"
                    Command1(0).Enabled = True
                    Command1(1).Enabled = False
                    Command1(0).Caption = "&Next >"
            End Select
        Case 2
            If MsgBox("Are you sure you want to cancel?", vbQuestion + vbYesNo, "Cancel") = vbYes Then
                Unload Me
            End If
    End Select
End Sub

Public Function SaveOperation()
    Dim rs As New ADODB.Recordset
    Dim updateID As String
    If EditID = 0 Then

        Dim Seed As String
        Seed = Replace(modGUID.GUID, "-", "")
        DBConn.rsQuery rs, "INSERT INTO Operations (Name) VALUES ('" & Seed & "');"
        DBConn.rsQuery rs, "SELECT * FROM Operations WHERE Name='" & Seed & "';"
        updateID = rs("ID")
    Else

        updateID = frmMain.lstObjects.SelectedItem.Tag
    End If
    
    DBConn.rsQuery rs, "UPDATE Operations SET Name='" & Replace(usrWizOperation1.Description, "'", "''") & "', " & _
                        "Enabled=" & CBool(usrWizTimer1.Enabled) & ", " & _
                        "ProcName='" & Replace(usrWizOperation1.Procedure, "'", "''") & "', " & _
                        "ScheduleType=" & usrWizTimer1.ScheduleType & ", " & _
                        "ExecuteDate='" & Replace(usrWizTimer1.ExecuteDate, "'", "''") & "', " & _
                        "ExecuteTime='" & Replace(usrWizTimer1.ExecuteTime, "'", "''") & "', " & _
                        "IncrementType=" & usrWizTimer1.IncrementType & ", " & _
                        "IncrementInterval=" & usrWizTimer1.IncrementInterval & " " & _
                        "WHERE ID=" & updateID & ";"
    
    DBConn.rsQuery rs, "DELETE * FROM OperationParams WHERE ParentID=" & updateID & ";"
    
    Dim lstItem
    Dim ParamType As String
    Dim ParamValue As String
    Dim Enumerator As String
    
    For Each lstItem In usrWizOperation1.Parameters.ListItems
        Select Case Trim(LCase(Procedures(usrWizOperation1.Procedure).Parameters(DisplayToScript(lstItem.Text)).ParamType))
            Case "boolean"
                ParamType = 1
                If LCase(Left(lstItem.SubItems(1), 3)) = "yes" Then
                    ParamValue = "True"
                Else
                    ParamValue = "False"
                End If
            Case "numeric", "integer", "long", "byte", "num"
                ParamType = 2
                ParamValue = lstItem.SubItems(1)
                If Not IsNumeric(lstItem.SubItems(1)) Then
                    ParamValue = 0
                End If
            Case "string", "browse"
                ParamType = 3
                ParamValue = lstItem.SubItems(1)
            Case Else
                ParamType = 4
                Enumerator = Procedures(usrWizOperation1.Procedure).Parameters(DisplayToScript(lstItem.Text)).ParamType
                If Not lstItem.SubItems(1) = "" Then
                    ParamValue = Enumerators(Enumerator).EnumValues(DisplayToScript(lstItem.SubItems(1))).EnumValue
                Else
                    ParamValue = Enumerators(Enumerator).EnumValues(1).EnumValue
                End If
        End Select
        
        DBConn.rsQuery rs, "INSERT INTO OperationParams (ParentID, ParamNum, ParamType, ParamValue) VALUES (" & updateID & "," & lstItem.Index & "," & ParamType & ",'" & Replace(ParamValue, "'", "''") & "')"
    
    Next

    If Not rs.State = 0 Then rs.Close
    Set rs = Nothing
    
    frmMain.SendMessage "updateoperation:" & Trim(updateID)
    
    Unload Me
End Function

Private Sub Form_Load()
    usrWizOperation1.InitializeProcedures
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = MsgBox("Are you sure you want to cancel?", vbQuestion + vbYesNo, "Cancel") = vbNo
    End If
End Sub
