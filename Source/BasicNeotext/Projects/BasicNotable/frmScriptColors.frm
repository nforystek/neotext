VERSION 5.00
Object = "{C98B112F-745F-4542-B5B3-DDFADF1F6E2F}#1177.0#0"; "NTControls22.ocx"
Begin VB.Form frmScriptColors 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Color Set"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin NTControls22.CodeEdit CodeEdit1 
      Height          =   3720
      Left            =   75
      TabIndex        =   5
      Top             =   30
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   6562
      FontSize        =   9
      Locked          =   -1  'True
      ColorDream1     =   8388736
      ColorDream2     =   8388608
      ColorDream3     =   8421376
      ColorDream4     =   32768
      ColorDream5     =   32896
      ColorDream6     =   16512
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   300
      Index           =   2
      Left            =   6030
      TabIndex        =   4
      Top             =   3840
      Width           =   990
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Apply"
      Height          =   300
      Index           =   1
      Left            =   4950
      TabIndex        =   3
      Top             =   3840
      Width           =   990
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Reset"
      Height          =   300
      Index           =   0
      Left            =   3870
      TabIndex        =   2
      Top             =   3840
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Pick Color"
      Height          =   300
      Left            =   2595
      TabIndex        =   1
      Top             =   3840
      Width           =   1185
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3840
      Width           =   2460
   End
End
Attribute VB_Name = "frmScriptColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private pLanguage As String
Private pColors(1 To 30) As OLE_COLOR

Public Property Get ColorProperty(ByVal Index As ColorProperties) As OLE_COLOR
    ColorProperty = pColors(Index)
End Property
Public Property Let ColorProperty(ByVal Index As ColorProperties, ByVal newVal As OLE_COLOR)
    pColors(Index) = newVal
    Select Case Index
        Case ColorProperties.BatchInkComment
            CodeEdit1.ColorDream1 = newVal
            CodeEdit1.LineDefines(1) = Dream1Index
        Case ColorProperties.BatchInkCommands
            CodeEdit1.ColorDream2 = newVal
            CodeEdit1.LineDefines(3) = Dream2Index
        Case ColorProperties.BatchInkFinished
            CodeEdit1.ColorDream3 = newVal
            CodeEdit1.LineDefines(5) = Dream3Index
        Case ColorProperties.BatchInkCurrently
            CodeEdit1.ColorDream4 = newVal
            CodeEdit1.LineDefines(7) = Dream4Index
        Case ColorProperties.BatchInkIncomming
            CodeEdit1.ColorDream5 = newVal
            CodeEdit1.LineDefines(9) = Dream5Index

        Case ColorProperties.JScriptComment
            CodeEdit1.ColorComment = newVal
        Case ColorProperties.JScriptStatements
            CodeEdit1.ColorStatement = newVal
        Case ColorProperties.JScriptOperators
            CodeEdit1.ColorOperator = newVal
        Case ColorProperties.JScriptVariables
            CodeEdit1.ColorVariable = newVal
        Case ColorProperties.JScriptValues
            CodeEdit1.ColorValue = newVal
        Case ColorProperties.JScriptError
            CodeEdit1.ColorError = newVal
            
        Case ColorProperties.VBScriptComment
            CodeEdit1.ColorComment = newVal
        Case ColorProperties.VBScriptStatements
            CodeEdit1.ColorStatement = newVal
        Case ColorProperties.VBScriptOperators
            CodeEdit1.ColorOperator = newVal
        Case ColorProperties.VBScriptVariables
            CodeEdit1.ColorVariable = newVal
        Case ColorProperties.VBScriptValues
            CodeEdit1.ColorValue = newVal
        Case ColorProperties.VBScriptError
            CodeEdit1.ColorError = newVal

        Case ColorProperties.NSISScriptComment
            CodeEdit1.ColorDream1 = newVal
            CodeEdit1.LineDefines(1) = Dream1Index
        Case ColorProperties.NSISScriptCommands
            CodeEdit1.ColorDream2 = newVal
            CodeEdit1.LineDefines(3) = Dream2Index
        Case ColorProperties.NSISScriptEqualJump
            CodeEdit1.ColorDream3 = newVal
            CodeEdit1.LineDefines(4) = Dream3Index
        Case ColorProperties.NSISScriptElseJump
            CodeEdit1.ColorDream4 = newVal
            CodeEdit1.LineDefines(5) = Dream4Index
        Case ColorProperties.NSISScriptAboveJump
            CodeEdit1.ColorDream5 = newVal
            CodeEdit1.LineDefines(6) = Dream5Index


        Case ColorProperties.AssemblyComment
            CodeEdit1.ColorDream1 = newVal
            CodeEdit1.LineDefines(1) = Dream1Index
        Case ColorProperties.AssemblyCommand
            CodeEdit1.ColorDream2 = newVal
            CodeEdit1.LineDefines(2) = Dream2Index
        Case ColorProperties.AssemblyNotation
            CodeEdit1.ColorDream3 = newVal
            CodeEdit1.LineDefines(3) = Dream3Index
        Case ColorProperties.AssemblyRegister
            CodeEdit1.ColorDream4 = newVal
            CodeEdit1.LineDefines(4) = Dream4Index
        Case ColorProperties.AssemblyParameter
            CodeEdit1.ColorDream5 = newVal
            CodeEdit1.LineDefines(5) = Dream5Index
        Case ColorProperties.AssemblyError
            CodeEdit1.ColorDream6 = newVal
            CodeEdit1.LineDefines(6) = Dream6Index
    End Select
    
End Property

Public Sub InitializeToScript(ByVal Language As String)
    pLanguage = Language
    Command2_Click 0
    
    Select Case Language
        Case "VBScript"
            Me.Caption = pLanguage & " " & Me.Caption
            CodeEdit1.Language = Language
            
            CodeEdit1.Text = "'this line is commented text" & vbCrLf & _
            "Function UserDefined(Arg1, Arg2)" & vbCrLf & _
            "" & vbCrLf & _
            "   Dim variable" & vbCrLf & _
            "   variable = ""some value""" & vbCrLf & _
            "   If (9 + 10) = 9 Mod Arg2 Then" & vbCrLf & _
            "       UserDefined = Empty" & vbCrLf & _
            "   End if" & vbCrLf & _
            "" & vbCrLf & _
            "   ....this line is an error!!!" & vbCrLf & _
            "End Function" & vbCrLf
            
            CodeEdit1.ColorComment = fMain.ColorProperty(VBScriptComment)
            CodeEdit1.ColorStatement = fMain.ColorProperty(VBScriptStatements)
            CodeEdit1.ColorOperator = fMain.ColorProperty(VBScriptOperators)
            CodeEdit1.ColorVariable = fMain.ColorProperty(VBScriptVariables)
            CodeEdit1.ColorValue = fMain.ColorProperty(VBScriptValues)
            CodeEdit1.ColorError = fMain.ColorProperty(VBScriptError)
            
            CodeEdit1.ErrorLine = 10
            
            Combo1.AddItem "Comment"
            Combo1.AddItem "Statements"
            Combo1.AddItem "Operators"
            Combo1.AddItem "Variables"
            Combo1.AddItem "Values"
            Combo1.AddItem "Error"
            Combo1.ListIndex = 0
            
        Case "JScript"
            Me.Caption = pLanguage & " " & Me.Caption
            CodeEdit1.Language = Language
            
            CodeEdit1.Text = "//this line is commented text" & vbCrLf & _
            "function userDefined(Arg1, Arg2) {" & vbCrLf & _
            "" & vbCrLf & _
            "   var variable = new String();" & vbCrLf & _
            "   variable = 'some value';" & vbCrLf & _
            "   if ((9 + 10) = 9 * Arg2) {" & vbCrLf & _
            "       return null;" & vbCrLf & _
            "   }" & vbCrLf & _
            "" & vbCrLf & _
            "   ....this line is an error!!!" & vbCrLf & _
            "}" & vbCrLf

            CodeEdit1.ErrorLine = 10
            
            CodeEdit1.ColorComment = fMain.ColorProperty(JScriptComment)
            CodeEdit1.ColorStatement = fMain.ColorProperty(JScriptStatements)
            CodeEdit1.ColorOperator = fMain.ColorProperty(JScriptOperators)
            CodeEdit1.ColorVariable = fMain.ColorProperty(JScriptVariables)
            CodeEdit1.ColorValue = fMain.ColorProperty(JScriptValues)
            CodeEdit1.ColorError = fMain.ColorProperty(JScriptError)

            Combo1.AddItem "Comment"
            Combo1.AddItem "Statements"
            Combo1.AddItem "Operators"
            Combo1.AddItem "Variables"
            Combo1.AddItem "Values"
            Combo1.AddItem "Error"
            Combo1.ListIndex = 0
        Case "BatchInk"
            Me.Caption = pLanguage & " " & Me.Caption
            
            CodeEdit1.Language = "Defined"
            
            CodeEdit1.Text = "rem this is a commented line" & vbCrLf & _
            "" & vbCrLf & _
            "copy autoexec.bat autoexec.bak" & vbCrLf & _
            "" & vbCrLf & _
            "echo color of executed output" & vbCrLf & _
            "" & vbCrLf & _
            "echo color of active output" & vbCrLf & _
            "" & vbCrLf & _
            "echo color of queued output" & vbCrLf
            
            CodeEdit1.ColorDream1 = fMain.ColorProperty(BatchInkComment)
            CodeEdit1.ColorDream2 = fMain.ColorProperty(BatchInkCommands)
            CodeEdit1.ColorDream3 = fMain.ColorProperty(BatchInkFinished)
            CodeEdit1.ColorDream4 = fMain.ColorProperty(BatchInkCurrently)
            CodeEdit1.ColorDream5 = fMain.ColorProperty(BatchInkIncomming)
            
            CodeEdit1.BackColor = fMain.BackColor
                                                 
            CodeEdit1.LineDefines(1) = CodeEdit1.ColorDream1
            CodeEdit1.LineDefines(3) = CodeEdit1.ColorDream2
            CodeEdit1.LineDefines(5) = CodeEdit1.ColorDream3
            CodeEdit1.LineDefines(7) = CodeEdit1.ColorDream4
            CodeEdit1.LineDefines(9) = CodeEdit1.ColorDream5
           
            Combo1.AddItem "Comment"
            Combo1.AddItem "Commands"
            Combo1.AddItem "Finished"
            Combo1.AddItem "Currently"
            Combo1.AddItem "Incomming"
            Combo1.ListIndex = 0
        Case "NSISScript"
            Me.Caption = pLanguage & " " & Me.Caption
            
            CodeEdit1.Language = "Defined"
            
            CodeEdit1.Text = ";this is a commented line" & vbCrLf & _
            ";at cursor, numeric jumps highlight:" & vbCrLf & _
            "normal text or command statements" & vbCrLf & _
            "jump to true, i.e. equals in IntCmp" & vbCrLf & _
            ";false jump and less then, i.e. IntCmp" & vbCrLf & _
            ";greater then jump line, i.e. IntCmp" & vbCrLf
 
            CodeEdit1.ColorDream1 = fMain.ColorProperty(NSISScriptComment)
            CodeEdit1.ColorDream2 = fMain.ColorProperty(NSISScriptCommands)
            CodeEdit1.ColorDream3 = fMain.ColorProperty(NSISScriptEqualJump)
            CodeEdit1.ColorDream4 = fMain.ColorProperty(NSISScriptElseJump)
            CodeEdit1.ColorDream5 = fMain.ColorProperty(NSISScriptAboveJump)
            
            'CodeEdit1.BackColor = fMain.BackColor
                        
            CodeEdit1.LineDefines(1) = Dream1Index
            CodeEdit1.LineDefines(2) = Dream1Index
            CodeEdit1.LineDefines(3) = Dream2Index
            CodeEdit1.LineDefines(4) = Dream3Index
            CodeEdit1.LineDefines(5) = Dream4Index
            CodeEdit1.LineDefines(6) = Dream5Index
 
            Combo1.AddItem "Comment"
            Combo1.AddItem "Commands"
            Combo1.AddItem "EqualJump"
            Combo1.AddItem "ElseJump"
            Combo1.AddItem "AboveJump"
            Combo1.ListIndex = 0
            
        Case "Assembly"
            Me.Caption = pLanguage & " " & Me.Caption
            
            CodeEdit1.Language = "Defined"
            
            
            CodeEdit1.Text = ";this is a comment..." & vbCrLf & _
            "MOV JMP POP etc.." & vbCrLf & _
            " , &H , &H , &H , " & vbCrLf & _
            "AX AH BX BH etc.." & vbCrLf & _
            " ""data or numbers for example"" " & vbCrLf & _
            "syntax error statement color" & vbCrLf
 
            CodeEdit1.ColorDream1 = fMain.ColorProperty(AssemblyComment)
            CodeEdit1.ColorDream2 = fMain.ColorProperty(AssemblyCommand)
            CodeEdit1.ColorDream3 = fMain.ColorProperty(AssemblyNotation)
            CodeEdit1.ColorDream4 = fMain.ColorProperty(AssemblyRegister)
            CodeEdit1.ColorDream5 = fMain.ColorProperty(AssemblyParameter)
            CodeEdit1.ColorDream6 = fMain.ColorProperty(AssemblyError)
            
            'CodeEdit1.BackColor = fMain.BackColor
                        
            CodeEdit1.Redraw
            
            CodeEdit1.LineDefines(1) = Dream1Index
            CodeEdit1.LineDefines(2) = Dream2Index
            CodeEdit1.LineDefines(3) = Dream3Index
            CodeEdit1.LineDefines(4) = Dream4Index
            CodeEdit1.LineDefines(5) = Dream5Index
            CodeEdit1.LineDefines(6) = Dream6Index
 
            Combo1.AddItem "Comment"
            Combo1.AddItem "Commands"
            Combo1.AddItem "Notation"
            Combo1.AddItem "Register"
            Combo1.AddItem "Parameter"
            Combo1.AddItem "Error"
            Combo1.ListIndex = 0
    End Select
    CodeEdit1.Redraw
    
End Sub
Private Sub PickColor(ByVal Index As Variant)
    On Error Resume Next
    
    fMain.CommonDialog1.Color = Me.ColorProperty(Index)
    fMain.CommonDialog1.CancelError = True
    fMain.CommonDialog1.Flags = cdlCCFullOpen Or cdlCCRGBInit
    fMain.CommonDialog1.ShowColor
    
    If Err = 0 Then
        
        Me.ColorProperty(Index) = fMain.CommonDialog1.Color

        Me.Tag = "True"
        Command2(2).Caption = "&Cancel"
        
    Else
        Err.Clear
    End If
    On Error GoTo 0
    
End Sub

Private Sub Command1_Click()
    Select Case pLanguage
        Case "Assembly"
            '+20
            PickColor Combo1.ListIndex + 25
        Case "NSISScript"
            '+14
            PickColor Combo1.ListIndex + 20
        Case "VBScript"
            '+14
            PickColor Combo1.ListIndex + 14
        Case "JScript"
            '+8
            PickColor Combo1.ListIndex + 8
        Case "BatchInk"
            '+3
            PickColor Combo1.ListIndex + 3
            
    End Select
End Sub

Private Sub Command2_Click(Index As Integer)
    Select Case Index
        Case 0 'reset
            Select Case pLanguage
                Case "VBScript"
                    Me.ColorProperty(ColorProperties.VBScriptComment) = fMain.ColorProperty(ColorProperties.VBScriptComment)
                    Me.ColorProperty(ColorProperties.VBScriptStatements) = fMain.ColorProperty(ColorProperties.VBScriptStatements)
                    Me.ColorProperty(ColorProperties.VBScriptOperators) = fMain.ColorProperty(ColorProperties.VBScriptOperators)
                    Me.ColorProperty(ColorProperties.VBScriptVariables) = fMain.ColorProperty(ColorProperties.VBScriptVariables)
                    Me.ColorProperty(ColorProperties.VBScriptValues) = fMain.ColorProperty(ColorProperties.VBScriptValues)
                    Me.ColorProperty(ColorProperties.VBScriptError) = fMain.ColorProperty(ColorProperties.VBScriptError)
                Case "JScript"
                    Me.ColorProperty(ColorProperties.JScriptComment) = fMain.ColorProperty(ColorProperties.JScriptComment)
                    Me.ColorProperty(ColorProperties.JScriptStatements) = fMain.ColorProperty(ColorProperties.JScriptStatements)
                    Me.ColorProperty(ColorProperties.JScriptOperators) = fMain.ColorProperty(ColorProperties.JScriptOperators)
                    Me.ColorProperty(ColorProperties.JScriptVariables) = fMain.ColorProperty(ColorProperties.JScriptVariables)
                    Me.ColorProperty(ColorProperties.JScriptValues) = fMain.ColorProperty(ColorProperties.JScriptValues)
                    Me.ColorProperty(ColorProperties.JScriptError) = fMain.ColorProperty(ColorProperties.JScriptError)
                Case "BatchInk"
                    Me.ColorProperty(ColorProperties.BatchInkComment) = fMain.ColorProperty(ColorProperties.BatchInkComment)
                    Me.ColorProperty(ColorProperties.BatchInkCommands) = fMain.ColorProperty(ColorProperties.BatchInkCommands)
                    Me.ColorProperty(ColorProperties.BatchInkFinished) = fMain.ColorProperty(ColorProperties.BatchInkFinished)
                    Me.ColorProperty(ColorProperties.BatchInkCurrently) = fMain.ColorProperty(ColorProperties.BatchInkCurrently)
                    Me.ColorProperty(ColorProperties.BatchInkIncomming) = fMain.ColorProperty(ColorProperties.BatchInkIncomming)
                Case "NSISScript"
                    Me.ColorProperty(ColorProperties.NSISScriptComment) = fMain.ColorProperty(ColorProperties.NSISScriptComment)
                    Me.ColorProperty(ColorProperties.NSISScriptCommands) = fMain.ColorProperty(ColorProperties.NSISScriptCommands)
                    Me.ColorProperty(ColorProperties.NSISScriptEqualJump) = fMain.ColorProperty(ColorProperties.NSISScriptEqualJump)
                    Me.ColorProperty(ColorProperties.NSISScriptElseJump) = fMain.ColorProperty(ColorProperties.NSISScriptElseJump)
                    Me.ColorProperty(ColorProperties.NSISScriptAboveJump) = fMain.ColorProperty(ColorProperties.NSISScriptAboveJump)
                Case "Assembly"
                    Me.ColorProperty(ColorProperties.AssemblyComment) = fMain.ColorProperty(ColorProperties.AssemblyComment)
                    Me.ColorProperty(ColorProperties.AssemblyCommand) = fMain.ColorProperty(ColorProperties.AssemblyCommand)
                    Me.ColorProperty(ColorProperties.AssemblyNotation) = fMain.ColorProperty(ColorProperties.AssemblyNotation)
                    Me.ColorProperty(ColorProperties.AssemblyRegister) = fMain.ColorProperty(ColorProperties.AssemblyRegister)
                    Me.ColorProperty(ColorProperties.AssemblyParameter) = fMain.ColorProperty(ColorProperties.AssemblyParameter)
                    Me.ColorProperty(ColorProperties.AssemblyError) = fMain.ColorProperty(ColorProperties.AssemblyError)
            End Select
            Me.Tag = "False"
            Command2(2).Caption = "&Close"
        Case 1 'apply
            Select Case pLanguage
                Case "VBScript"
                    fMain.ColorProperty(ColorProperties.VBScriptComment) = Me.ColorProperty(ColorProperties.VBScriptComment)
                    fMain.ColorProperty(ColorProperties.VBScriptStatements) = Me.ColorProperty(ColorProperties.VBScriptStatements)
                    fMain.ColorProperty(ColorProperties.VBScriptOperators) = Me.ColorProperty(ColorProperties.VBScriptOperators)
                    fMain.ColorProperty(ColorProperties.VBScriptVariables) = Me.ColorProperty(ColorProperties.VBScriptVariables)
                    fMain.ColorProperty(ColorProperties.VBScriptValues) = Me.ColorProperty(ColorProperties.VBScriptValues)
                    fMain.ColorProperty(ColorProperties.VBScriptError) = Me.ColorProperty(ColorProperties.VBScriptError)
                Case "JScript"
                    fMain.ColorProperty(ColorProperties.JScriptComment) = Me.ColorProperty(ColorProperties.JScriptComment)
                    fMain.ColorProperty(ColorProperties.JScriptStatements) = Me.ColorProperty(ColorProperties.JScriptStatements)
                    fMain.ColorProperty(ColorProperties.JScriptOperators) = Me.ColorProperty(ColorProperties.JScriptOperators)
                    fMain.ColorProperty(ColorProperties.JScriptVariables) = Me.ColorProperty(ColorProperties.JScriptVariables)
                    fMain.ColorProperty(ColorProperties.JScriptValues) = Me.ColorProperty(ColorProperties.JScriptValues)
                    fMain.ColorProperty(ColorProperties.JScriptError) = Me.ColorProperty(ColorProperties.JScriptError)
                Case "BatchInk"
                    fMain.ColorProperty(ColorProperties.BatchInkComment) = Me.ColorProperty(ColorProperties.BatchInkComment)
                    fMain.ColorProperty(ColorProperties.BatchInkCommands) = Me.ColorProperty(ColorProperties.BatchInkCommands)
                    fMain.ColorProperty(ColorProperties.BatchInkFinished) = Me.ColorProperty(ColorProperties.BatchInkFinished)
                    fMain.ColorProperty(ColorProperties.BatchInkCurrently) = Me.ColorProperty(ColorProperties.BatchInkCurrently)
                    fMain.ColorProperty(ColorProperties.BatchInkIncomming) = Me.ColorProperty(ColorProperties.BatchInkIncomming)
                Case "NSISScript"
                    fMain.ColorProperty(ColorProperties.NSISScriptComment) = Me.ColorProperty(ColorProperties.NSISScriptComment)
                    fMain.ColorProperty(ColorProperties.NSISScriptCommands) = Me.ColorProperty(ColorProperties.NSISScriptCommands)
                    fMain.ColorProperty(ColorProperties.NSISScriptEqualJump) = Me.ColorProperty(ColorProperties.NSISScriptEqualJump)
                    fMain.ColorProperty(ColorProperties.NSISScriptElseJump) = Me.ColorProperty(ColorProperties.NSISScriptElseJump)
                    fMain.ColorProperty(ColorProperties.NSISScriptAboveJump) = Me.ColorProperty(ColorProperties.NSISScriptAboveJump)
                Case "Assembly"
                    fMain.ColorProperty(ColorProperties.AssemblyComment) = Me.ColorProperty(ColorProperties.AssemblyComment)
                    fMain.ColorProperty(ColorProperties.AssemblyCommand) = Me.ColorProperty(ColorProperties.AssemblyCommand)
                    fMain.ColorProperty(ColorProperties.AssemblyNotation) = Me.ColorProperty(ColorProperties.AssemblyNotation)
                    fMain.ColorProperty(ColorProperties.AssemblyRegister) = Me.ColorProperty(ColorProperties.AssemblyRegister)
                    fMain.ColorProperty(ColorProperties.AssemblyParameter) = Me.ColorProperty(ColorProperties.AssemblyParameter)
                    fMain.ColorProperty(ColorProperties.AssemblyError) = Me.ColorProperty(ColorProperties.AssemblyError)
            End Select
            Me.CodeEdit1.Redraw
            fMain.txtMain.Redraw
            Me.Tag = "False"
            Command2(2).Caption = "&Close"
        Case 2 'cancel
            If Me.Tag = "True" Then
                If MsgBox("Are you sure you want to cancel the changes?", vbQuestion + vbYesNo, "Cancel") = vbYes Then
                    Unload Me
                End If
            Else
                Unload Me
            End If
    End Select
End Sub

Private Sub Form_Load()
    Me.Tag = "False"
    Command2(2).Caption = "&Close"
End Sub
