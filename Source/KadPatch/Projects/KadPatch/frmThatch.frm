VERSION 5.00
Object = "{BA98913A-7219-4720-8E5D-F3D8E058DF1B}#323.0#0"; "NTImaging10.ocx"
Begin VB.Form frmThatch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Project Matting Size"
   ClientHeight    =   5055
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   3285
   ClipControls    =   0   'False
   Icon            =   "frmThatch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   3285
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   288
      Left            =   7815
      TabIndex        =   8
      Text            =   "4"
      Top             =   435
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   336
      Left            =   2190
      TabIndex        =   7
      Top             =   4545
      Width           =   936
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Okay"
      Default         =   -1  'True
      Height          =   336
      Left            =   2190
      TabIndex        =   6
      Top             =   4200
      Width           =   936
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   288
      Left            =   2130
      TabIndex        =   4
      Text            =   "96"
      Top             =   285
      Width           =   708
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   288
      Left            =   795
      TabIndex        =   2
      Text            =   "96"
      Top             =   285
      Width           =   708
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   8925
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   435
      Visible         =   0   'False
      Width           =   705
   End
   Begin NTImaging10.Gallery Gallery1 
      Height          =   4230
      Left            =   90
      TabIndex        =   11
      Top             =   660
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   7461
      FilePath        =   ""
      Stretch         =   -1  'True
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   2355
      Left            =   75
      TabIndex        =   13
      Top             =   1425
      Visible         =   0   'False
      Width           =   3120
   End
   Begin VB.Label Label6 
      Caption         =   "Number of blocks, and thatch back."
      Height          =   225
      Left            =   75
      TabIndex        =   12
      Top             =   15
      Width           =   3165
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      ForeColor       =   &H80000010&
      Height          =   540
      Left            =   5700
      TabIndex        =   10
      Top             =   825
      Visible         =   0   'False
      Width           =   6165
   End
   Begin VB.Label Label5 
      Caption         =   "Square block measurement:"
      Height          =   165
      Left            =   5715
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Label Label3 
      Caption         =   "in"
      Height          =   210
      Left            =   8640
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label2 
      Caption         =   "width:"
      Height          =   165
      Left            =   1605
      TabIndex        =   3
      Top             =   330
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "height:"
      Height          =   255
      Left            =   165
      TabIndex        =   0
      Top             =   330
      Width           =   540
   End
End
Attribute VB_Name = "frmThatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If Not (IsNumeric(Text1.Text) And IsNumeric(Text2.Text)) And Text1.Visible Then
        MsgBox "The values for Height and Width must numerical and above zero and below five hundred and tweleve."
    Else
        If (Not ((CSng(Text1.Text) > 0) And (CSng(Text2.Text) > 0) And (CSng(Text1.Text) <= 512) And (CSng(Text2.Text) <= 512))) And Text1.Visible Then
            MsgBox "The values for Height and Width must numerical and above zero and below five hundred and tweleve."
        Else
'            If BlockRow(CSng(Text1.Text), CSng(Text3.Text)) <= 0 Or BlockCol(CSng(Text2.Text), CSng(Text3.Text)) <= 0 Then
'                MsgBox "The values you entered did not yield a row and/or a column count of at least one block or more."
'            Else
                Tag = "OK"
                Visible = False
                
'            End If
        End If
    End If
End Sub

Private Sub Command2_Click()
    Tag = False
    Visible = False
End Sub

Private Sub Form_Initialize()
    frmThatch.Tag = ""
End Sub

Public Sub SetupExport()
    Me.Caption = "Project Display Size"
    Text1.Text = 6
    Text2.Text = 6
    Label6.Caption = "Enter the display picture size in inches"
    Label7.Caption = "A PDF file will be created with two pages:" & vbCrLf & vbCrLf & "The first page will be the visual rendering of the project with the above defined width and height in inches." & vbCrLf & vbCrLf & "The second page will be a absolute sized (or greater) black and white rendering of the projet in pattern symbols."
    Label1.Visible = True
    Label7.Visible = True
    Label2.Visible = True
    Text1.Visible = True
    Text2.Visible = True
    Gallery1.Visible = False
    'Command1.Top = Command1.Top - 3500
    'Command2.Top = Command2.Top - 3500
    'Me.Height = Me.Height - 3500
    
    Combo1.AddItem "mm" '1/16th of a inch square
    Combo1.AddItem "in" 'always round up to cover
    Combo1.AddItem "cm"
    Combo1.ListIndex = 0
    Approximate

'        Gallery1.Orientation = 0
'    Gallery1.Width = Me.ScaleWidth - (Gallery1.Left * 2)
End Sub

Public Sub SetupThatch()
    Me.Caption = "Project Matting Size"
    Label6.Caption = "Number of blocks, and thatch back."

    Label1.Visible = True
    Label2.Visible = True
    Label7.Visible = False
    Text1.Visible = True
    Text2.Visible = True
    
    Combo1.AddItem "mm" '1/16th of a inch square
    Combo1.AddItem "in" 'always round up to cover
    Combo1.AddItem "cm"
    Combo1.ListIndex = 0
    Approximate

    Gallery1.Stretch = True
    Gallery1.FilePath = AppPath & "Base\Stitchings\Mattings"
    Gallery1.ListIndex = ThatchIndex - 1
'        Gallery1.Orientation = 0
'    Gallery1.Width = Me.ScaleWidth - (Gallery1.Left * 2)
End Sub

Public Sub SetupSymbols()
    Me.Caption = "Set Symbol Color"
    Label6.Caption = "Select a color to be used with the symbol."
    Label1.Visible = False
    Label7.Visible = False
    Label2.Visible = False
    Text1.Visible = False
    Text2.Visible = False

    Gallery1.Stretch = True
    Gallery1.FilePath = AppPath & "Base\Stitchings\FlossThreads"
    'Gallery1.ListIndex = ColorIndex - 1
        
'        Gallery1.Orientation = 0
'    Gallery1.Width = Me.ScaleWidth - (Gallery1.Left * 2)
End Sub

Public Sub SetupColors()
    Me.Caption = "Set Color Symbol"
    Label6.Caption = "Select a symbol to be used with the color."
    Label1.Visible = False
    Label2.Visible = False
    Text1.Visible = False
    Text2.Visible = False


    Gallery1.Stretch = False
    Gallery1.FilePath = AppPath & "Base\Stitchings\LegendKeys"
        
'        Gallery1.Orientation = 0
'    Gallery1.Width = Me.ScaleWidth - (Gallery1.Left * 2)
End Sub


Public Property Get ThatchWidth() As Single
    ThatchWidth = CSng(Text2.Text)
End Property
Public Property Get ThatchHeight() As Single
    ThatchHeight = CSng(Text1.Text)
End Property
Public Property Get BlockSize() As Single
    BlockSize = CSng(Text3.Text)
End Property
Public Property Get ThatchIndex() As Byte
    ThatchIndex = (Gallery1.ListIndex + 1)
    
End Property

Public Sub SetProperties(ByVal ToForm As Boolean)

    Dim w1 As Single
    Dim w2 As Single
    Dim h1 As Single
    Dim h2 As Single
    Dim b1 As Single
    Dim b2 As Single
    Dim t1 As String
    Dim t2 As String
    Select Case Combo1.ListIndex
        Case 0 'millimeters
            If IsNumeric(Text1.Text) And IsNumeric(Text2.Text) Then
                b1 = 0.0393701
                t1 = "inches"
                b2 = 0.1
                t2 = "centimeters"
            End If
        Case 1 'inches
            If IsNumeric(Text1.Text) And IsNumeric(Text2.Text) Then
                b1 = 25.4
                t1 = "millimeters"
                b2 = 2.54
                t2 = "centimeters"
            End If
        Case 2 'centimeters
            If IsNumeric(Text1.Text) And IsNumeric(Text2.Text) Then
                b1 = 0.393701
                t1 = "inches"
                b2 = 10
                t2 = "millimeters"
            End If
    End Select
    If t1 <> "" Then
        w1 = Round((CSng(Text1.Text) * b1), 2)
        h1 = Round((CSng(Text2.Text) * b1), 2)
        b1 = Round(w1 * h1 * b1, 2)
        w2 = Round((CDbl(Text1.Text) * b2), 2)
        h2 = Round((CDbl(Text2.Text) * b2), 2)
        b2 = Round(w2 * h2 * b2, 2)
        Label4.Caption = ((t1 * w1) / h1) & " Blocks In " & t1 & ":  Width " & w1 & ",  Height " & h1 & ",  Block Size " & b1 & "." & vbCrLf
        Label4.Caption = Label4.Caption & ((t2 * w2) / h2) & " Blocks In " & t2 & ":  Width " & w2 & ",  Height " & h2 & ",  Block Size " & b2 & "." & vbCrLf
    End If


'    If ToForm Then
'        If modProj.Height > modProj.Blocks Then
'
'            Debug.Print ((modProj.Header.Scalar * 65355) / modProj.Header.Blocks)
'            Debug.Print ((65355 * modProj.Header.Blocks) / modProj.Header.Height)
'        Else
'            Debug.Print ((modProj.Header.Scalar * modProj.Header.Height) / modProj.Header.Blocks)
'            Debug.Print ((modProj.Header.Height * modProj.Header.Blocks) / 65355)
'
'        End If
'        Debug.Print
'
'    '    Text1.Text = modProj.Header.Height
'    '    Text2.Text = modProj.Header.Width
'    '    Text3.Text = modProj.Header.Resolute
'        Combo1.ListIndex = modProj.Header.Scalar
'    Else
'        If CSng(Text1.Text) > CSng(Text2.Text) Then
'
'    '    modProj.Header.Width = Text1.Text
'    '    modProj.Header.Height = Text2.Text
'    '    modProj.Header.Resolute = Text3.Text
'        modProj.Header.Scalar = Combo1.ListIndex
'        End If
'    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 And Visible Then
        Tag = ""
        Visible = False
        Cancel = True
    End If
End Sub

Private Sub Approximate()
    Dim w1 As Single
    Dim w2 As Single
    Dim h1 As Single
    Dim h2 As Single
    Dim b1 As Single
    Dim b2 As Single
    Dim t1 As String
    Dim t2 As String
    Select Case Combo1.ListIndex
        Case 0 'millimeters
            If IsNumeric(Text1.Text) And IsNumeric(Text2.Text) Then
                b1 = 0.0393701
                t1 = "inches"
                b2 = 0.1
                t2 = "centimeters"
            End If
        Case 1 'inches
            If IsNumeric(Text1.Text) And IsNumeric(Text2.Text) Then
                b1 = 25.4
                t1 = "millimeters"
                b2 = 2.54
                t2 = "centimeters"
            End If
        Case 2 'centimeters
            If IsNumeric(Text1.Text) And IsNumeric(Text2.Text) Then
                b1 = 0.393701
                t1 = "inches"
                b2 = 10
                t2 = "millimeters"
            End If
    End Select
    If t1 <> "" Then
        w1 = Round((CSng(Text1.Text) * b1), 2)
        h1 = Round((CSng(Text2.Text) * b1), 2)
        b1 = Round(Round(w1 * h1, 2) * b1, 2)
        w2 = Round((CDbl(Text1.Text) * b2), 2)
        h2 = Round((CDbl(Text2.Text) * b2), 2)
        b2 = Round(Round(w2 * h2, 2) * b2, 2)
        On Error Resume Next
        Dim temp As String
        temp = Round(((h1 * w1) * b1), 0) + IIf((h1 * w1) Mod b1 <> 0, 1, 0) & " Blocks In " & t1 & ":  Width " & w1 & ",  Height " & h1 & ",  Blocks " & b1 & "." & vbCrLf
        If Err Then
            temp = ""
            Err.Clear
        End If
        Label4.Caption = temp
        
        temp = Round(((h2 * w2) * b2), 0) + IIf((h2 * w2) Mod b2 <> 0, 1, 0) & " Blocks In " & t2 & ":  Width " & w2 & ",  Height " & h2 & ",  Blocks " & b2 & "." & vbCrLf
        If Err Then
            temp = ""
            Err.Clear
        End If
        Label4.Caption = Label4.Caption & temp
    End If
End Sub


Private Sub Text1_Change()
    Approximate
End Sub

Private Sub Text2_Change()
    Approximate
End Sub

Private Sub Combo1_Click()
    Approximate
End Sub





