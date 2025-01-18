VERSION 5.00
Begin VB.Form frmColor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Color"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4980
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   185
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   332
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHex 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   19
      Text            =   "0"
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox txtBlue 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   4320
      MaxLength       =   3
      TabIndex        =   18
      Text            =   "0"
      Top             =   1890
      Width           =   375
   End
   Begin VB.TextBox txtGreen 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   4320
      MaxLength       =   3
      TabIndex        =   17
      Text            =   "0"
      Top             =   1590
      Width           =   375
   End
   Begin VB.TextBox txtRed 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   4320
      MaxLength       =   3
      TabIndex        =   16
      Text            =   "0"
      Top             =   1290
      Width           =   375
   End
   Begin VB.TextBox txtBri 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   4320
      MaxLength       =   3
      TabIndex        =   15
      Text            =   "0"
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox txtSat 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   4320
      MaxLength       =   3
      TabIndex        =   14
      Text            =   "0"
      Top             =   420
      Width           =   375
   End
   Begin VB.TextBox txtHue 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   4320
      MaxLength       =   3
      TabIndex        =   13
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox picBrightness 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1995
      Left            =   3600
      ScaleHeight     =   131
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4710
      TabIndex        =   22
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4710
      TabIndex        =   21
      Top             =   420
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "°"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4710
      TabIndex        =   20
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "B:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   11
      Top             =   1890
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "G:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   1590
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "R:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      Top             =   1290
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "B:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "S:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   420
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "H:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Color :"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblCurrColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   2280
      Width           =   375
   End
   Begin VB.Shape selector 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   195
      Left            =   480
      Top             =   240
      Width           =   195
   End
   Begin VB.Label pc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   195
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Color As Long

Const SWIDTH = -1
Const MAXCOL = 18
Const CSTEP = 8
Private ClrMdl As ColorConverter



Private Sub cmdCancel_Click()
    Color = -1
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Color = lblCurrColor.BackColor
    Me.Hide
End Sub

Private Sub Form_Load()
    Color = -1
    Set ClrMdl = New ColorConverter
    pc(0).Left = 8
    pc(0).Top = 8
    selector.Left = pc(0).Left
    selector.Top = pc(0).Top
   
    loadSwatches
'    picBrightness.Left = pc(15).Left + pc(15).Width
    LoadColors 100
   ' If strTitle <> "" Then Me.Caption = strTitle
   ' If lngDefaultColor >= 0 Then lblCurrColor.BackColor = lngDefaultColor
End Sub

Private Sub loadSwatches()
    Dim i, col, step As Integer
    Dim cTop As Long
    Dim R, G, B As Integer
    cTop = pc(0).Top
    Dim hue, sat, brightness As Long
    For i = 1 To 197 Step 1
        DoEvents
        
        
        Load pc(i)
        col = col + 1
        If col < MAXCOL Then
            With pc(i)
                .Left = pc(i - 1).Left + pc(i - 1).Width + SWIDTH
                .Top = cTop
                .Visible = True
            End With
        Else
            col = 0
            cTop = cTop + pc(0).Height + SWIDTH
            With pc(i)
                .Left = pc(0).Left
                .Top = cTop
                .Visible = True
            End With
        End If
    Next i
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Color = -1
        Me.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "Color = " & Color
End Sub

Private Sub pc_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        SelectColor Index
    End If
End Sub

Private Sub LoadColors(brightness As Long)
    Dim i, col, row As Integer
    row = 0
    Dim foundColor As Boolean
    Dim selColIndex As Integer
    Dim hue, sat As Long
    For i = 0 To 197 Step 1
        DoEvents
        col = col + 1
        If col >= MAXCOL Then
            col = 0
            row = row + 1
        End If
        pc(i).BackColor = ClrMdl.HSBToRGB(ClrMdl.HSBToLong(hue, sat, brightness))
        hue = col * 20
        sat = row * 10
    Next i
    selector.Visible = foundColor
    If foundColor Then
        SelectColor selColIndex
    End If
End Sub

Private Sub FadeColor(Color As Long)
    Dim hue, sat As Long
    hue = ClrMdl.hue(ClrMdl.RGBToHSB(Color))
    sat = ClrMdl.Saturation(ClrMdl.RGBToHSB(Color))
'    Debug.Print "## FadeColor , Hue =" & hue & ", Sat = " & sat
    Dim myHeight As Long
    myHeight = picBrightness.ScaleHeight
    Dim i As Integer
    With picBrightness
        For i = 0 To myHeight Step 2
            picBrightness.Line (0, i)-(.Width, i), ClrMdl.HSBToRGB(ClrMdl.HSBToLong(hue, sat, 120 - i))
            picBrightness.Line (0, i + 1)-(.Width, i + 1), ClrMdl.HSBToRGB(ClrMdl.HSBToLong(hue, sat, 120 - i))
        Next i
    End With
End Sub

Private Sub SelectColor(myIndex As Integer)
    'picBrightness.BackColor = Color
    With pc(myIndex)
        selector.Left = .Left
        selector.Top = .Top
    End With
    selector.Visible = True
    lblCurrColor.BackColor = pc(myIndex).BackColor
    FadeColor pc(myIndex).BackColor
    UpdateTextValues pc(myIndex).BackColor
End Sub


Private Sub picBrightness_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        lblCurrColor.BackColor = picBrightness.Point(X, Y)
        UpdateTextValues lblCurrColor.BackColor
        findcolor lblCurrColor.BackColor
    End If
End Sub

Private Sub UpdateTextValues(myColor As Long)
    With ClrMdl
        txtHue.Text = .hue(.RGBToHSB(myColor))
        txtSat.Text = .Saturation(.RGBToHSB(myColor))
        txtBri.Text = .brightness(.RGBToHSB(myColor))
        txtRed.Text = .Red(myColor)
        txtGreen.Text = .Green(myColor)
        txtBlue.Text = .Blue(myColor)
        txtHex.Text = GetHexString(myColor)
    End With
End Sub

Private Sub UpdateRGBValues(myColor As Long)
    With ClrMdl
        txtRed.Text = .Red(myColor)
        txtGreen.Text = .Green(myColor)
        txtBlue.Text = .Blue(myColor)
        txtHex.Text = GetHexString(myColor)
    End With
End Sub

Private Sub UpdateHSBValues(myColor As Long)
    With ClrMdl
        txtHue.Text = .hue(.RGBToHSB(myColor))
        txtSat.Text = .Saturation(.RGBToHSB(myColor))
        txtBri.Text = .brightness(.RGBToHSB(myColor))
        txtHex.Text = GetHexString(myColor)
    End With
End Sub

Private Sub UpdateHSBRGBValues(myColor As Long)
    With ClrMdl
        txtHue.Text = .hue(.RGBToHSB(myColor))
        txtSat.Text = .Saturation(.RGBToHSB(myColor))
        txtBri.Text = .brightness(.RGBToHSB(myColor))
        txtRed.Text = .Red(myColor)
        txtGreen.Text = .Green(myColor)
        txtBlue.Text = .Blue(myColor)
    End With
End Sub


Private Function GetHexString(clr As Long) As String
    Dim tmpval As String
    tmpval = HFormat(Hex(clr))
    GetHexString = Mid(tmpval, 5, 2) & Mid(tmpval, 3, 2) & Mid(tmpval, 1, 2)
End Function


Private Function HFormat(hstring As String) As String
    Dim i As Integer
    Dim retval As String
    For i = Len(hstring) To 5 Step 1
        retval = retval & "0"
    Next i
    HFormat = retval & hstring
End Function

Private Function SelBox(tbox As TextBox)
    tbox.SelStart = 0
    tbox.SelLength = Len(tbox.Text)
End Function

Private Sub ValidateText(tbox As TextBox, MaxVal As Integer, Optional MinVal As Integer = 0)
    Dim myVal As Long
    If IsNumeric(tbox.Text) Then
        myVal = CLng(tbox.Text)
        If myVal > MaxVal Then tbox.Text = MaxVal
        If myVal < MinVal Then tbox.Text = MinVal
    Else
        tbox.Text = MinVal
    End If
End Sub


Private Sub txtHex_KeyUp(KeyCode As Integer, Shift As Integer)
    UpdateFromHex
End Sub

Private Sub txtHex_LostFocus()
    txtHex.Text = HFormat(txtHex.Text)
End Sub

Private Sub txtHue_KeyUp(KeyCode As Integer, Shift As Integer)
    ValidateText txtHue, 360
    UpdateHSBFromText
End Sub


Private Sub txtSat_KeyUp(KeyCode As Integer, Shift As Integer)
    ValidateText txtSat, 100
    UpdateHSBFromText
End Sub

Private Sub txtBri_KeyUp(KeyCode As Integer, Shift As Integer)
    ValidateText txtBri, 100
    UpdateHSBFromText
End Sub

Private Sub txtRed_KeyUp(KeyCode As Integer, Shift As Integer)
    ValidateText txtRed, 255
    UpdateRGBFromText
End Sub

Private Sub txtGreen_KeyUp(KeyCode As Integer, Shift As Integer)
    ValidateText txtGreen, 255
    UpdateRGBFromText
End Sub

Private Sub txtBlue_KeyUp(KeyCode As Integer, Shift As Integer)
    ValidateText txtBlue, 255
    UpdateRGBFromText
End Sub


Private Sub txtHue_GotFocus()
    SelBox txtHue
End Sub

Private Sub txtSat_GotFocus()
    SelBox txtSat
End Sub

Private Sub txtBri_GotFocus()
    SelBox txtBri
End Sub

Private Sub txtRed_GotFocus()
    SelBox txtRed
End Sub

Private Sub txtGreen_GotFocus()
    SelBox txtGreen
End Sub

Private Sub txtBlue_GotFocus()
    SelBox txtBlue
End Sub

Private Sub txtHex_GotFocus()
    SelBox txtHex
End Sub

Private Sub UpdateHSBFromText()
    lblCurrColor.BackColor = ClrMdl.HSBToRGB(ClrMdl.HSBToLong(CLng(txtHue.Text), CLng(txtSat.Text), CLng(txtBri.Text)))
    FadeColor lblCurrColor.BackColor
    UpdateRGBValues lblCurrColor.BackColor
    findcolor lblCurrColor.BackColor
End Sub

Private Sub UpdateRGBFromText()
    lblCurrColor.BackColor = RGB(CLng(txtRed.Text), CLng(txtGreen.Text), CLng(txtBlue.Text))
    FadeColor lblCurrColor.BackColor
    UpdateHSBValues lblCurrColor.BackColor
    findcolor lblCurrColor.BackColor
End Sub

Private Sub UpdateFromHex()
    lblCurrColor.BackColor = GetColor(txtHex.Text)
    FadeColor lblCurrColor.BackColor
    UpdateHSBRGBValues lblCurrColor.BackColor
    findcolor lblCurrColor.BackColor
End Sub


Private Sub findcolor(myColor As Long)
    Dim i As Integer
    selector.Visible = False
    For i = 0 To 197 Step 1
        If pc(i).BackColor = myColor Then
            selector.Left = pc(i).Left
            selector.Top = pc(i).Top
            selector.Visible = True
            Exit For
        End If
    Next i
End Sub

Private Function GetColor(hexstring As String) As Long
    Dim pos As Integer
    Dim redstring, greenstring, bluestring As String
    Dim R, G, B As Long
    pos = InStr(1, hexstring, "#")
    'Debug.Print pos
    hexstring = Mid(hexstring, pos + 1)
    redstring = Mid(hexstring, 1, 2)
    greenstring = Mid(hexstring, 3, 2)
    bluestring = Mid(hexstring, 5, 2)
    R = Val("&H" & redstring)
    G = Val("&H" & greenstring)
    B = Val("&H" & bluestring)
    'Debug.Print "####HEXSTRING = " & hexstring
    GetColor = RGB(R, G, B)
End Function

