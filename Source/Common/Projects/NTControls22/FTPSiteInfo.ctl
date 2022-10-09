VERSION 5.00
Begin VB.UserControl FTPSiteInfo 
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6060
   LockControls    =   -1  'True
   ScaleHeight     =   1695
   ScaleWidth      =   6060
   ToolboxBitmap   =   "FTPSiteInfo.ctx":0000
   Begin VB.Frame Frame1 
      Height          =   1635
      Left            =   75
      TabIndex        =   6
      Top             =   -15
      Width           =   5925
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2340
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1155
         Width           =   2115
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   4545
         MaxLength       =   11
         TabIndex        =   11
         Text            =   "10000-20000"
         Top             =   1155
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Save Login"
         Height          =   240
         Left            =   2325
         TabIndex        =   3
         Top             =   780
         Width           =   1125
      End
      Begin NTControls22.URLBox URLBox1 
         Height          =   315
         Left            =   555
         TabIndex        =   0
         Top             =   300
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   556
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   0
         Left            =   555
         TabIndex        =   1
         Top             =   750
         Width           =   1665
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4545
         TabIndex        =   4
         Text            =   "21"
         Top             =   750
         Width           =   480
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   555
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1170
         Width           =   1665
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Pasv"
         Height          =   240
         Left            =   5115
         TabIndex        =   5
         Top             =   795
         Width           =   675
      End
      Begin VB.Label Label4 
         Caption         =   "Port"
         Height          =   195
         Left            =   4170
         TabIndex        =   10
         Top             =   810
         Width           =   315
      End
      Begin VB.Label Label3 
         Caption         =   "Pass"
         Height          =   255
         Left            =   135
         TabIndex        =   9
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "User"
         Height          =   255
         Left            =   135
         TabIndex        =   8
         Top             =   795
         Width           =   360
      End
      Begin VB.Label Label1 
         Caption         =   "URL"
         Height          =   210
         Left            =   135
         TabIndex        =   7
         Top             =   360
         Width           =   390
      End
   End
End
Attribute VB_Name = "FTPSiteInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary

Public sHostURL As Object
Public sUserName As Object
Public sPassword As Object
Public sPort As Object
Public sSavePass As Object
Public sPassive As Object
Public sAdapter As Object
Public sPortRange As Object

Private tmpRange As String

Public Event Change()
Public Event KeyDown(ByVal KeyCode As Integer)

Public Property Get AutoTypeCombo()
    Set AutoTypeCombo = URLBox1.AutoTypeCombo
End Property
Public Property Let Caption(ByVal newVal As String)
    Frame1.Caption = newVal
End Property
Public Property Get Caption() As String
    Caption = Frame1.Caption
End Property

Public Sub Reset()
    
    Check1.Value = 0
    Check3.Value = 0
    
    sHostURL.Text = ""
    sPort.Text = "21"
    sUserName.Text = ""
    sPassword.Text = ""
    
    sPortRange.Text = "10000-20000"
    
    If Combo1.ListCount > 0 Then Combo1.ListIndex = 0

End Sub
Public Property Let ShowAdvSettings(ByVal newVal As Boolean)

    Label4.Left = IIf(newVal, 4170, 4920)
    Text3.Left = IIf(newVal, 4545, 5295)

    Combo1.Visible = newVal
    Check3.Visible = newVal
    Text1.Visible = newVal
    
End Property
Public Property Get ShowAdvSettings() As Boolean
    ShowAdvSettings = Check3.Visible
End Property

Public Sub Refresh()
    SetPasswordBox
End Sub

Private Function IsHostName(ByVal sHost As String) As Boolean
    Dim isHost As Boolean
    Dim tPos As Integer
    isHost = False
    tPos = InStr(sHost, ".")
    If tPos > 0 Then
        tPos = InStr(tPos, sHost, ".")
        If tPos > 0 Then
            tPos = InStr(tPos, sHost, ".")
            If tPos > 0 Then
                tPos = InStr(sHost, "/")
                If tPos = 0 Then
                    tPos = InStr(sHost, "\")
                    If tPos = 0 Then
                        isHost = True
                    End If
                End If
            End If
        End If
    End If
    
    IsHostName = isHost
    
End Function
Public Function GetUserName(ByVal theURL As String) As String

Dim Login As String
theURL = Trim(theURL)
If (Left(theURL, 6) = "ftp://" Or Left(theURL, 6) = "ntp://") Or Left(theURL, 7) = "http://" Then
    Dim testURL As String
    If Left(theURL, 7) = "http://" Then
        testURL = Mid(theURL, 8)
    Else
        testURL = Mid(theURL, 7)
        End If
    If InStr(testURL, "@") > 0 Then
        testURL = Left(testURL, InStr(testURL, "@") - 1)
        If InStr(testURL, ":") > 0 Then
            Login = Left(testURL, InStr(testURL, ":") - 1)
        Else
            Login = testURL
            End If
    Else
        Login = ""
        End If
Else
    Login = ""
    End If
GetUserName = Login

End Function
Public Function GetPassword(ByVal theURL As String) As String

Dim Password As String
theURL = Trim(theURL)
If (Left(theURL, 6) = "ftp://" Or Left(theURL, 6) = "ntp://") Or Left(theURL, 7) = "http://" Then
    Dim testURL As String
    If Left(theURL, 7) = "http://" Then
        testURL = Mid(theURL, 8)
    Else
        testURL = Mid(theURL, 7)
        End If
    If InStr(testURL, "@") > 0 Then
        testURL = Left(testURL, InStr(testURL, "@") - 1)
        If InStr(testURL, ":") > 0 Then
            Password = Mid(testURL, InStr(testURL, ":") + 1)
        Else
            Password = ""
            End If
    Else
        Password = ""
        End If
Else
    Password = ""
    End If
GetPassword = Password

End Function
Public Function GetType(ByVal theURL As String) As Integer

Dim whatsURL As Integer
If (LCase(Left(Trim(theURL), 6)) = "ftp://" Or LCase(Left(Trim(theURL), 6)) = "ntp://") Then
    whatsURL = 3
Else
    If (LCase(Left(Trim(theURL), 7)) = "http://") Or _
        Left(Trim(LCase(theURL)), 3) = "www" Or _
        Right(Trim(LCase(theURL)), 4) = ".txt" Or _
        Right(Trim(LCase(theURL)), 4) = ".htm" Or _
        Right(Trim(LCase(theURL)), 5) = ".html" Or _
        Right(Trim(LCase(theURL)), 4) = ".asp" Then

        whatsURL = 4
    Else
        If LCase(Left(Trim(theURL), 8)) = "file:///" Or LCase(Mid(Trim(theURL), 2, 1)) = ":" Then
            whatsURL = 1
        Else
            If LCase(Left(Trim(theURL), 7)) = "file://" Or LCase(Left(Trim(theURL), 2)) = "\\" Then
                whatsURL = 2
            Else
                whatsURL = 0
            End If
        End If
    End If
End If
GetType = whatsURL

End Function
Private Sub SetPasswordBox()
    If ((GetUserName(URLBox1.Text) <> "") Or (GetPassword(URLBox1.Text) <> "")) Then
    
        Label2.Enabled = False
        Label3.Enabled = False
        sUserName.Enabled = False
        sPassword.Enabled = False
        sUserName.Locked = True
        sPassword.Locked = True
        
        Combo1.Enabled = False
        Check1.Enabled = False
    
        Label4.Enabled = True
        Text3.Enabled = True
        Check3.Enabled = True

        Text1.Enabled = True
        
        sUserName.Text = GetUserName(URLBox1.Text)
        sPassword.Text = GetPassword(URLBox1.Text)

    ElseIf (GetType(URLBox1.Text) = 3 Or IsHostName(URLBox1.Text)) Then
        
        If Not sUserName.Enabled And sUserName.Text <> "" Then sUserName.Text = ""
        If Not sPassword.Enabled And sPassword.Text <> "" Then sPassword.Text = ""
        
        Label2.Enabled = True
        Label3.Enabled = True
        sUserName.Enabled = True
        sPassword.Enabled = True
        sUserName.Locked = False
        sPassword.Locked = False
        
        Combo1.Enabled = True
        Check1.Enabled = True
        
        Label4.Enabled = True
        Text3.Enabled = True
        Check3.Enabled = True
        
        Text1.Enabled = True
        
    Else
    
        Label2.Enabled = False
        Label3.Enabled = False
        sUserName.Enabled = False
        sPassword.Enabled = False
        sUserName.Locked = True
        sPassword.Locked = True
    
        Combo1.Enabled = False
        Check1.Enabled = False
    
        Label4.Enabled = False
        Text3.Enabled = False
        Check3.Enabled = False
        Text1.Enabled = False
        
        sUserName.Text = ""
        sPassword.Text = ""

    End If
    
    If Check3.Enabled Then

        Combo1.Enabled = (Check3.Value = 0)
        Text1.Enabled = (Check3.Value = 0)

    End If
    
End Sub

Private Sub Check3_Click()
    SetPasswordBox
End Sub

Private Sub URLBox1_Change()
    SetPasswordBox
    RaiseEvent Change
End Sub

Public Property Get hWnd() As Long
    On Error Resume Next
    hWnd = UserControl.Parent.hWnd
    Err.Clear
End Property

Private Sub UserControl_Initialize()
    
    Set sHostURL = URLBox1
    Set sUserName = Text2(0)
    Set sPassword = Text2(1)
    Set sPort = Text3
    Set sSavePass = Check1
    Set sPassive = Check3
    Set sAdapter = Combo1
    Set sPortRange = Text1
    
    Dim col As Collection
    Dim ftp As Object
    Set ftp = CreateObject("NTAdvFTP61.Client")
    Set col = ftp.alladapters
    Set ftp = Nothing
    
    If col.Count > 0 Then
        Dim cnt As Long
        For cnt = 1 To col.Count
            Combo1.AddItem col.Item(cnt)
        Next
    Else
        Combo1.AddItem "(Unknown)"
    End If

End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 6060
    UserControl.Height = 1695
End Sub

Private Sub UserControl_Terminate()

    Set sHostURL = Nothing
    Set sUserName = Nothing
    Set sPassword = Nothing
    Set sPort = Nothing
    Set sSavePass = Nothing
    Set sPassive = Nothing
    Set sAdapter = Nothing
    Set sPortRange = Nothing

End Sub

