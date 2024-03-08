VERSION 5.00
Begin VB.Form frmCert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Certificate"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   Icon            =   "frmCert.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Remember my choice (until a restart)"
      Height          =   210
      Left            =   150
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   3030
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Accept"
      Height          =   375
      Left            =   3555
      TabIndex        =   11
      Top             =   3420
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5175
      TabIndex        =   10
      Top             =   3420
      Width           =   1470
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      Height          =   3240
      Left            =   105
      ScaleHeight     =   3180
      ScaleWidth      =   6465
      TabIndex        =   0
      Top             =   105
      Width           =   6525
      Begin VB.PictureBox KeyUsage 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   7
         Left            =   6165
         Picture         =   "frmCert.frx":0442
         ScaleHeight     =   337.5
         ScaleMode       =   0  'User
         ScaleWidth      =   187.5
         TabIndex        =   23
         ToolTipText     =   "TrustList"
         Top             =   1935
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.PictureBox KeyUsage 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   6
         Left            =   5790
         Picture         =   "frmCert.frx":0784
         ScaleHeight     =   337.5
         ScaleMode       =   0  'User
         ScaleWidth      =   187.5
         TabIndex        =   22
         ToolTipText     =   "Microsoft Trust List Signing"
         Top             =   1935
         Width           =   225
      End
      Begin VB.PictureBox KeyUsage 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   5
         Left            =   5130
         Picture         =   "frmCert.frx":0AC6
         ScaleHeight     =   337.5
         ScaleMode       =   0  'User
         ScaleWidth      =   187.5
         TabIndex        =   21
         ToolTipText     =   "Encrypting File System"
         Top             =   1935
         Width           =   225
      End
      Begin VB.PictureBox KeyUsage 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   4
         Left            =   4410
         Picture         =   "frmCert.frx":0E08
         ScaleHeight     =   337.5
         ScaleMode       =   0  'User
         ScaleWidth      =   187.5
         TabIndex        =   20
         ToolTipText     =   "Time Stamping"
         Top             =   1935
         Width           =   225
      End
      Begin VB.PictureBox KeyUsage 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   3
         Left            =   3720
         Picture         =   "frmCert.frx":114A
         ScaleHeight     =   337.5
         ScaleMode       =   0  'User
         ScaleWidth      =   187.5
         TabIndex        =   19
         ToolTipText     =   "Secure Email"
         Top             =   1935
         Width           =   225
      End
      Begin VB.PictureBox KeyUsage 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   2
         Left            =   3030
         Picture         =   "frmCert.frx":148C
         ScaleHeight     =   337.5
         ScaleMode       =   0  'User
         ScaleWidth      =   187.5
         TabIndex        =   18
         ToolTipText     =   "Code Signing"
         Top             =   1935
         Width           =   225
      End
      Begin VB.PictureBox KeyUsage 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   1
         Left            =   2310
         Picture         =   "frmCert.frx":17CE
         ScaleHeight     =   337.5
         ScaleMode       =   0  'User
         ScaleWidth      =   187.5
         TabIndex        =   17
         ToolTipText     =   "Client Authentication"
         Top             =   1935
         Width           =   225
      End
      Begin VB.PictureBox KeyUsage 
         BackColor       =   &H80000014&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   0
         Left            =   1620
         Picture         =   "frmCert.frx":1B10
         ScaleHeight     =   337.5
         ScaleMode       =   0  'User
         ScaleWidth      =   187.5
         TabIndex        =   16
         ToolTipText     =   "Server Authentication"
         Top             =   1935
         Width           =   225
      End
      Begin VB.TextBox Issuer 
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   1215
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   1365
         Width           =   4965
      End
      Begin VB.TextBox Subject 
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   840
         Width           =   4965
      End
      Begin VB.Label Ciphers 
         BackStyle       =   0  'Transparent
         Height          =   270
         Left            =   1590
         TabIndex        =   24
         Top             =   2280
         Width           =   4440
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Key Usage:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   315
         TabIndex        =   15
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Line Line2 
         X1              =   195
         X2              =   6180
         Y1              =   2625
         Y2              =   2625
      End
      Begin VB.Label Serial 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00 00 00 00"
         Height          =   240
         Left            =   750
         TabIndex        =   9
         Top             =   345
         Width           =   5355
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Serial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   5205
         TabIndex        =   8
         Top             =   120
         Width           =   930
      End
      Begin VB.Label EndDate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3975
         TabIndex        =   7
         Top             =   2745
         Width           =   2220
      End
      Begin VB.Label BeginDate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   2745
         Width           =   2160
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "to"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3630
         TabIndex        =   5
         Top             =   2730
         Width           =   285
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Valid from"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   375
         TabIndex        =   4
         Top             =   2730
         Width           =   930
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   225
         Picture         =   "frmCert.frx":1E52
         Top             =   135
         Width           =   480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Server Certificate Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   105
         Width           =   2730
      End
      Begin VB.Line Line1 
         X1              =   195
         X2              =   6180
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Issued to:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         TabIndex        =   2
         Top             =   840
         Width           =   930
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Issued by:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         TabIndex        =   1
         Top             =   1365
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmCert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub ViewCert(ByRef cert As Certificate)
    
    Dim cnt As Long

    cnt = 1
    Do While cert.Exists("ID_" & (CInt(CertificateFields.Subject) + cnt))
        Subject.Text = Subject.Text & cert.Terms(CInt(CertificateFields.Subject) + cnt) & ", "
        cnt = cnt + 1
    Loop
    If Len(Subject.Text) > 2 Then Subject.Text = Left(Subject.Text, Len(Subject.Text) - 2)


    cnt = 1
    Do While cert.Exists("ID_" & (CInt(CertificateFields.Issuer) + cnt))
        If Left(cert.Namely(("ID_" & (CInt(CertificateFields.Issuer) + cnt))), 7) = "Subject" Then

        Else
            Issuer.Text = Issuer.Text & cert.Terms(CInt(CertificateFields.Issuer) + cnt) & ", "

        End If
        cnt = cnt + 1
    Loop
    If Len(Issuer.Text) > 2 Then Issuer.Text = Left(Issuer.Text, Len(Issuer.Text) - 2)
    
    Ciphers.Caption = cert.RSAKeySize & "bit"
        
    If cert.Exists("ID_" & (CInt(CertificateFields.SignatureAlgorithm) + 1)) Then
        Ciphers.Caption = Ciphers.Caption & " Signature-Algorithm,"
    ElseIf cert.Exists("ID_" & (CInt(CertificateFields.Algorithm) + 1)) Then
        Ciphers.Caption = Ciphers.Caption & " Algorithm,"
    End If

    If cert.Terms(CInt(CertificateFields.PublicKeyBlock)) <> "" Or cert.Terms(CertificateFields.Signature) <> "" Then
    
        If cert.Terms(CInt(CertificateFields.PublicKeyBlock)) <> "" Then
            Ciphers.Caption = Ciphers.Caption & " Public Key,"
        End If
    
        If cert.Terms(CertificateFields.Signature) <> "" Then
            Ciphers.Caption = Ciphers.Caption & " Signature,"
        End If
    End If
    If Right(Ciphers.Caption, 1) = "," Then Ciphers.Caption = Left(Ciphers.Caption, Len(Ciphers.Caption) - 1)
        
    BeginDate.Caption = cert.Terms(ValidityBeginDate)
    EndDate.Caption = cert.Terms(ValidityExpireDate)
    Serial.Caption = cert.Terms(SerialNumber)
    

    Me.Show

    Do
        DoLoop
    Loop Until Me.Visible = False
    
End Sub
Public Function CheckCert(ByRef cert As Certificate) As Boolean
    
    Command2.Visible = True
    Command1.Caption = "&Decline"
    Check1.Visible = True

    
    ViewCert cert
    
    cert.Accepted = (Me.Tag = "ACCEPT")
    cert.NoPrompt = (Check1.Value = 1)

    
    CheckCert = cert.Accepted
End Function


Private Sub Command1_Click()
    Me.Tag = "DECLINE"
    Me.Hide
    
End Sub

Private Sub Command2_Click()
    Me.Tag = "ACCEPT"
    Me.Hide
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 1 Then
        Me.Tag = "CANCEL"
        Me.Hide
    End If
End Sub

