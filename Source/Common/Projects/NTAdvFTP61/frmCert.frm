VERSION 5.00
Begin VB.Form frmCert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Certificate"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   Icon            =   "frmCert.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Remember my choice (until a restart)"
      Height          =   195
      Left            =   150
      TabIndex        =   12
      Top             =   2880
      Visible         =   0   'False
      Width           =   3030
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Accept"
      Height          =   375
      Left            =   3540
      TabIndex        =   11
      Top             =   2820
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   2820
      Width           =   1470
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      Height          =   2625
      Left            =   105
      ScaleHeight     =   2565
      ScaleWidth      =   6465
      TabIndex        =   0
      Top             =   105
      Width           =   6525
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
      Begin VB.Line Line2 
         X1              =   195
         X2              =   6180
         Y1              =   1950
         Y2              =   1950
      End
      Begin VB.Label Serial 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   240
         Left            =   855
         TabIndex        =   9
         Top             =   390
         Width           =   5325
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
         Left            =   3900
         TabIndex        =   7
         Top             =   2145
         Width           =   2220
      End
      Begin VB.Label BeginDate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   1245
         TabIndex        =   6
         Top             =   2145
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
         Left            =   3555
         TabIndex        =   5
         Top             =   2130
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
         Left            =   300
         TabIndex        =   4
         Top             =   2130
         Width           =   930
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   225
         Picture         =   "frmCert.frx":0442
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
    

   ' Subject.Text = cert.Terms(CInt(CertificateFields.Subject))
   ' Issuer.Text = cert.Terms(CInt(CertificateFields.Issuer))
    
    Dim cnt As Long

    cnt = 1
    Do While cert.Exists("ID_" & (CInt(CertificateFields.Subject) + cnt))
        Subject.Text = Subject.Text & cert.Terms(CInt(CertificateFields.Subject) + cnt) & ", "
        cnt = cnt + 1
    Loop
    If Len(Subject.Text) > 2 Then Subject.Text = Left(Subject.Text, Len(Subject.Text) - 2)
    cnt = 1
    Do While cert.Exists("ID_" & (CInt(CertificateFields.Issuer) + cnt))
        Issuer.Text = Issuer.Text & cert.Terms(CInt(CertificateFields.Issuer) + cnt) & ", "
        cnt = cnt + 1
    Loop
    If Len(Issuer.Text) > 2 Then Issuer.Text = Left(Issuer.Text, Len(Issuer.Text) - 2)
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

