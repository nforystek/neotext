VERSION 5.00
Object = "{C98B112F-745F-4542-B5B3-DDFADF1F6E2F}#1180.0#0"; "NTControls22.ocx"
Begin VB.Form frmFavoriteSite 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Favorite Site Information"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   Icon            =   "frmFavoriteSite.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   7500
   Visible         =   0   'False
   Begin VB.CheckBox Check1 
      Caption         =   "Secure to me"
      Height          =   225
      Left            =   6135
      TabIndex        =   4
      Top             =   915
      Width           =   1275
   End
   Begin NTControls22.SiteInformation SiteInformation2 
      Height          =   1695
      Left            =   15
      TabIndex        =   1
      Top             =   1695
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   2990
   End
   Begin NTControls22.SiteInformation SiteInformation1 
      Height          =   1695
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   2990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   345
      Index           =   0
      Left            =   6105
      TabIndex        =   2
      Top             =   90
      Width           =   1305
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   345
      Index           =   1
      Left            =   6105
      TabIndex        =   3
      Top             =   480
      Width           =   1305
   End
End
Attribute VB_Name = "frmFavoriteSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private LastFileName As String

Public Property Get FileName() As String
    FileName = LastFileName
End Property

Public Sub LoadSite(ByVal FileName As String)
    Dim enc As New NTCipher10.ncode

    LastFileName = FileName

    Dim FileNum As Integer
    
    SiteInformation1.sHostURL.Text = ""
    SiteInformation1.sPort.Text = "21"
    SiteInformation1.sPassive.Value = BoolToCheck(dbSettings.GetProfileSetting("ConnectionMode") = 0)
    SiteInformation1.sPortRange.Text = dbSettings.GetProfileSetting("DefaultPortRange")
    SiteInformation1.sUserName.Text = ""
    SiteInformation1.sPassword.Text = ""
    SiteInformation1.sSavePass.Value = 0
    SiteInformation1.sSSL.Value = 0
    SiteInformation1.sAdapter.ListIndex = (dbSettings.GetProfileSetting("AdapterIndex") - 1)
    
    SiteInformation2.sHostURL.Text = ""
    SiteInformation2.sPort.Text = "21"
    SiteInformation2.sPassive.Value = BoolToCheck(dbSettings.GetProfileSetting("ConnectionMode") = 0)
    SiteInformation2.sPortRange.Text = dbSettings.GetProfileSetting("DefaultPortRange")
    SiteInformation2.sUserName.Text = ""
    SiteInformation2.sPassword.Text = ""
    SiteInformation2.sSavePass.Value = 0
    SiteInformation2.sSSL.Value = 0
    SiteInformation2.sAdapter.ListIndex = (dbSettings.GetProfileSetting("AdapterIndex") - 1)
    
    Check1.Value = 0
                
    FileNum = FreeFile
    Open FileName For Input As #FileNum
        Dim InData As String
        Dim inVar As String
        Do Until EOF(FileNum)
            Line Input #FileNum, InData
            If (InStr(InData, "=") = 0) And (Not (Left(InData, 1) = "'")) Then
                Check1.Value = 1
                InData = enc.DecryptString(InData, dbSettings.CryptKey)
            End If
            If Left$(InData, 1) <> "'" Then
                If InStr(InData, "=") > 0 Then
                    inVar = Trim(LCase(Left(InData, InStr(InData, "=") - 1)))
                    InData = Mid(InData, InStr(InData, "=") + 1)

                    Select Case inVar
                    
                    Case "plocation"
                        SiteInformation1.sHostURL.Text = Trim(InData)
                    Case "slocation"
                        SiteInformation2.sHostURL.Text = Trim(InData)
                        
                    Case "pport"
                        SiteInformation1.sPort.Text = Trim(InData)
                    Case "sport"
                        SiteInformation2.sPort.Text = Trim(InData)
                    
                    Case "ppassive"
                        SiteInformation1.sPassive.Value = CInt(InData)
                    Case "spassive"
                        SiteInformation2.sPassive.Value = CInt(InData)
                    
                    Case "pportrange"
                        SiteInformation1.sPortRange.Text = Trim(InData)
                    Case "sportrange"
                        SiteInformation2.sPortRange.Text = Trim(InData)
                    
                    Case "padapter"
                        SiteInformation1.sAdapter.ListIndex = CInt(InData) - 1
                    Case "sadapter"
                        SiteInformation2.sAdapter.ListIndex = CInt(InData) - 1
                    
                    Case "psavepass"
                        SiteInformation1.sSavePass.Value = CInt(InData)
                    Case "ssavepass"
                        SiteInformation2.sSavePass.Value = CInt(InData)
                    
                    Case "plogin"
                        If Not InData = "" Then SiteInformation1.sUserName.Text = IIf(Check1.Value = 1, enc.DecryptString(InData, dbSettings.CryptKey), InData)
                    Case "slogin"
                        If Not InData = "" Then SiteInformation2.sUserName.Text = IIf(Check1.Value = 1, enc.DecryptString(InData, dbSettings.CryptKey), InData)
                    
                    Case "ppassword"
                        If Not InData = "" Then SiteInformation1.sPassword.Text = IIf(Check1.Value = 1, enc.DecryptString(InData, dbSettings.CryptKey(SiteInformation1.sUserName.Text)), InData)
                    Case "spassword"
                        If Not InData = "" Then SiteInformation2.sPassword.Text = IIf(Check1.Value = 1, enc.DecryptString(InData, dbSettings.CryptKey(SiteInformation2.sUserName.Text)), InData)
                    
                    Case "pssl"
                        SiteInformation1.sSSL.Value = CInt(InData)
                    Case "sssl"
                        SiteInformation2.sSSL.Value = CInt(InData)
                    
                    End Select

                    End If
                End If
            Loop
    Close #FileNum

    Set enc = Nothing
End Sub

Public Sub SaveSite(ByVal FileName As String)
    Dim enc As New NTCipher10.ncode

    Dim FileNum As Integer
    FileNum = FreeFile
    Open FileName For Output As #FileNum
    If Check1.Value = 0 Then
        Print #FileNum, "pLocation =" & SiteInformation1.sHostURL.Text
        Print #FileNum, "sLocation =" & SiteInformation2.sHostURL.Text
    
        Print #FileNum, "pPort =" & SiteInformation1.sPort.Text
        Print #FileNum, "sPort =" & SiteInformation2.sPort.Text
        
        Print #FileNum, "pPassive =" & CStr(SiteInformation1.sPassive.Value)
        Print #FileNum, "sPassive =" & CStr(SiteInformation2.sPassive.Value)
        
        Print #FileNum, "pPortRange =" & CStr(SiteInformation1.sPortRange.Text)
        Print #FileNum, "sPortRange =" & CStr(SiteInformation2.sPortRange.Text)
        
        Print #FileNum, "pAdapter =" & CStr(SiteInformation1.sAdapter.ListIndex + 1)
        Print #FileNum, "sAdapter =" & CStr(SiteInformation2.sAdapter.ListIndex + 1)
        
        Print #FileNum, "pSavePass =" & CStr(SiteInformation1.sSavePass.Value)
        Print #FileNum, "sSavePass =" & CStr(SiteInformation2.sSavePass.Value)
        
        If SiteInformation1.sSavePass.Value = 1 Then
            If (SiteInformation1.sUserName.Text = "") Then
                Print #FileNum, "pLogin ="
            Else
                Print #FileNum, "pLogin =" & SiteInformation1.sUserName.Text
            End If
        End If
        If SiteInformation2.sSavePass.Value = 1 Then

            If (SiteInformation2.sUserName.Text = "") Then
                Print #FileNum, "sLogin ="
            Else
                Print #FileNum, "sLogin =" & SiteInformation2.sUserName.Text
            End If

        End If
        
        If SiteInformation1.sSavePass.Value = 1 Then

            If (SiteInformation1.sPassword.Text = "") Then
                Print #FileNum, "pPassword ="
            Else
                Print #FileNum, "pPassword =" & SiteInformation1.sPassword.Text
            End If

        End If
        If SiteInformation2.sSavePass.Value = 1 Then
            If (SiteInformation2.sPassword.Text = "") Then
                Print #FileNum, "sPassword ="
            Else
                Print #FileNum, "sPassword =" & SiteInformation2.sPassword.Text
            End If
        End If
        
        Print #FileNum, "pSSL =" & CStr(SiteInformation1.sSSL.Value)
        Print #FileNum, "sSSL =" & CStr(SiteInformation2.sSSL.Value)
    Else
        Print #FileNum, enc.EncryptString("pLocation =" & SiteInformation1.sHostURL.Text, dbSettings.CryptKey)
        Print #FileNum, enc.EncryptString("sLocation =" & SiteInformation2.sHostURL.Text, dbSettings.CryptKey)
    
        Print #FileNum, enc.EncryptString("pPort =" & SiteInformation1.sPort.Text, dbSettings.CryptKey)
        Print #FileNum, enc.EncryptString("sPort =" & SiteInformation2.sPort.Text, dbSettings.CryptKey)
        
        Print #FileNum, enc.EncryptString("pPassive =" & CStr(SiteInformation1.sPassive.Value), dbSettings.CryptKey)
        Print #FileNum, enc.EncryptString("sPassive =" & CStr(SiteInformation2.sPassive.Value), dbSettings.CryptKey)
        
        Print #FileNum, enc.EncryptString("pPortRange =" & CStr(SiteInformation1.sPortRange.Text), dbSettings.CryptKey)
        Print #FileNum, enc.EncryptString("sPortRange =" & CStr(SiteInformation2.sPortRange.Text), dbSettings.CryptKey)
        
        Print #FileNum, enc.EncryptString("pAdapter =" & CStr(SiteInformation1.sAdapter.ListIndex + 1), dbSettings.CryptKey)
        Print #FileNum, enc.EncryptString("sAdapter =" & CStr(SiteInformation2.sAdapter.ListIndex + 1), dbSettings.CryptKey)
        
        Print #FileNum, enc.EncryptString("pSavePass =" & CStr(SiteInformation1.sSavePass.Value), dbSettings.CryptKey)
        Print #FileNum, enc.EncryptString("sSavePass =" & CStr(SiteInformation2.sSavePass.Value), dbSettings.CryptKey)
        
        If SiteInformation1.sSavePass.Value = 1 Then
            If (SiteInformation1.sUserName.Text = "") Then
                Print #FileNum, enc.EncryptString("pLogin =", dbSettings.CryptKey)
            Else
                Print #FileNum, enc.EncryptString("pLogin =" & enc.EncryptString(SiteInformation1.sUserName.Text, dbSettings.CryptKey), dbSettings.CryptKey)
            End If
        End If
        If SiteInformation2.sSavePass.Value = 1 Then
            If (SiteInformation2.sUserName.Text = "") Then
                Print #FileNum, enc.EncryptString("sLogin =", dbSettings.CryptKey)
            Else
                Print #FileNum, enc.EncryptString("sLogin =" & enc.EncryptString(SiteInformation2.sUserName.Text, dbSettings.CryptKey), dbSettings.CryptKey)
            End If
        End If
        
        If SiteInformation1.sSavePass.Value = 1 Then
            If (SiteInformation1.sPassword.Text = "") Then
                Print #FileNum, enc.EncryptString("pPassword =", dbSettings.CryptKey)
            Else
                Print #FileNum, enc.EncryptString("pPassword =" & enc.EncryptString(SiteInformation1.sPassword.Text, dbSettings.CryptKey(SiteInformation1.sUserName.Text)), dbSettings.CryptKey)
            End If
        End If
        If SiteInformation2.sSavePass.Value = 1 Then
            If (SiteInformation2.sPassword.Text = "") Then
                Print #FileNum, enc.EncryptString("sPassword =", dbSettings.CryptKey)
            Else
                Print #FileNum, enc.EncryptString("sPassword =" & enc.EncryptString(SiteInformation2.sPassword.Text, dbSettings.CryptKey(SiteInformation2.sUserName.Text)), dbSettings.CryptKey)
            End If
        End If
        
        Print #FileNum, enc.EncryptString("pSSL =" & CStr(SiteInformation1.sSSL.Value), dbSettings.CryptKey)
        Print #FileNum, enc.EncryptString("sSSL =" & CStr(SiteInformation2.sSSL.Value), dbSettings.CryptKey)
    End If
    
    Close #FileNum

    Set enc = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
    frmMain.ValidDataPortRange SiteInformation1.sPortRange
    frmMain.ValidDataPortRange SiteInformation2.sPortRange
    Select Case Index
        Case 0
            If LastFileName <> "" Then SaveSite LastFileName
    End Select
    Unload Me
End Sub

Private Sub Form_Load()
    
    Me.Left = ((Screen.Width / 2) - (Me.Width / 2))
    Me.Top = ((Screen.Height / 2) - (Me.Height / 2))
    Dim frm As Form
    For Each frm In Forms
        If (TypeName(frm) = TypeName(Me)) Then
            If Not (frm.hwnd = Me.hwnd) Then
                Me.Left = frm.Left + (Screen.TwipsPerPixelX * 32)
                Me.Top = frm.Top + (Screen.TwipsPerPixelY * 32)
            End If
        End If
    Next
    
    If ((Me.Left + Me.Width) > Screen.Width) Or _
        ((Me.Top + Me.Height) > Screen.Height) Then
        Me.Left = (32 * Screen.TwipsPerPixelX)
        Me.Top = (32 * Screen.TwipsPerPixelY)
    End If
        
    SetAutoTypeList Me, SiteInformation1.AutoTypeCombo
    SetAutoTypeList Me, SiteInformation2.AutoTypeCombo

    SiteInformation1.sPassive.Value = BoolToCheck(dbSettings.GetProfileSetting("ConnectionMode") = 0)
    SiteInformation2.sPassive.Value = BoolToCheck(dbSettings.GetProfileSetting("ConnectionMode") = 0)

    SiteInformation1.sSSL.Value = dbSettings.GetProfileSetting("SSL")
    SiteInformation2.sSSL.Value = dbSettings.GetProfileSetting("SSL")
    
    SiteInformation1.sPortRange.Text = dbSettings.GetProfileSetting("DefaultPortRange")
    SiteInformation2.sPortRange.Text = dbSettings.GetProfileSetting("DefaultPortRange")
    
    SiteInformation1.sAdapter.ListIndex = (dbSettings.GetProfileSetting("AdapterIndex") - 1)
    SiteInformation2.sAdapter.ListIndex = (dbSettings.GetProfileSetting("AdapterIndex") - 1)
    
    SiteInformation1.ShowAdvSettings = dbSettings.GetProfileSetting("ShowAdvSettings")
    If Not SiteInformation1.ShowAdvSettings Then
        SiteInformation1.sPassive.Value = IIf((dbSettings.GetProfileSetting("ConnectionMode") = 0), 1, 0)
    End If
    SiteInformation2.ShowAdvSettings = dbSettings.GetProfileSetting("ShowAdvSettings")
    If Not SiteInformation2.ShowAdvSettings Then
        SiteInformation2.sPassive.Value = IIf((dbSettings.GetProfileSetting("ConnectionMode") = 0), 1, 0)
    End If
    
End Sub
