VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test"
   ClientHeight    =   6000
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8490
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "SSL"
      Height          =   195
      Left            =   7080
      TabIndex        =   17
      Top             =   360
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Cancel"
      Height          =   252
      Left            =   5700
      TabIndex        =   15
      Top             =   96
      Width           =   720
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   6516
      ScaleHeight     =   240
      ScaleWidth      =   1830
      TabIndex        =   10
      Top             =   120
      Width           =   1824
      Begin VB.OptionButton Option4 
         Caption         =   "Passive"
         Height          =   195
         Left            =   768
         TabIndex        =   14
         Top             =   0
         Value           =   -1  'True
         Width           =   945
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Active"
         Height          =   195
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   768
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1560
      ScaleHeight     =   240
      ScaleWidth      =   2190
      TabIndex        =   9
      Top             =   372
      Width           =   2196
      Begin VB.OptionButton Option2 
         Caption         =   "Large File"
         Height          =   195
         Left            =   1068
         TabIndex        =   12
         Top             =   0
         Width           =   1044
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Small File"
         Height          =   195
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Value           =   -1  'True
         Width           =   1044
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Disconnect"
      Height          =   252
      Left            =   4656
      TabIndex        =   1
      Top             =   90
      Width           =   1050
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Local List"
      Height          =   252
      Left            =   3732
      TabIndex        =   7
      Top             =   90
      Width           =   930
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Upload"
      Height          =   252
      Left            =   2928
      TabIndex        =   5
      Top             =   90
      Width           =   810
   End
   Begin VB.TextBox Text2 
      Height          =   2595
      Left            =   132
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   3285
      Width           =   8220
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Local"
      Height          =   252
      Left            =   2292
      TabIndex        =   6
      Top             =   90
      Width           =   645
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Download"
      Height          =   252
      Left            =   1368
      TabIndex        =   4
      Top             =   90
      Width           =   930
   End
   Begin VB.CommandButton Command3 
      Caption         =   "List"
      Height          =   252
      Left            =   924
      TabIndex        =   3
      Top             =   90
      Width           =   450
   End
   Begin VB.TextBox Text1 
      Height          =   2556
      Left            =   132
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   615
      Width           =   8220
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   810
   End
   Begin VB.Label Label1 
      Height          =   240
      Left            =   3720
      TabIndex        =   16
      Top             =   360
      Width           =   2700
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
'TOP DOWN

Public WithEvents loc As NTAdvFTP61.client
Attribute loc.VB_VarHelpID = -1
Public WithEvents hdd As NTAdvFTP61.client
Attribute hdd.VB_VarHelpID = -1
Public WithEvents ftp As NTAdvFTP61.client
Attribute ftp.VB_VarHelpID = -1


Private Sub Check1_Click()
    
    ftp.ImplicitSSL = (Check1.Value = 1)
End Sub

Private Sub Form_Initialize()
On Error GoTo catch

    Set ftp = New NTAdvFTP61.client
    Set loc = New NTAdvFTP61.client
    Set hdd = New NTAdvFTP61.client


    ftp.username = "anonymous"
    ftp.password = "maxftp@neotext.org"
    ftp.DataPortRange = "3000-6000"

    
    Exit Sub
catch:
    DebugText "Error: " & Err.Description
    Err.Clear
End Sub


Private Sub Command1_Click()
On Error GoTo catch
    
    If Check1.Value = 1 Then
        ftp.ImplicitSSL = True
        ftp.Connect "ftps://www.neotext.org:990/upload"
    Else
        ftp.ImplicitSSL = False
        ftp.Connect "ftp://www.neotext.org:21/upload"
    End If
    loc.Connect "C:\Windows\Temp"
    hdd.Connect "C:\Temp"
    ftp.ConnectionMode = IIf(Option3.Value, ConnectionModes.active, ConnectionModes.Passive)

    ftp.ListContents "C:\WINDOWS\TEMP\LIST.TXT"
    Exit Sub
catch:
    DebugText "Error: " & Err.Description
    Err.Clear
End Sub

Private Sub Command2_Click()
On Error GoTo catch

    ftp.Disconnect
    loc.Disconnect
    hdd.Disconnect
    
    Exit Sub
catch:
    DebugText "Error: " & Err.Description
    Err.Clear
End Sub

Private Sub Command3_Click()
On Error GoTo catch

    Text2.Text = "(Listing Contents)"
    ftp.ConnectionMode = IIf(Option3.Value, ConnectionModes.active, ConnectionModes.Passive)

    ftp.ListContents "C:\WINDOWS\TEMP\LIST.TXT"

    Exit Sub
catch:
    DebugText "Error: " & Err.Description
    Err.Clear

End Sub

Private Sub Command4_Click()
On Error GoTo catch

    ftp.ConnectionMode = IIf(Option3.Value, ConnectionModes.active, ConnectionModes.Passive)
    If Option2.Value Then
        ftp.transferfile "TEST.tmp", loc
    Else
        ftp.transferfile "Max-FTP v6.1.0.exe", loc
    End If
    Exit Sub
catch:
    DebugText "Error: " & Err.Description
    Err.Clear

End Sub

Private Sub Command5_Click()
On Error GoTo catch

    ftp.ConnectionMode = IIf(Option3.Value, ConnectionModes.active, ConnectionModes.Passive)
    If Option2.Value Then
        loc.transferfile "TEST.tmp", ftp
    Else
        loc.transferfile "Max-FTP v6.1.0.exe", ftp
    End If
    Exit Sub
catch:
    DebugText "Error: " & Err.Description
    Err.Clear

End Sub

Private Sub Command6_Click()
On Error GoTo catch

    ftp.ConnectionMode = IIf(Option3.Value, ConnectionModes.active, ConnectionModes.Passive)
    If Option2.Value Then
        hdd.transferfile "TEST.tmp", loc
    Else
        hdd.transferfile "Max-FTP v6.1.0.exe", loc
    End If
    Exit Sub
catch:
    DebugText "Error: " & Err.Description
    Err.Clear

End Sub

Private Sub Command7_Click()
On Error GoTo catch
    Text2.Text = "(Listing Contents)"
    ftp.ConnectionMode = IIf(Option3.Value, ConnectionModes.active, ConnectionModes.Passive)
    hdd.ListContents "C:\WINDOWS\TEMP\LIST.TXT"
    Exit Sub
catch:
    DebugText "Error: " & Err.Description
    Err.Clear

End Sub

Private Sub Command8_Click()
On Error GoTo catch

    loc.CancelTransfer

    hdd.CancelTransfer
    
    ftp.CancelTransfer


    Exit Sub
catch:
    DebugText "Error: " & Err.Description
    Err.Clear

End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo catch
    If loc.transfering Then loc.CancelTransfer
    If hdd.transfering Then hdd.CancelTransfer
    If ftp.transfering Then ftp.CancelTransfer

    Exit Sub
catch:
    DebugText "Error: " & Err.Description
    Err.Clear
End Sub

Private Sub Form_Terminate()
On Error GoTo catch

    Set ftp = Nothing
    Set loc = Nothing
    Set hdd = Nothing
    Exit Sub
catch:
    DebugText "Error: " & Err.Description
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo catch

    If ftp.ConnectedState Then ftp.Disconnect
    If loc.ConnectedState Then loc.Disconnect
    If hdd.ConnectedState Then hdd.Disconnect
    Exit Sub
catch:
    DebugText "Error: " & Err.Description
    Err.Clear
End Sub

Private Sub ftp_DataComplete(ByVal ProgressType As NTAdvFTP61.ProgressTypes)
On Error GoTo catch

    DebugText "ftp_DataComplete " & ProgressType & " " & ftp.transfering
    If ProgressType = NTAdvFTP61.ProgressTypes.FileListing Then GetListing ftp
    Exit Sub
catch:
    DebugText "Error: " & Err.Description
    Err.Clear
End Sub
Private Sub GetListing(ByRef obj As client)
On Error GoTo catch

    Dim col As New Collection
    Dim itm As Variant
    obj.parselisting ReadFile("C:\WINDOWS\TEMP\LIST.TXT"), col
    Text2.Text = ""
    For Each itm In col
        Text2.Text = Text2.Text & itm & vbCrLf
    Next
    Exit Sub
catch:
    DebugText "Error: " & Err.Description
    Err.Clear
End Sub
Private Sub ftp_DataProgress(ByVal ProgressType As NTAdvFTP61.ProgressTypes, ByVal ReceivedBytes As Double)
On Error GoTo catch

    DebugText "ftp_DataProgress " & ProgressType & " " & ReceivedBytes & " " & ftp.transfering
    Exit Sub
catch:
    DebugText "Error: " & Err.Description
    Err.Clear
End Sub

Private Sub ftp_Error(ByVal Number As Long, ByVal Source As String, ByVal Description As String)
On Error GoTo catch

    DebugText "Error: " & Description
    
    Exit Sub
catch:
    DebugText "Error: " & Err.Description
    Err.Clear
End Sub

Private Sub ftp_LogMessage(ByVal MessageType As NTAdvFTP61.MessageTypes, ByVal AddedText As String)
On Error GoTo catch

    DebugText AddedText
    
    Exit Sub
catch:
    DebugText "Error: " & Err.Description
    Err.Clear
End Sub

Private Sub hdd_DataComplete(ByVal ProgressType As NTAdvFTP61.ProgressTypes)
On Error GoTo catch

    DebugText "hdd_DataComplete " & ProgressType & " " & hdd.transfering
    If ProgressType = NTAdvFTP61.ProgressTypes.FileListing Then GetListing hdd
    Exit Sub
catch:
    DebugText "Error: " & Err.Description
    Err.Clear
End Sub

Private Sub hdd_DataProgress(ByVal ProgressType As NTAdvFTP61.ProgressTypes, ByVal ReceivedBytes As Double)
On Error GoTo catch

    DebugText "hdd_DataProgress " & ProgressType & " " & ReceivedBytes & " " & hdd.transfering
    
    Exit Sub
catch:
    DebugText "Error: " & Err.Description
    Err.Clear
End Sub

Private Sub hdd_Error(ByVal Number As Long, ByVal Source As String, ByVal Description As String)
On Error GoTo catch

    DebugText "Error: " & Description
    
    Exit Sub
catch:
    DebugText "Error: " & Err.Description
    Err.Clear
End Sub

Private Sub loc_DataComplete(ByVal ProgressType As NTAdvFTP61.ProgressTypes)
On Error GoTo catch

    DebugText "loc_DataComplete " & ProgressType & " " & loc.transfering
    If ProgressType = NTAdvFTP61.ProgressTypes.FileListing Then GetListing loc
    Exit Sub
catch:
    DebugText "Error: " & Err.Description
    Err.Clear
End Sub

Private Sub loc_DataProgress(ByVal ProgressType As NTAdvFTP61.ProgressTypes, ByVal ReceivedBytes As Double)
On Error GoTo catch

    DebugText "loc_DataProgress " & ProgressType & " " & ReceivedBytes & " " & loc.transfering
    Exit Sub
catch:
    DebugText "Error: " & Err.Description
    Err.Clear
End Sub

Private Sub loc_Error(ByVal Number As Long, ByVal Source As String, ByVal Description As String)
On Error GoTo catch

    DebugText "Error: " & Description
    
    Exit Sub
catch:
    DebugText "Error: " & Err.Description
    Err.Clear
End Sub

Private Sub DebugText(ByVal str As String)
    Label1.Caption = ftp.transfering & " " & hdd.transfering & " " & loc.transfering

    Dim cnt As Long
    cnt = CountWord(Text1.Text, vbCrLf)
    Text1.Text = Text1.Text & vbCrLf & str
    str = Trim(Replace(Text1.Text, Chr(0), ""))
    If cnt > 1000 Then
        RemoveNextArg str, vbCrLf
        Text1.Text = str
    End If
    Text1.SelStart = InStrRev(Text1.Text, vbCrLf) + 2

    DoEvents
End Sub
