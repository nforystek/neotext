VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   4410
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   2295
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "frmSetup.frx":000C
      Top             =   1710
      Width           =   4170
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   345
      Index           =   1
      Left            =   3435
      TabIndex        =   2
      Top             =   90
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "New"
      Height          =   345
      Index           =   0
      Left            =   2640
      TabIndex        =   1
      Top             =   90
      Width           =   705
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   660
      TabIndex        =   0
      Text            =   "Combo2"
      Top             =   105
      Width           =   1860
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Tablet Controls"
      Height          =   255
      Left            =   105
      TabIndex        =   3
      Top             =   615
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Sound FX"
      Height          =   195
      Left            =   3300
      TabIndex        =   5
      Top             =   615
      Value           =   1  'Checked
      Width           =   1035
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Ambient"
      Height          =   195
      Left            =   2295
      TabIndex        =   4
      Top             =   630
      Value           =   1  'Checked
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   360
      Left            =   3210
      TabIndex        =   9
      Top             =   4170
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Default         =   -1  'True
      Height          =   360
      Left            =   2070
      TabIndex        =   8
      Top             =   4170
      Width           =   1065
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Full Screen"
      Height          =   225
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2340
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1050
      Width           =   1950
   End
   Begin VB.Label Label4 
      Caption         =   "Persistant Bindings:"
      Height          =   270
      Left            =   150
      TabIndex        =   14
      Top             =   1455
      Width           =   1605
   End
   Begin VB.Label Label3 
      Caption         =   "Souund:"
      Height          =   225
      Left            =   1590
      TabIndex        =   12
      Top             =   630
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Profile:"
      Height          =   225
      Left            =   120
      TabIndex        =   11
      Top             =   150
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Resolution:"
      Height          =   240
      Left            =   1485
      TabIndex        =   10
      Top             =   1095
      Width           =   885
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public Property Get Play() As Boolean
    Play = (Me.Tag = "Play")
End Property

Private Sub Check1_Click()
    FullScreen = (Check1.Value = 1)
    db.dbQuery "UPDATE Profiles SET Windowed = " & IIf(Check1.Value = 1, "No", "Yes") & " WHERE Profilename = '" & Replace(Combo2.Text, "'", "''") & "';"
End Sub

Private Sub Check4_Click()
    Tablet = (Check4.Value = 1)
    db.dbQuery "UPDATE Profiles SET Tablet = " & IIf(Check4.Value = 1, "Yes", "No") & " WHERE Profilename = '" & Replace(Combo2.Text, "'", "''") & "';"
End Sub

Private Sub Check2_Click()
    AmbientFX = (Check2.Value = 1)
    db.dbQuery "UPDATE Profiles SET Ambient = " & IIf(Check2.Value = 1, "No", "Yes") & " WHERE Profilename = '" & Replace(Combo2.Text, "'", "''") & "';"
End Sub

Private Sub Check3_Click()
    SoundFX = (Check3.Value = 1)
    db.dbQuery "UPDATE Profiles SET SoundFX = " & IIf(Check3.Value = 1, "No", "Yes") & " WHERE Profilename = '" & Replace(Combo2.Text, "'", "''") & "';"
End Sub

Private Sub Combo1_Click()
    Resolution = Combo1.List(Combo1.ListIndex)
    db.dbQuery "UPDATE Profiles SET Resolution = '" & Resolution & "' WHERE Profilename = '" & Replace(Combo2.Text, "'", "''") & "';"
End Sub

'Private Sub Combo2_Change()
''    Debug.Print "Change " & Combo2.Text & " " & Combo2.ListIndex
'End Sub

Private Sub Combo2_Click()
   ' Debug.Print "Click " & Combo2.Text & " " & Combo2.ListIndex
    If Not RefreshRecord Then
        
    End If
End Sub

Private Sub Command1_Click()
    db.rsQuery rs, "SELECT * FROM Profiles WHERE Profilename = '" & Replace(Combo2.Text, "'", "''") & "';"
    If db.rsEnd(rs) Then
        MsgBox "Please select a value profile to play.", vbExclamation
    Else
                    
        Me.Tag = "Play"
        Me.Visible = False
    End If
End Sub

Private Sub Command2_Click()
    Me.Tag = "Close"
    Me.Visible = False
End Sub

Private Sub EnumDispModes()

    Combo1.Clear
    
    Dim dx As DirectX8
    Dim D3D As Direct3D8
    
    Set dx = New DirectX8
    Set D3D = dx.Direct3DCreate
    
    Dim i As Integer
    For i = 0 To D3D.GetAdapterModeCount(D3DADAPTER_DEFAULT) - 1
        D3D.EnumAdapterModes D3DADAPTER_DEFAULT, i, Display
        '32 bit pixel format
        If Display.Format = D3DFMT_R8G8B8 Or Display.Format = D3DFMT_X8R8G8B8 Or Display.Format = D3DFMT_A8R8G8B8 Then
            If (D3D.CheckDeviceType(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Display.Format, Display.Format, False) >= 0) And _
                (Not (Combo1.List(Combo1.ListCount - 1) = Display.Width & "x" & Display.Height)) Then
                Combo1.AddItem Display.Width & "x" & Display.Height
            End If
            '16 bit pixel format
          Else
            If (D3D.CheckDeviceType(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Display.Format, Display.Format, False) >= 0) And _
                (Not (Combo1.List(Combo1.ListCount - 1) = Display.Width & "x" & Display.Height)) Then
                Combo1.AddItem Display.Width & "x" & Display.Height
            End If
        End If
    Next i
    
    Set dx = Nothing
    Set D3D = Nothing
    
End Sub
Public Sub EnumProfiles()
    db.rsQuery rs, "SELECT * FROM Profiles;"
    If Not db.rsEnd(rs) Then
        Combo2.Clear
        Do While Not db.rsEnd(rs)
    
            Combo2.AddItem rs("Profilename")
            
            rs.MoveNext
        Loop
    End If
End Sub

Private Sub Command3_Click(Index As Integer)
    Select Case Index
        Case 0
            If Combo2.Text = "" Then Combo2.Text = Replace(GetUserLoginName, "'", "''")
            
            If Combo2.Text <> "" Then
                If IsAlphaNumeric(Replace(Combo2.Text, " ", "")) Then
                    'Debug.Print "SELECT * FROM Profiles WHERE Profilename = '" & Replace(Combo2.Text, "'", "''") & "';"
                    db.rsQuery rs, "SELECT * FROM Profiles WHERE Profilename = '" & Replace(Combo2.Text, "'", "''") & "';"
                    If db.rsEnd(rs) Then
                        NewProfile Replace(Combo2.Text, "'", "''")
                        'Debug.Print "UPDATE Users SET Profilename='" & Replace(Combo2.Text, "'", "''") & "' WHERE Username='" & Replace(GetUserLoginName, "'", "''") & "';"
                        db.dbQuery "UPDATE Users SET Profilename='" & Replace(Combo2.Text, "'", "''") & "' WHERE Username='" & Replace(GetUserLoginName, "'", "''") & "';"
                        Combo2.AddItem Combo2.Text
                        MsgBox "Profile """ & Combo2.Text & """ has been created.", vbExclamation
                    Else
                        MsgBox "The profile name already exists.", vbExclamation
                    End If
                Else
                    MsgBox "Please choose a profile name that is of alphabet, numeric and space only.", vbExclamation
                    Combo2.SetFocus
                End If
            Else
                MsgBox "Please enter a profile name in the text entry left of the new button.", vbExclamation
                Combo2.SetFocus
                Combo1.SelStart = 0
                Combo1.SelLength = Len(Combo2.Text)
            End If
        Case 1
            'Debug.Print "SELECT * FROM Profiles WHERE Profilename='" & Replace(Combo2.Text, "'", "''") & "';"
            db.rsQuery rs, "SELECT * FROM Profiles WHERE Profilename='" & Replace(Combo2.Text, "'", "''") & "';"
            If Not db.rsEnd(rs) Then
                'Debug.Print "DELETE FROM Profiles WHERE Profilename='" & Replace(Combo2.Text, "'", "''") & "';"
                db.dbQuery "DELETE FROM Profiles WHERE Profilename='" & Replace(Combo2.Text, "'", "''") & "';"
                MsgBox "Profile " & Combo2.Text & " has been deleted.", vbExclamation
                
                Dim cnt As Long
                Do While cnt <= Combo2.ListCount And cnt > -1
                    If Combo2.List(cnt) = Combo2.Text Then
                        Combo2.RemoveItem cnt
                        cnt = cnt - 1
                    Else
                        cnt = cnt + 1
                    End If
                Loop
                Combo2.Text = ""
                If Combo2.ListCount > 0 Then Combo2.ListIndex = 0
            Else
                MsgBox "Profile not found!", vbExclamation
            End If
    End Select
End Sub

Private Sub NewProfile(ByVal Profilename As String)

    'Debug.Print "INSERT INTO Profiles (Profilename, Resolution, Windowed, Ambient, SoundFX, Tablet) VALUES ('" & Replace(Profilename, "'", "''") & "', '1024x768', Yes, Yes, Yes, No);"
    db.dbQuery "INSERT INTO Profiles (Profilename, Resolution, Windowed, Ambient, SoundFX, Tablet) VALUES ('" & Replace(Profilename, "'", "''") & "', '1024x768', Yes, Yes, Yes, No);"
    Dim cnt As Long
    For cnt = 0 To Combo1.ListCount - 1
        If Combo1.List(cnt) = Resolution Then
            Combo1.ListIndex = cnt
        End If
    Next

End Sub

Private Sub Form_Load()

    EnumDispModes
    
    EnumProfiles
        
    SetDefaults
    
    'Debug.Print "SELECT * FROM Users WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"
    db.rsQuery rs, "SELECT * FROM Users WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"
    If db.rsEnd(rs) Then
        'Debug.Print " " & "INSERT INTO Users (Username, Profilename) VALUES ('" & Replace(GetUserLoginName, "'", "''") & "','" & Replace(GetUserLoginName, "'", "''") & "');"
        db.dbQuery "INSERT INTO Users (Username, Profilename) VALUES ('" & Replace(GetUserLoginName, "'", "''") & "','" & Replace(GetUserLoginName, "'", "''") & "');"
        
        'Debug.Print "SELECT * FROM Users WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"
        db.rsQuery rs, "SELECT * FROM Users WHERE Username = '" & Replace(GetUserLoginName, "'", "''") & "';"
        Combo2.Text = Replace(GetUserLoginName, "'", "''")
        Combo2.AddItem Combo2.Text
    Else
        Combo2.Text = rs("Profilename")
    End If

    db.rsClose rs
    
    If Not RefreshRecord Then
        Combo2.Text = Replace(GetUserLoginName, "'", "''")
        NewProfile Replace(Combo2.Text, "'", "''")
        
        SetProfile

    End If

    If (Combo1.Text = "") And (Combo1.ListCount > 0) Then Combo1.ListIndex = 0
    
    db.rsClose rs
End Sub

Private Sub SetProfile()
    'Debug.Print "UPDATE Users SET Profilename='" & Replace(Combo2.Text, "'", "''") & "' WHERE Username='" & Replace(GetUserLoginName, "'", "''") & "';"
    db.dbQuery "UPDATE Users SET Profilename='" & Replace(Combo2.Text, "'", "''") & "' WHERE Username='" & Replace(GetUserLoginName, "'", "''") & "';"
End Sub

Private Function RefreshRecord() As Boolean
    'Debug.Print "SELECT * FROM Profiles WHERE Profilename = '" & Replace(Combo2.Text, "'", "''") & "';"
    db.rsQuery rs, "SELECT * FROM Profiles WHERE Profilename = '" & Replace(Combo2.Text, "'", "''") & "';"
    If Not db.rsEnd(rs) Then
        Dim cnt As Long
        For cnt = 0 To Combo1.ListCount - 1
            If Combo1.List(cnt) = rs("Resolution") Then
                Combo1.ListIndex = cnt
            End If
        Next
        Check1.Value = -CInt(Not CBool(rs("Windowed")))
        Check2.Value = -CInt(Not CBool(rs("Ambient")))
        Check3.Value = -CInt(Not CBool(rs("SoundFX")))
        Check4.Value = -CInt(CBool(rs("Tablet")))

        RefreshRecord = True
        
        SetProfile
        
    End If
    
End Function

Private Sub SetDefaults()
    Resolution = "1024x768"
    FullScreen = (Check1.Value = 1)
    AmbientFX = (Check2.Value = 1)
    SoundFX = (Check3.Value = 1)
    Tablet = (Check4.Value = 1)
End Sub
