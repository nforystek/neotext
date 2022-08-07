
VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blacklawn Ship Editor"
   ClientHeight    =   7575
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   11580
   FillStyle       =   0  'Solid
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11580
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   -45
      Width           =   9420
      Begin VB.OptionButton Option1 
         Caption         =   "Orange"
         Height          =   225
         Index           =   9
         Left            =   8205
         TabIndex        =   14
         Top             =   180
         Width           =   915
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Grey"
         Height          =   225
         Index           =   8
         Left            =   7425
         TabIndex        =   13
         Top             =   180
         Width           =   645
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Purple"
         Height          =   225
         Index           =   7
         Left            =   6525
         TabIndex        =   12
         Top             =   180
         Width           =   915
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Blue"
         Height          =   225
         Index           =   6
         Left            =   5790
         TabIndex        =   11
         Top             =   180
         Width           =   690
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Terqoise"
         Height          =   225
         Index           =   5
         Left            =   4770
         TabIndex        =   10
         Top             =   180
         Width           =   915
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Green"
         Height          =   225
         Index           =   4
         Left            =   3915
         TabIndex        =   9
         Top             =   180
         Width           =   915
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Yellow"
         Height          =   225
         Index           =   3
         Left            =   3045
         TabIndex        =   8
         Top             =   180
         Width           =   915
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Red"
         Height          =   225
         Index           =   2
         Left            =   2325
         TabIndex        =   7
         Top             =   180
         Width           =   690
      End
      Begin VB.OptionButton Option1 
         Caption         =   "White"
         Height          =   225
         Index           =   1
         Left            =   1530
         TabIndex        =   6
         Top             =   180
         Width           =   765
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Current Custom"
         Height          =   225
         Index           =   0
         Left            =   60
         TabIndex        =   5
         Top             =   180
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6765
      Top             =   75
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   165
      TabIndex        =   1
      Top             =   -45
      Width           =   1755
      Begin VB.OptionButton Option2 
         Caption         =   "Delete"
         Height          =   210
         Index           =   1
         Left            =   885
         TabIndex        =   3
         Top             =   195
         Width           =   825
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Draw"
         Height          =   210
         Index           =   0
         Left            =   105
         TabIndex        =   2
         Top             =   195
         Value           =   -1  'True
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   6990
      Left            =   120
      MousePointer    =   2  'Cross
      ScaleHeight     =   462
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   752
      TabIndex        =   0
      Top             =   495
      Width           =   11340
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN
Option Compare Text
Private Const Title = "Blacklawn Ship Editor"

Private Const GridWidth = 750
Private Const GridHeight = 450

Private Const GridX = 10
Private Const GridY = 10

Private Const aModeDraw = 0
Private Const aModeDelete = 1
Private ShiftToggle As Boolean
Private ActionMode As Integer

Private MouseX As Long
Private MouseY As Long

Private IsDrawing As Long
Private LineStartX As Long
Private LineStartY As Long

Private pFileName As String
Private pChanged As Boolean

Private Edits As New Collection

Private Function GetColor(Optional ByVal Index As Long = -1) As Long
    If Index > -1 Then
        Select Case Index
            Case 0
                GetColor = RGB(0, 0, 0)
            Case 1
                GetColor = RGB(255, 255, 255)
            Case 2
                GetColor = RGB(255, 0, 0)
            Case 3
                GetColor = RGB(255, 255, 0)
            Case 4
                GetColor = RGB(0, 255, 0)
            Case 5
                GetColor = RGB(0, 255, 255)
            Case 6
                GetColor = RGB(0, 0, 255)
            Case 7
                GetColor = RGB(255, 0, 255)
            Case 8
                GetColor = RGB(128, 128, 128)
            Case 9
                GetColor = RGB(255, 128, 0)
        End Select
    Else
        If Option1(0).Value Then
            GetColor = 0
        ElseIf Option1(1).Value Then
            GetColor = 1
        ElseIf Option1(2).Value Then
            GetColor = 2
        ElseIf Option1(3).Value Then
            GetColor = 3
        ElseIf Option1(4).Value Then
            GetColor = 4
        ElseIf Option1(5).Value Then
            GetColor = 5
        ElseIf Option1(6).Value Then
            GetColor = 6
        ElseIf Option1(7).Value Then
            GetColor = 7
        ElseIf Option1(8).Value Then
            GetColor = 8
        ElseIf Option1(9).Value Then
            GetColor = 9
        End If
    End If
End Function

Private Property Get FileName() As String
    FileName = pFileName
End Property
Private Property Let FileName(ByVal NewVal As String)
    pFileName = NewVal
    SetCaption
End Property
Private Property Get Changed() As Boolean
    Changed = pChanged
End Property
Private Property Let Changed(ByVal NewVal As Boolean)
    pChanged = NewVal
    SetCaption
End Property
Private Sub SetCaption()
    If pChanged Then
        Me.Caption = Title & " - [" & pFileName & "] *"
    Else
        Me.Caption = Title & " - [" & pFileName & "]"
    End If
End Sub

Public Sub ClearObjects()
    Do Until Edits.Count = 0
        Edits.Remove 1
    Loop

End Sub

Public Sub NewDesign()
    ClearObjects
    Changed = False
    FileName = "Untitled.sx"
    DrawDesign
End Sub
Public Sub OpenDesign()
    On Error Resume Next
    
    With CommonDialog1
        .CancelError = True
        .DialogTitle = "Open Design"
        .Filter = "Blacklawn Ships|*.sx|All Files|*.*"
        .FilterIndex = 1
        .ShowOpen
    End With
    If Err.Number = 0 Then
        'loaddocument
        ClearObjects
        
        Changed = False
        FileName = CommonDialog1.FileName
    
        LoadFile CommonDialog1.FileName
        DrawDesign
    
    Else
        Err.Clear
    End If

End Sub
Public Sub SaveDesign(ByVal ShowDialog As Boolean)
    If ShowDialog Or Not FileExists(FileName) Then
        On Error Resume Next
        With CommonDialog1
            .CancelError = True
            .DialogTitle = "Save Design As..."
            .Filter = "Blacklawn Ships|*.sx|All Files|*.*"
            .FilterIndex = 1
            .ShowSave
        End With
        If Err.Number = 0 Then
            FileName = CommonDialog1.FileName
        Else
            Err.Clear
            Exit Sub
        End If
    
    End If
    
    SaveFile FileName
    
    
    Changed = False
End Sub

Public Sub DrawDesign()
    Picture1.Cls
    DrawGrid
    DrawEdits
    
End Sub

Public Sub DrawEdits()

    Picture1.FillStyle = 0
    
    Dim cntX As Long
    Dim cntY As Long
    
    Dim dLine
    For Each dLine In Edits
        Picture1.DrawWidth = 1
        If dLine.Selected And ActionMode = aModeDelete Then
        
            Picture1.Line (dLine.X1, dLine.Y1)-(dLine.X2, dLine.Y2), RGB(128 + 64, 128 + 64, 128 + 64), BF
        
        Else
        
            Picture1.Line (dLine.X1, dLine.Y1)-(dLine.X2, dLine.Y2), GetColor(dLine.Color), BF
            
        End If
    Next
    
    If IsDrawing Then

        Dim dx As Long, dy As Long
        dx = MouseX: dy = MouseY
        FindNearestGridDot dx, dy
        
        Picture1.Line (LineStartX, LineStartY)-(dx, LineStartY), GetColor
        Picture1.Line (dx, LineStartY)-(dx, dy), GetColor
        Picture1.Line (dx, dy)-(LineStartX, dy), GetColor
        Picture1.Line (LineStartX, dy)-(LineStartX, LineStartY), GetColor
        
    End If
End Sub
Public Sub DrawGrid()

    Picture1.FillStyle = 0

    Dim X As Long
    Dim Y As Long
    
    Picture1.DrawWidth = 1
    For X = 0 To Picture1.ScaleWidth Step GridX
        For Y = 0 To Picture1.ScaleHeight Step GridY
            Picture1.PSet (X, Y), RGB(0, 0, 0)
        Next
    Next
        
End Sub


Private Sub Form_Load()
    NewDesign
    
    ActionMode = aModeDraw
    
    IsDrawing = 0
    
    LineStartX = -1
    LineStartY = -1
        
    DrawDesign
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        If Changed Then
            Cancel = (MsgBox("The current design has changed, do you want to proceed?", vbYesNo) = vbNo)
        End If
    End If
End Sub

Private Sub mnuNew_Click()
    Dim Continue As Boolean
    Continue = True
    
    If Changed Then
        Continue = (MsgBox("The current design has changed, do you want to proceed?", vbYesNo) = vbYes)
    End If
    
    If Continue Then
        NewDesign
    End If
End Sub

Private Sub mnuOpen_Click()
    Dim Continue As Boolean
    Continue = True
    
    If Changed Then
        Continue = (MsgBox("The current design has changed, do you want to proceed?", vbYesNo) = vbYes)
    End If
    
    If Continue Then
        OpenDesign
    End If
End Sub

Private Sub mnuSave_Click()
    SaveDesign False
End Sub

Private Sub mnuSaveAs_Click()
    SaveDesign True
End Sub

Private Sub Option1_Click(Index As Integer)
    
    Picture1.SetFocus
End Sub

Private Sub Option2_Click(Index As Integer)
    ActionMode = Index
    Option2(ActionMode).Value = True
    
    IsDrawing = 0
    Picture1.SetFocus
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift And (Not ShiftToggle) Then
        ShiftToggle = True
        If Option2(0).Value Then
            ActionMode = aModeDelete
            Option2(1).Value = True
        Else
            ActionMode = aModeDraw
            Option2(0).Value = True
        End If
    Else
        ShiftToggle = False
    End If
End Sub

Private Sub RemoveSelected()
    Dim delIndex As Long
    Dim cnt As Long
    
    delIndex = 0
    If Edits.Count > 0 Then
        For cnt = 1 To Edits.Count
            If Edits(cnt).Selected Then
                delIndex = cnt
            End If
        Next
        If delIndex > 0 Then
            Edits.Remove delIndex
        End If
    End If
    
    DrawDesign
End Sub


Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
    ShiftToggle = False
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
    
    Dim dx As Long
    Dim dy As Long
    dx = MouseX
    dy = MouseY
    
    FindNearestGridDot dx, dy
    
    If Button = 1 Then
        Changed = True
        Select Case ActionMode
            Case aModeDraw

                If IsDrawing > 0 Then
                    If (Not (LineStartX = dx)) And (Not (LineStartY = dy)) Then
                    
                        Dim newLine As New clsShipEdit
    
                        IsDrawing = 0
                        With newLine
                            .X1 = IIf(LineStartX > dx, dx, LineStartX)
                            .Y1 = IIf(LineStartY > dy, dy, LineStartY)
                            .X2 = IIf(LineStartX > dx, LineStartX, dx)
                            .Y2 = IIf(LineStartY > dy, LineStartY, dy)
                            .Color = GetColor()
                        End With
                        
                        Edits.Add newLine
                        
                        LineStartX = -1
                        LineStartY = -1
                            
                        Set newLine = Nothing
                    Else
                        MsgBox "Invalid Square Size"
                    End If
                Else
                    IsDrawing = 1
                    LineStartX = dx
                    LineStartY = dy
                End If
        
            Case aModeDelete
                RemoveSelected
            
        End Select
    ElseIf Button = 2 Then
        Select Case ActionMode
            Case aModeDraw
                Option2_Click aModeDelete
            Case aModeDelete
                Option2_Click aModeDraw
        End Select
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
    
    Dim dx As Long
    Dim dy As Long
    dx = MouseX
    dy = MouseY
            
    Select Case ActionMode
        Case aModeDraw
            Picture1.MousePointer = 2

            FindNearestGridDot dx, dy
            
            DrawDesign
            
            Picture1.FillStyle = 1
            
            Picture1.DrawWidth = 1
            Picture1.Circle (dx, dy), 3, RGB(255, 0, 0)
        
            If IsDrawing > 0 Then
                Picture1.DrawWidth = 1
                Picture1.Circle (LineStartX, LineStartY), 3, RGB(255, 0, 0)
            End If
            
        Case aModeDelete
            Picture1.MousePointer = 1
            
            Dim shortest As Single
            Dim nextLine As Single
            Dim dLine
            Dim selLine
            Set selLine = Nothing
            
            For Each dLine In Edits
                nextLine = DistToSegment(dx, dy, dLine.X1, dLine.Y1, dLine.X2, dLine.Y2)
                If InsideSquare(dx, dy, dLine.X1, dLine.Y1, dLine.X2, dLine.Y2) Or ((nextLine < shortest) Or (shortest = 0)) Then
                    shortest = nextLine
                    Set selLine = dLine
                End If
                dLine.Selected = False
            Next
            If (Not selLine Is Nothing) Then
                selLine.Selected = True
            End If
            
            Set selLine = Nothing
            
            DrawDesign
        
    End Select
End Sub

Private Sub FindNearestGridDot(ByRef X As Long, ByRef Y As Long)
    X = (X / GridX)
    Y = (Y / GridY)
    X = X * GridX
    Y = Y * GridY
End Sub

Public Function LoadFile(ByVal FileName As String)
    Dim FileNum As Long
    Dim inLine As String
    Dim dLine As clsShipEdit
    FileNum = FreeFile
    Open FileName For Input As #FileNum
        Do Until EOF(FileNum)
            Line Input #FileNum, inLine
            Set dLine = New clsShipEdit
            With dLine
                .X1 = GridWidth - -(RemoveNextParam(inLine, ":") * GridX)
                .Y1 = (-((RemoveNextParam(inLine, ":") - 1) * GridY) + GridHeight)
                .X2 = GridWidth - -(RemoveNextParam(inLine, ":") * GridX)
                .Y2 = (-((RemoveNextParam(inLine, ":") - 1) * GridY) + GridHeight)
                .Color = RemoveNextParam(inLine, ":")
            End With
            Edits.Add dLine
            Set dLine = Nothing
        Loop
    Close #FileNum
End Function

Public Function SaveFile(ByVal FileName As String)
    Dim dLine
    Dim FileNum As Long
    FileNum = FreeFile
    Open FileName For Output As #FileNum
        For Each dLine In Edits
            Print #FileNum, (-(GridWidth - dLine.X1) / GridX) & ":" & (((-dLine.Y1 + GridHeight) / GridY) + 1) & ":" & (-(GridWidth - dLine.X2) / GridX) & ":" & (((-dLine.Y2 + GridHeight) / GridY) + 1) & ":" & dLine.Color
        Next
    Close #FileNum
End Function

Function RemoveNextParam(ByRef TheParams As String, ByVal TheSeperator As String) As String
    Dim retVal As String
    If InStr(TheParams, TheSeperator) > 0 Then
        retVal = Trim(Left(TheParams, InStr(TheParams, TheSeperator) - 1))
        TheParams = Trim(Mid(TheParams, InStr(TheParams, TheSeperator) + Len(TheSeperator)))
    Else
        retVal = Trim(TheParams)
        TheParams = ""
    End If
    RemoveNextParam = retVal
End Function

Public Function FileExists(ByVal NewVal As String) As Boolean
    On Error Resume Next
    Dim l As Long
    l = FileLen(NewVal)
    FileExists = (Err.Number = 0)
    Err.Clear
End Function
Public Function InsideSquare(ByVal px As Single, ByVal py _
    As Single, ByVal X1 As Single, ByVal Y1 As Single, _
    ByVal X2 As Single, ByVal Y2 As Single) As Boolean
    If X1 < X2 Then Swap X1, X2
    If Y1 < Y2 Then Swap Y1, Y2

    InsideSquare = (px >= X1 And px <= X2) And (py >= Y1 And py <= Y2)
End Function
Public Sub Swap(ByRef val1 As Single, ByRef val2 As Single)
    Dim tmp As Single
    tmp = val1
    val1 = val2
    val2 = tmp
End Sub
Public Function DistToSegment(ByVal px As Single, ByVal py _
    As Single, ByVal X1 As Single, ByVal Y1 As Single, _
    ByVal X2 As Single, ByVal Y2 As Single) As Single

    On Error Resume Next
    
    Dim dx As Single
    Dim dy As Single
    Dim t As Single

    dx = X2 - X1
    dy = Y2 - Y1
    If dx = 0 And dy = 0 Then
        ' It's a point not a line segment.
        dx = px - X1
        dy = py - Y1
        DistToSegment = Sqr(dx * dx + dy * dy)
        Exit Function
    End If

    t = (px + py - X1 - Y1) / (dx + dy)

    If t < 0 Then
        dx = px - X1
        dy = py - Y1
    ElseIf t > 1 Then
        dx = px - X2
        dy = py - Y2
    Else
        X2 = X1 + t * dx
        Y2 = Y1 + t * dy
        dx = px - X2
        dy = py - Y2
    End If
    DistToSegment = Sqr(dx * dx + dy * dy)
    Err.Clear
End Function


