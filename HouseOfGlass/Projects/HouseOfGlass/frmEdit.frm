VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "House Of Glass Level Editor"
   ClientHeight    =   10410
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   15915
   FillStyle       =   0  'Solid
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10410
   ScaleWidth      =   15915
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   -45
      Width           =   9420
      Begin VB.OptionButton Option1 
         Caption         =   "Starting Block"
         Height          =   315
         Index           =   4
         Left            =   5865
         TabIndex        =   10
         Top             =   135
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ending Block"
         Height          =   315
         Index           =   5
         Left            =   7275
         TabIndex        =   9
         Top             =   135
         Width           =   1365
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Blue Glass Wall"
         Height          =   315
         Index           =   3
         Left            =   4335
         TabIndex        =   8
         Top             =   135
         Width           =   1425
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Solid White Wall"
         Height          =   315
         Index           =   2
         Left            =   2760
         TabIndex        =   7
         Top             =   135
         Width           =   1590
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Hollow Wall"
         Height          =   225
         Index           =   1
         Left            =   1530
         TabIndex        =   6
         Top             =   180
         Width           =   1320
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Tint Glass Wall"
         Height          =   225
         Index           =   0
         Left            =   60
         TabIndex        =   5
         Top             =   180
         Value           =   -1  'True
         Width           =   1470
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
      Height          =   9765
      Left            =   120
      MousePointer    =   2  'Cross
      ScaleHeight     =   647
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1040
      TabIndex        =   0
      Top             =   495
      Width           =   15660
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
Private Const Title = "House Of Glass Level Editor"

Private Const GridWidth = 750
Private Const GridHeight = 450

Private Const GridX = 20
Private Const GridY = 20

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

Private Function GetWallType(Optional ByVal Index As Long = -1) As Long
    If Index > -1 Then
        Select Case Index
            Case 0
                GetWallType = RGB(192, 0, 0)
            Case 1
                GetWallType = RGB(0, 192, 0)
            Case 2
                GetWallType = RGB(192, 192, 192)
            Case 3
                GetWallType = RGB(128, 255, 255)
            Case 4
                GetWallType = RGB(192, 192, 0)
            Case 5
                GetWallType = RGB(0, 192, 192)
        End Select
    Else
        If Option1(0).Value Then
            GetWallType = 0
        ElseIf Option1(1).Value Then
            GetWallType = 1
        ElseIf Option1(2).Value Then
            GetWallType = 2
        ElseIf Option1(3).Value Then
            GetWallType = 3
        ElseIf Option1(4).Value Then
            GetWallType = 4
        ElseIf Option1(5).Value Then
            GetWallType = 5
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
    FileName = "Untitled.hog"
    DrawDesign
End Sub
Public Sub OpenDesign()
    On Error Resume Next
    
    With CommonDialog1
        .CancelError = True
        .DialogTitle = "Open Design"
        .Filter = "House Of Glass Level|*.hog|All Files|*.*"
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
            .Filter = "House Of Glass Level|*.hog|All Files|*.*"
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
        Picture1.DrawWidth = 2
        If dLine.Selected And ActionMode = aModeDelete Then
        
            If dLine.WallType <= 3 Then
                Picture1.Line (dLine.X1, dLine.Y1)-(dLine.X2, dLine.Y2), RGB(0, 0, 255)
            Else
                Picture1.Line (dLine.X1, dLine.Y1)-(dLine.X2, dLine.Y2), RGB(0, 0, 255), BF
            End If
        
        Else
            If dLine.WallType <= 3 Then
                Picture1.Line (dLine.X1, dLine.Y1)-(dLine.X2, dLine.Y2), GetWallType(dLine.WallType)
            Else
                Picture1.Line (dLine.X1, dLine.Y1)-(dLine.X2, dLine.Y2), GetWallType(dLine.WallType), BF
            End If
        End If
    Next
    
    If IsDrawing Then

        Dim dx As Long, dy As Long
        dx = MouseX: dy = MouseY
        FindNearestGridDot dx, dy
        
        If GetWallType <= 3 Then
            Picture1.Line (LineStartX, LineStartY)-(dx, dy), GetWallType
        Else
            Picture1.Line (LineStartX, LineStartY)-(dx, LineStartY), GetWallType
            Picture1.Line (dx, LineStartY)-(dx, dy), GetWallType
            Picture1.Line (dx, dy)-(LineStartX, dy), GetWallType
            Picture1.Line (LineStartX, dy)-(LineStartX, LineStartY), GetWallType
        End If
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
                If (GetWallType <= 3) Or (((Not (LineStartX = dx)) And (Not (LineStartY = dy))) And (GetWallType > 3)) Then
                    
                    If (IsDrawing > 0) Then
                        If (((Abs((LineStartX / GridX) - (dx / GridX)) = 1) And (Abs((LineStartY / GridY) - (dy / GridY)) = 0)) Or _
                             ((Abs((LineStartX / GridX) - (dx / GridX)) = 0) And (Abs((LineStartY / GridY) - (dy / GridY)) = 1))) Or _
                             Not ((Abs((LineStartX / GridX) - (dx / GridX)) = 0) And (Abs((LineStartY / GridY) - (dy / GridY)) = 0)) Then
                    
                            Dim newLine As New clsWallEdit
        
                            IsDrawing = 0
                            With newLine
                                .X1 = IIf(LineStartX > dx, dx, LineStartX)
                                .Y1 = IIf(LineStartY > dy, dy, LineStartY)
                                .X2 = IIf(LineStartX > dx, LineStartX, dx)
                                .Y2 = IIf(LineStartY > dy, LineStartY, dy)
                                .WallType = GetWallType()
                            End With
                            
                            Edits.Add newLine
                            
                            LineStartX = -1
                            LineStartY = -1
                                
                            Set newLine = Nothing
                        Else
                            MsgBox "Invalid Wall Size"
                        End If
                    Else
                        IsDrawing = 1
                        LineStartX = dx
                        LineStartY = dy
                    End If
                    
                Else
                    MsgBox "Invalid Block Size"
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
    Dim dLine As clsWallEdit
    FileNum = FreeFile
    Open FileName For Input As #FileNum
        Do Until EOF(FileNum)
            Line Input #FileNum, inLine
            Set dLine = New clsWallEdit
            With dLine
                .X1 = (RemoveNextParam(inLine, ":") * GridX)
                .Y1 = (RemoveNextParam(inLine, ":") * GridY)
                .X2 = (RemoveNextParam(inLine, ":") * GridX)
                .Y2 = (RemoveNextParam(inLine, ":") * GridY)
                .WallType = RemoveNextParam(inLine, ":")
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
            Print #FileNum, (dLine.X1 / GridX) & ":" & (dLine.Y1 / GridY) & ":" & (dLine.X2 / GridX) & ":" & (dLine.Y2 / GridY) & ":" & dLine.WallType
        Next
    Close #FileNum
End Function

Function RemoveNextParam(ByRef TheParams As String, ByVal TheSeperator As String) As String
    Dim RetVal As String
    If InStr(TheParams, TheSeperator) > 0 Then
        RetVal = Trim(Left(TheParams, InStr(TheParams, TheSeperator) - 1))
        TheParams = Trim(Mid(TheParams, InStr(TheParams, TheSeperator) + Len(TheSeperator)))
    Else
        RetVal = Trim(TheParams)
        TheParams = ""
    End If
    RemoveNextParam = RetVal
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



