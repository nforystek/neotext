VERSION 5.00
Begin VB.Form frmSwap 
   AutoRedraw      =   -1  'True
   Caption         =   "Visual Swap"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11625
   Icon            =   "frmSwap.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   11625
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2355
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   4560
      Width           =   2835
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   90
      TabIndex        =   6
      Top             =   4560
      Width           =   2025
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      Height          =   435
      Index           =   1
      Left            =   2700
      TabIndex        =   5
      Top             =   3060
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   435
      Index           =   0
      Left            =   2100
      TabIndex        =   4
      Top             =   3060
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Each next swap uses new locations"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6300
      Top             =   2880
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   4395
      Left            =   3300
      ScaleHeight     =   4335
      ScaleWidth      =   8175
      TabIndex        =   2
      Top             =   60
      Width           =   8235
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Swap"
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   3060
      Width           =   1875
   End
   Begin VB.TextBox Text1 
      Height          =   2355
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmSwap.frx":0442
      Top             =   660
      Width           =   3135
   End
End
Attribute VB_Name = "frmSwap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private elements As New Collection

Private swaps As New Collection

Private current As Long

Public Function RemoveNextArg(ByRef TheParams As Variant, ByVal TheSeperator As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare, Optional ByVal TrimResult As Boolean = True) As String
    If TrimResult Then
        If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
            RemoveNextArg = Trim(Left(TheParams, InStr(1, TheParams, TheSeperator, Compare) - 1))
            TheParams = Trim(Mid(TheParams, InStr(1, TheParams, TheSeperator, Compare) + Len(TheSeperator)))
        Else
            RemoveNextArg = Trim(TheParams)
            TheParams = ""
        End If
    Else
        If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
            RemoveNextArg = Left(TheParams, InStr(1, TheParams, TheSeperator, Compare) - 1)
            TheParams = Mid(TheParams, InStr(1, TheParams, TheSeperator, Compare) + Len(TheSeperator))
        Else
            RemoveNextArg = TheParams
            TheParams = ""
        End If
    End If
End Function

Public Function KeyExists(ByRef Col As Collection, ByVal Key As String) As Boolean
    If Col.Count > 0 Then
        On Error GoTo notexists
        Dim tmp
        If IsObject(Col.Item(1)) Then
            Set tmp = Col.Item(Key)
        Else
            tmp = Col.Item(Key)
        End If
        KeyExists = True
        Exit Function
notexists:
        KeyExists = False
        Err.Clear
    
    End If
End Function
Private Function AddElement(ByVal itm As Variant, Optional ByVal inClr As Variant) As Boolean
    Dim at As Long
    Dim ele As Element
    If (Not KeyExists(elements, itm)) And (itm <> "") Then
        at = 1
        Set ele = New Element
        ele.Holder = itm
        ele.Display = itm
        If Not IsMissing(inClr) Then SetColor ele, inClr
        If elements.Count > 0 Then
            Do While Asc(elements.Item(at).Holder) < Asc(itm) And at < elements.Count
                at = at + 1
            Loop
            If Asc(elements.Item(at).Holder) > Asc(itm) Then
                elements.Add ele, itm, at
            Else
                elements.Add ele, itm
            End If
        Else
            elements.Add ele, itm
        End If
        Set ele = Nothing
        AddElement = True
    End If
End Function
Private Sub SetColor(ByRef ele As Element, ByVal inClr As Variant)
    Select Case LCase(Trim(inClr))
        Case "blue"
            ele.Color = &HFF0000
        Case "white"
            ele.Color = vbWhite
        Case "green"
            ele.Color = &H8000&
        Case "red"
            ele.Color = &HC0&
        Case "yellow"
            ele.Color = &HC0C0&
        Case "magenta"
            ele.Color = &H800080
        Case "black"
            ele.Color = vbBlack
        Case Else
            ele.Color = LCase(Trim(inClr))
    End Select
End Sub
Private Sub ReadySwaps()
    If current < 0 Then
        
        Dim inLine As String
        Dim inText As String
        Dim inClr As Variant
        Dim inCmd As String
        Dim itm As Variant
        Dim obj As Object
        Dim at As Long
        Dim swp As Swap
        Dim ele As Element
        
        Do Until swaps.Count = 0
            swaps.Remove 1
        Loop
        Do Until elements.Count = 0
            elements.Remove 1
        Loop
        
        inText = Text1.Text
        Do Until inText = ""
            inLine = RemoveNextArg(inText, vbCrLf)
            inCmd = LCase(RemoveNextArg(inLine, " "))
            Select Case inCmd
                Case "swap"
                    Set swp = New Swap
                    
                    Do While inLine <> ""
                        swp.RightToLeft.Add UCase(Replace(RemoveNextArg(inLine, ","), ".", ""))
                    Loop
                    If swaps.Count > 0 Then
                        swaps.Add swp, , , swaps.Count
                    Else
                        swaps.Add swp
                    End If
                    Set swp = Nothing
                Case "pause", "break", "stop"
                    Set swp = New Swap
                    
                    swp.Action = IIf(inCmd = "pause", 1, IIf(inCmd = "stop", 3, 2))
                    If swaps.Count > 0 Then
                        swaps.Add swp, , , swaps.Count
                    Else
                        swaps.Add swp
                    End If
                    Set swp = Nothing
            End Select
            
        Loop
        
        For Each obj In swaps
            For Each itm In obj.RightToLeft
                AddElement itm
            Next
        Next
        
        inText = Text1.Text
        Do Until inText = ""
            inLine = RemoveNextArg(inText, vbCrLf)
            inCmd = UCase(RemoveNextArg(inLine, " "))
            For Each ele In elements
                If inCmd = ele.Holder Then
                    Select Case LCase(Trim(inLine))
                        Case "blue"
                            ele.Color = &HFF0000
                        Case "white"
                            ele.Color = vbWhite
                        Case "green"
                            ele.Color = &H8000&
                        Case "red"
                            ele.Color = &HC0&
                        Case "yellow"
                            ele.Color = &HC0C0&
                        Case "magenta"
                            ele.Color = &H800080
                        Case "black"
                            ele.Color = vbBlack
                        Case Else
                            ele.Color = LCase(Trim(inLine))
                    End Select
                End If
            Next

        Loop

        If elements.Count > 0 Then
            For at = 1 To elements.Count
                Set ele = elements.Item(at)
                
                ele.Origin.X = (Screen.TwipsPerPixelX * 15) + (((Screen.TwipsPerPixelX * 15) * 4) * at)
                ele.Origin.Y = (Picture1.ScaleHeight / 2)
                ele.Located.X = ele.Origin.X
                ele.Located.Y = ele.Origin.Y
                Set ele = Nothing
                
            Next
        End If
    End If
    current = 0
End Sub
Private Sub StopState()
    Command1.Enabled = True
    Command2(0).Enabled = True
    Command2(1).Enabled = True
    Timer1.Enabled = False
    Text1.Enabled = True
    Command1.Caption = "Swap"
End Sub

Private Sub Command1_Click()
    
    If Command1.Caption = "Swap" Then
    
        current = -1
        Text1.Enabled = False
        Command1.Caption = "Stop"
        Command2(0).Enabled = False
        Command2(1).Enabled = False
        
        ReadySwaps

        If swaps.Count > 0 And elements.Count > 0 Then
            Timer1.Enabled = True
        Else
            StopState
            current = -1
        End If
    ElseIf Command1.Caption = "Continue" Then
        Timer1.Enabled = True
        Command1.Caption = "Stop"
    Else
        StopState
        current = -1
    End If
End Sub
Private Function GetUseSwap(ByVal itm As Variant) As String
    If Check1.Value = 1 Then
        GetUseSwap = elements.Item(itm).Holder
    Else
        GetUseSwap = elements.Item(itm).Display
    End If
End Function
Private Sub PopulateCoords(ByVal fromitm As String, ByVal toitm As String)
    Dim fobj As Element
    Dim tobj As Element
    Dim itm As Variant
    
    For itm = 1 To elements.Count
        Select Case GetUseSwap(itm)
            Case fromitm
                Set fobj = elements.Item(itm)
                If Not tobj Is Nothing Then Exit For
            Case toitm
                Set tobj = elements.Item(itm)
                If Not fobj Is Nothing Then Exit For
        End Select
    Next
        
    If Not ((tobj Is Nothing) Or (fobj Is Nothing)) Then
        
        If fobj.Located.X < tobj.Located.X Then
            Bezier fobj.Coordinates, Vec(tobj.Located.X, tobj.Located.Y), _
                Vec(tobj.Located.X + ((fobj.Located.X - tobj.Located.X) / 4), (tobj.Located.Y - (Picture1.ScaleHeight / 4))), _
                Vec(tobj.Located.X + (((fobj.Located.X - tobj.Located.X) / 4) * 3), (tobj.Located.Y - (Picture1.ScaleHeight / 4))), _
                Vec(fobj.Located.X, fobj.Located.Y), 4, False
        Else
            Bezier fobj.Coordinates, Vec(fobj.Located.X, fobj.Located.Y), _
                Vec(fobj.Located.X + ((tobj.Located.X - fobj.Located.X) / 4), (fobj.Located.Y + (Picture1.ScaleHeight / 4))), _
                Vec(fobj.Located.X + (((tobj.Located.X - fobj.Located.X) / 4) * 3), (fobj.Located.Y + (Picture1.ScaleHeight / 4))), _
                Vec(tobj.Located.X, tobj.Located.Y), 4, True
        End If
    
    End If
    Set fobj = Nothing
    Set tobj = Nothing
End Sub
Private Function RenderMovement() As Boolean
    Dim ele As Object
    Dim pt As Vector
    
    For Each ele In elements
        If ele.Coordinates.Count > 0 Then
            RenderMovement = True
            Set pt = ele.Coordinates.Item(1)
            ele.Coordinates.Remove 1
            Set ele.Located = Nothing
            Set ele.Located = pt
        Else
            Set pt = ele.Located
        End If

        Picture1.Circle (pt.X, pt.Y), Screen.TwipsPerPixelX * 15, ele.Color
        Picture1.CurrentX = pt.X - (Picture1.TextWidth(ele.Holder) / 2)
        Picture1.CurrentY = pt.Y - (Picture1.TextHeight(ele.Holder) / 2)
        Picture1.ForeColor = ele.Color
        Picture1.Print ele.Holder
    
    Next

End Function

Private Sub RenderCurrentText()

    If swaps.Count > 0 Then

        Dim inc As Long
        
        Dim at As Long
        at = current - 4
        Do While at < 1
            at = at + swaps.Count
        Loop
        
        inc = 0
        Do Until inc = 8
        
            Select Case inc
                Case 0, 6
                    Picture1.ForeColor = &HC0C0C0
                    Picture1.Font.underline = False
                Case 1, 5
                    Picture1.ForeColor = &H808080
                    Picture1.Font.underline = False
                Case 2, 4
                    Picture1.ForeColor = &H404040
                    Picture1.Font.underline = False
                Case 3
                    Picture1.ForeColor = &H0&
                    Picture1.Font.underline = True
            End Select
            
            Picture1.CurrentX = (Picture1.ScaleWidth / 2) - (Picture1.TextWidth(swaps.Item(at).ToString) / 2) - (Picture1.TextWidth("W") * (3 - inc))
            Picture1.CurrentY = Picture1.ScaleHeight - (Picture1.TextHeight("W") * (7 - inc))
            
            Picture1.Print swaps.Item(at).ToString
            
            inc = inc + 1
            at = at + 1
            Do While at > swaps.Count
                at = at - swaps.Count
            Loop
        Loop

    End If
End Sub

Public Sub Swap(ByRef Var1, ByRef Var2, Optional ByRef Var3, Optional ByRef Var4, Optional ByRef Var5, Optional ByRef Var6)
    Dim Var0
    If (IsObject(Var1) Or TypeName(Var1) = "Nothing") Or _
        (IsObject(Var2) Or TypeName(Var2) = "Nothing") Then
        
        If IsMissing(Var3) Then
            Set Var0 = Var1
            Set Var1 = Var2
            Set Var2 = Var0
        ElseIf IsMissing(Var4) Then
            Set Var0 = Var1
            Set Var1 = Var2
            Set Var2 = Var3
            Set Var3 = Var0
        ElseIf IsMissing(Var5) Then
            Set Var0 = Var1
            Set Var1 = Var2
            Set Var2 = Var3
            Set Var3 = Var4
            Set Var4 = Var0
        ElseIf IsMissing(Var6) Then
            Set Var0 = Var1
            Set Var1 = Var2
            Set Var2 = Var3
            Set Var3 = Var4
            Set Var4 = Var5
            Set Var5 = Var0
        Else
            Set Var0 = Var1
            Set Var1 = Var2
            Set Var2 = Var3
            Set Var3 = Var4
            Set Var4 = Var5
            Set Var5 = Var6
            Set Var6 = Var0
        End If
    
    Else
        
        If IsMissing(Var3) Then
            Var0 = Var1
            Var1 = Var2
            Var2 = Var0
        ElseIf IsMissing(Var4) Then
            Var0 = Var1
            Var1 = Var2
            Var2 = Var3
            Var3 = Var0
        ElseIf IsMissing(Var5) Then
            Var0 = Var1
            Var1 = Var2
            Var2 = Var3
            Var3 = Var4
            Var4 = Var0
        ElseIf IsMissing(Var6) Then
            Var0 = Var1
            Var1 = Var2
            Var2 = Var3
            Var3 = Var4
            Var4 = Var5
            Var5 = Var0
        Else
            Var0 = Var1
            Var1 = Var2
            Var2 = Var3
            Var3 = Var4
            Var4 = Var5
            Var5 = Var6
            Var6 = Var0
        End If
    End If
End Sub

Private Sub Command2_Click(Index As Integer)
    
    Text1.Enabled = False
    Command1.Enabled = False
    Command2(0).Enabled = False
    Command2(1).Enabled = False

    If current < 1 Then
        ReadySwaps
    End If

    If Index = 1 Then

        current = current + 1
        If current > swaps.Count Then current = 1
        SetupCurrentSwap

    Else

        current = current - 1
        If current < 1 Then current = swaps.Count
        SetupCurrentSwap

    End If

    Timer1.Enabled = True

End Sub

Private Sub Form_Load()
    current = -1
    List1.AddItem "Swap"
    List1.AddItem "Stop"
    List1.AddItem "Break"
    List1.AddItem "Pause"
    List1.AddItem "Blue"
    List1.AddItem "White"
    List1.AddItem "Green"
    List1.AddItem "Red"
    List1.AddItem "Yellow"
    List1.AddItem "Magenta"
    List1.AddItem "Black"
    List1.ListIndex = 0
    List1_Click
    
    Text1.Text = ReadFile(App.path & "\VisualSwap.txt")
End Sub

Public Function ReadFile(ByVal path As String) As String
    Dim num As Long
    Dim Text As String
    num = FreeFile
    Open path For Binary Shared As #num Len = 1 'Access Read Lock Write As num Len = 1 'LenB(Chr(CByte(0)))
        Text = String(LOF(num), " ")
        Get #num, 1, Text
    Close #num
    ReadFile = Text
End Function

Public Sub WriteFile(ByVal path As String, ByRef Text As String)
    If (GetAttr(path) And vbReadOnly) = 0 Then
        Dim num As Integer
        num = FreeFile
        Open path For Output Shared As #num Len = 1  'Len = LenB(Chr(CByte(0)))
        Close #num
        Open path For Binary Shared As #num Len = 1 'Access Write Lock Read As #num Len = 1  'Len = LenB(Chr(CByte(0)))
            Put #num, 1, Text
        Close #num
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Text1.Left = (Screen.TwipsPerPixelX * 5)
    Text1.Top = (Screen.TwipsPerPixelY * 5) + IIf(Check1.Visible, Check1.Top + Check1.Height, 0)
    Command1.Top = Me.ScaleHeight - ((Screen.TwipsPerPixelY * 5) + Command1.Height) - (Screen.TwipsPerPixelY * 10) - List1.Height
    Command2(0).Top = Command1.Top
    Command2(1).Top = Command1.Top
    Command2(1).Left = Text1.Left + Text1.Width - ((Command2(1).Width) + (Screen.TwipsPerPixelY * 5))
    Command2(0).Left = Command2(1).Left - ((Screen.TwipsPerPixelY * 5) + Command2(1).Width)
    Command1.Width = Command2(0).Left - (Command1.Left + (Screen.TwipsPerPixelY * 5))
    Command2(0).Height = Command1.Height
    Command2(1).Height = Command1.Height
    Text1.Height = Me.ScaleHeight - (((Screen.TwipsPerPixelY * 5) * 2) + Command1.Height) - Text1.Top - (Screen.TwipsPerPixelY * 10) - List1.Height
    Command1.Left = (Screen.TwipsPerPixelX * 5)
    Picture1.Top = (Screen.TwipsPerPixelY * 5)
    Picture1.Left = ((Screen.TwipsPerPixelY * 5) * 2) + Text1.Width
    Picture1.Width = Me.ScaleWidth - Picture1.Left - (Screen.TwipsPerPixelY * 5)
    Picture1.Height = Me.ScaleHeight - ((Screen.TwipsPerPixelY * 5) * 2) - (Screen.TwipsPerPixelY * 10) - List1.Height
    
    List1.Left = (Screen.TwipsPerPixelX * 5)
    List1.Top = Command.Top + Command1.Height + (Screen.TwipsPerPixelY * 5)
    Text2.Top = List1.Top
    Text2.Left = List1.Left + List1.Width + (Screen.TwipsPerPixelX * 5)
    Text2.Width = Me.ScaleWidth - Text2.Left - (Screen.TwipsPerPixelX * 5)
    Text2.Height = List1.Height
    
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteFile App.path & "\VisualSwap.txt", Text1.Text
End Sub

Private Sub List1_Click()
    If List1.ListIndex > -1 Then
        Select Case List1.List(List1.ListIndex)
            Case "Swap"
                Text2.Text = "Usage: SWAP <?>, <?>[, <?>][, <?>][, <?>][, <?>] Notes: Up to six arguments seperated by comma. Example: SWAP A, B"
            Case "Stop"
                Text2.Text = "Usage: STOP Notes: Stops the swapping animation, and maybe used anywhere between swaps."
            Case "Break"
                Text2.Text = "Usage: BREAK Notes: Stops the swapping animation but holds the location among SWAPS and allows you to continue."
            Case "Pause"
                Text2.Text = "Usage: PAUASE Notes: Temporarily pauses the swapping animation between the swaps that it is placed, and then continues automatically."
            Case "Blue"
                Text2.Text = "Usage: BLUE <?> Notes: Colors the argument with the color blue in the swap animation, and must come after all swaps. Example: BLUE A"
            Case "White"
                Text2.Text = "Usage: WHITE <?> Notes: Colors the argument with the color white in the swap animation, and must come after all swaps. Example: WHITE A"
            Case "Green"
                Text2.Text = "Usage: GREEN <?> Notes: Colors the argument with the color green in the swap animation, and must come after all swaps. Example: GREEN A"
            Case "Red"
                Text2.Text = "Usage: RED <?> Notes: Colors the argument with the color red in the swap animation, and must come after all swaps. Example: RED A"
            Case "Yellow"
                Text2.Text = "Usage: YELLOW <?> Notes: Colors the argument with the color yellow in the swap animation, and must come after all swaps. Example: YELLOW A"
            Case "Magenta"
                Text2.Text = "Usage: MAGENTA <?> Notes: Colors the argument with the color magenta in the swap animation, and must come after all swaps. Example: MAGENTA A"
            Case "Black"
                Text2.Text = "Usage: BLACK <?> Notes: Colors the argument with the color black in the swap animation, and must come after all swaps. Example: BLACK A"
        End Select
    End If
End Sub

Private Sub Text1_Change()
    current = -1
End Sub
Private Sub SetupCurrentSwap()
    Dim swp As Swap
    Dim itm As Variant
    Dim ele As Element
    Set swp = swaps.Item(current)

    For Each ele In elements

        Do Until ele.Coordinates.Count = 0

            ele.Coordinates.Remove 1

        Loop
    
        For itm = 1 To elements.Count

            If ele.Located.X = elements.Item(itm).Origin.X And ele.Located.Y = elements.Item(itm).Origin.Y Then

                ele.Display = elements.Item(itm).Holder

            End If
        Next
    Next

    If swp.RightToLeft.Count > 0 Then
        For itm = 1 To swp.RightToLeft.Count
            If itm = 1 Then
                PopulateCoords swp.RightToLeft(1), swp.RightToLeft(swp.RightToLeft.Count)
            Else
                PopulateCoords swp.RightToLeft(itm), swp.RightToLeft(itm - 1)
            End If
        Next

    ElseIf swp.Action = 2 Then
        Timer1.Enabled = False
        Command1.Caption = "Continue"
        
    End If
    
    If swp.Action > 0 Then
        Picture1.Cls
        
        RenderMovement
        current = current + 1
        If current > swaps.Count Then current = 1
        RenderCurrentText
        
        current = current - 1
        If current <= 0 Then current = 1
        DoEvents
    End If
    
    If swp.Action = 1 Then
    
        Sleep 2000
    ElseIf swp.Action = 3 Then
        StopState
        current = -1
    End If
    
End Sub

Private Sub Timer1_Timer()

    If current <= 0 Then current = 1

    If current <= swaps.Count And current > 0 Then

        Picture1.Cls
        
        RenderCurrentText
        
        If (Not RenderMovement) Then
        
            SetupCurrentSwap

            If (Not Command1.Caption = "Swap") Then
                current = current + 1
                If current > swaps.Count Then current = 1
            End If

            If Command1.Caption = "Swap" Then
                Timer1.Enabled = False
                StopState

            Else
                Timer1.Interval = 100

            End If
            
        Else
            Timer1.Interval = 1
        End If

    End If
        
End Sub
