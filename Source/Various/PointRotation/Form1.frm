VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rotation"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12390
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   12390
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   3960
      Left            =   8280
      ScaleHeight     =   260
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   260
      TabIndex        =   2
      Top             =   120
      Width           =   3960
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   3960
      Left            =   4200
      ScaleHeight     =   260
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   260
      TabIndex        =   1
      Top             =   120
      Width           =   3960
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3960
      Left            =   120
      ScaleHeight     =   260
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   260
      TabIndex        =   0
      Top             =   120
      Width           =   3960
   End
   Begin VB.Label Label1 
      Caption         =   "Use any of three buttons on a mouse clicking on any of the canvas above to move the rotaitons around."
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   4200
      Width           =   7575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Me.MousePointer = 2
    
    NewPoints

End Sub

Private Sub Form_Paint()
    
    DrawAxis "X Axis", Picture1
    DrawAxis "Y Axis", Picture2
    DrawAxis "Z Axis", Picture3
        
    DrawPointRotation Picture1, Vertex(1).Y, Vertex(1).Z, Rotate(1).X, 255, 0, 0
    DrawPointRotation Picture2, Vertex(1).Z, Vertex(1).X, Rotate(1).Y, 255, 0, 0
    DrawPointRotation Picture3, Vertex(1).X, Vertex(1).Y, Rotate(1).Z, 255, 0, 0
    
    DrawPointRotation Picture1, Vertex(2).Y, Vertex(2).Z, Rotate(2).X, 0, 255, 0
    DrawPointRotation Picture2, Vertex(2).Z, Vertex(2).X, Rotate(2).Y, 0, 255, 0
    DrawPointRotation Picture3, Vertex(2).X, Vertex(2).Y, Rotate(2).Z, 0, 255, 0

    DrawPointRotation Picture1, Vertex(3).Y, Vertex(3).Z, Rotate(3).X, 0, 0, 255
    DrawPointRotation Picture2, Vertex(3).Z, Vertex(3).X, Rotate(3).Y, 0, 0, 255
    DrawPointRotation Picture3, Vertex(3).X, Vertex(3).Y, Rotate(3).Z, 0, 0, 255
        
End Sub

Private Sub DrawAxis(ByVal Axis As String, ByRef PicBox As PictureBox)
    'draw z axis
    PicBox.Cls
    PicBox.CurrentX = 1
    PicBox.CurrentY = 1
    PicBox.Print Axis
    Select Case Axis
        Case "X Axis"
            PicBox.CurrentX = 2
            PicBox.CurrentY = (PicBox.ScaleHeight / 2) + 1
            PicBox.Print "Y"
            PicBox.CurrentX = (PicBox.ScaleWidth / 2) + 3
            PicBox.CurrentY = 1
            PicBox.Print "Z"
        Case "Y Axis"
            PicBox.CurrentX = 2
            PicBox.CurrentY = (PicBox.ScaleHeight / 2) + 1
            PicBox.Print "Z"
            PicBox.CurrentX = (PicBox.ScaleWidth / 2) + 3
            PicBox.CurrentY = 1
            PicBox.Print "X"
        Case "Z Axis"
            PicBox.CurrentX = 2
            PicBox.CurrentY = (PicBox.ScaleHeight / 2) + 1
            PicBox.Print "X"
            PicBox.CurrentX = (PicBox.ScaleWidth / 2) + 3
            PicBox.CurrentY = 1
            PicBox.Print "Y"
    End Select
    PicBox.Line (0, (PicBox.ScaleHeight / 2))-(PicBox.ScaleWidth, (PicBox.ScaleHeight / 2)), &H808080
    PicBox.Line ((PicBox.ScaleWidth / 2), 0)-((PicBox.ScaleWidth / 2), PicBox.ScaleHeight), &H808080
End Sub

Private Sub DrawPointRotation(ByRef PicBox As PictureBox, ByVal vX As Double, ByVal vY As Double, ByVal rZ As Double, ByVal Red As Long, ByVal Green As Long, ByVal Blue As Long)

    Dim Distance As Single
    Distance = (((vX ^ 2) + (vY ^ 2)) ^ (1 / 2))

    PicBox.Line ((PicBox.ScaleWidth / 2), (PicBox.ScaleHeight / 2))-( _
                ((PicBox.ScaleWidth / 2) + (Distance * Sin(rZ))), _
                ((PicBox.ScaleHeight / 2) - (Distance * Cos(rZ)))), RGB(Red, Green, Blue)

    PicBox.Line ((PicBox.ScaleWidth / 2), (PicBox.ScaleHeight / 2))-( _
                ((PicBox.ScaleWidth / 2) + (Distance * Sine(vX, vY))), _
                ((PicBox.ScaleHeight / 2) - (Distance * Cosine(vX, vY)))), RGB(IIf(Red = 0, 128, Red), IIf(Green = 0, 128, Green), IIf(Blue = 0, 128, Blue))

    PicBox.Circle (((PicBox.ScaleWidth / 2) + vX), ((PicBox.ScaleHeight / 2) + vY)), 6, RGB(Red, Green, Blue)

    PicBox.Circle (((PicBox.ScaleWidth / 2) + vX), ((PicBox.ScaleHeight / 2) + vY)), 6, RGB(IIf(Red = 0, 128, Red), IIf(Green = 0, 128, Green), IIf(Blue = 0, 128, Blue))

End Sub

Public Sub CornerPrint(ByRef PicBox As PictureBox, ByVal Txt As String)
    PicBox.CurrentY = PicBox.ScaleHeight - PicBox.TextHeight(Txt)
    PicBox.CurrentX = PicBox.ScaleWidth - PicBox.TextWidth(Txt)
    PicBox.Print Txt
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then NewPoints
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePoint Button, , X - (Picture1.ScaleHeight / 2), Y - (Picture1.ScaleWidth / 2)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePoint Button, , X - (Picture1.ScaleHeight / 2), Y - (Picture1.ScaleWidth / 2)
End Sub

Private Sub Picture2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then NewPoints
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePoint Button, Y - (Picture2.ScaleWidth / 2), , X - (Picture2.ScaleHeight / 2)
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePoint Button, Y - (Picture2.ScaleWidth / 2), , X - (Picture2.ScaleHeight / 2)
End Sub

Private Sub Picture3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then NewPoints
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePoint Button, X - (Picture3.ScaleWidth / 2), Y - (Picture3.ScaleHeight / 2)
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePoint Button, X - (Picture3.ScaleWidth / 2), Y - (Picture3.ScaleHeight / 2)
End Sub

Private Sub MousePoint(Button As Integer, Optional ByRef X As Variant, Optional ByRef Y As Variant, Optional ByRef Z As Variant)

    If Button <> 0 Then
        If Button = 4 Then Button = 3
        
        If Button = 1 Or Button = 2 Or Button = 3 Then
            
            Dim v As Point
            Dim r As Point
            Set v = Vertex(Button)
            Set r = Rotate(Button)
            If IsMissing(X) Then
                GetNewPoint v, , Y, Z
            ElseIf IsMissing(Y) Then
                GetNewPoint v, X, , Z
            ElseIf IsMissing(Z) Then
                GetNewPoint v, X, Y
            End If
            
    
            Set r = AnglesOfPoint2(v)

            If Button < Vertex.Count Then
                Vertex.Remove Button
                Rotate.Remove Button
                
                Vertex.Add v, , Button
                Rotate.Add r, , Button
            Else
                Vertex.Remove Button
                Rotate.Remove Button
                
                Vertex.Add v
                Rotate.Add r
            End If
        End If
    End If
    
    Form_Paint
End Sub
Public Function GetNewPoint(Optional ByRef v As Point, Optional ByRef X As Variant, Optional ByRef Y As Variant, Optional ByRef Z As Variant) As Point
    If IsMissing(X) And IsMissing(Y) And IsMissing(Z) Then
        Set v = RandomPoint(New Plane, -(Picture1.ScaleWidth / 2), (Picture1.ScaleWidth / 2))
    Else

        If Not IsMissing(X) Then v.X = X
        If Not IsMissing(Y) Then v.Y = Y
        If Not IsMissing(Z) Then v.Z = Z
        Dim dist As Double
        dist = DistanceEx(MakePoint(0, 0, 0), v)
        If dist > 0 Then
            Dim n As Point
            Set n = PointNormalize(v)
            Set v = DistanceSet(MakePoint(0, 0, 0), n, dist)
            Set n = Nothing
        End If

    End If
    Set GetNewPoint = v
End Function

Public Sub NewPoints()

    Do While Vertex.Count > 0
        Vertex.Remove 1
    Loop
    
    Do While Rotate.Count > 0
        Rotate.Remove 1
    Loop
    
    Vertex.Add GetNewPoint
    Vertex.Add GetNewPoint
    Vertex.Add GetNewPoint
    
    Rotate.Add AnglesOfPoint1(Vertex(1))
    Rotate.Add AnglesOfPoint1(Vertex(2))
    Rotate.Add AnglesOfPoint1(Vertex(3))
    
    Form_Paint
End Sub
