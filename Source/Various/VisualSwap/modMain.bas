Attribute VB_Name = "modMain"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)

'Bezier curve code modified from original found at:
'http://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=21402&lngWId=1&txtForceRefresh=71201913472648124

'A Vector
Public Type TVector
    X As Single
    Y As Single
End Type 'TVector


'Vector addition
Function Add(l As TVector, R As TVector) As TVector
    Dim Vec As TVector  'Attention aux Effets de bords
    Vec.X = l.X + R.X
    Vec.Y = l.Y + R.Y
    Add = Vec
End Function 'Add


'Scalar-Vector multiplication
Function Mul(S As Single, V As TVector) As TVector
    Dim Vec As TVector
    Vec.X = V.X * S
    Vec.Y = V.Y * S
    Mul = Vec
End Function 'Mul


'Return a vector
Public Function Vec(X As Single, Y As Single) As TVector
    Vec.X = X
    Vec.Y = Y
End Function 'Vec

'Draw a Bezier curve on pct with the specified control points. Depth is Recursive depth for Bezier calc.
Sub Bezier(pct As Collection, P0 As TVector, P1 As TVector, P2 As TVector, P3 As TVector, Depth As Byte, Rtl As Boolean)
    
    Dim nP0 As TVector
    Dim nP1 As TVector
    Dim nP2 As TVector
    Dim nP3 As TVector
    
    'Depth :
    If Depth > 0 Then

        'left
        nP0 = P0
        nP1 = Add(Mul(1 / 2, P0), Mul(1 / 2, P1))
        nP2 = Add(Add(Mul(1 / 4, P0), Mul(1 / 2, P1)), Mul(1 / 4, P2))
        nP3 = Add(Add(Add(Mul(1 / 8, P0), Mul(3 / 8, P1)), Mul(3 / 8, P2)), Mul(1 / 8, P3))
        Bezier pct, nP0, nP1, nP2, nP3, Depth - 1, Rtl


        'right
        nP0 = P3
        nP1 = Add(Mul(1 / 2, P3), Mul(1 / 2, P2))
        nP2 = Add(Add(Mul(1 / 4, P3), Mul(1 / 2, P2)), Mul(1 / 4, P1))
        nP3 = Add(Add(Add(Mul(1 / 8, P3), Mul(3 / 8, P2)), Mul(3 / 8, P1)), Mul(1 / 8, P0))
        Bezier pct, nP0, nP1, nP2, nP3, Depth - 1, Rtl
    Else

        'pct.Line (P0.X, P0.Y)-(P1.X, P1.Y)
        'pct.Line -(P2.X, P2.Y)
        'pct.Line -(P3.X, P3.Y)
        
        AddCoord pct, P0, Rtl
        AddCoord pct, P1, Rtl
        AddCoord pct, P2, Rtl
        AddCoord pct, P3, Rtl
    
    End If
End Sub 'Draw

Private Sub AddCoord(pct As Collection, P0 As TVector, ByVal Rtl As Boolean)
    Dim p4 As Vector
    
    Set p4 = New Vector
    p4.X = P0.X
    p4.Y = P0.Y
    If pct.Count > 0 Then
        Dim cnt As Long
        cnt = 1
        Do While (((pct.Item(cnt).X < P0.X) And (Not Rtl)) Xor ((pct.Item(cnt).X > P0.X) And Rtl)) And (cnt < pct.Count)
            cnt = cnt + 1
        Loop
        
        If (((pct.Item(cnt).X < P0.X) And (Not Rtl)) Xor ((pct.Item(cnt).X > P0.X) And Rtl)) Then
            pct.Add p4
        Else
            pct.Add p4, , cnt
        End If
    Else
        pct.Add p4
    End If
    
    Set p4 = Nothing
        
End Sub

