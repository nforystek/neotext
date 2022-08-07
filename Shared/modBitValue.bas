Attribute VB_Name = "modBitValue"
#Const [True] = -1
#Const [False] = 0

Option Explicit
'TOP DOWN
Option Compare Binary

'byte
Public Const Bit1 As Byte = &H1
Public Const Bit2 As Byte = &H2
Public Const Bit3 As Byte = &H4
Public Const Bit4 As Byte = &H8
Public Const Bit5 As Byte = &H10
Public Const Bit6 As Byte = &H20
Public Const Bit7 As Byte = &H40
Public Const Bit8 As Byte = &H80

'integer
Public Const Bit01 = &H1
Public Const Bit02 = &H2
Public Const Bit03 = &H4
Public Const Bit04 = &H8
Public Const Bit05 = &H10
Public Const Bit06 = &H20
Public Const Bit07 = &H40
Public Const Bit08 = &H80

Public Const Bit09 = &H100
Public Const Bit10 = &H200
Public Const Bit11 = &H400
Public Const Bit12 = &H800
Public Const Bit13 = &H1000
Public Const Bit14 = &H2000
Public Const Bit15 = &H4000
Public Const Bit16 = &H8000

'long
Public Const Bit17 = &H10000
Public Const Bit18 = &H20000
Public Const Bit19 = &H40000
Public Const Bit20 = &H80000
Public Const Bit21 = &H100000
Public Const Bit22 = &H200000
Public Const Bit23 = &H400000
Public Const Bit24 = &H800000

Public Const Bit25 = &H1000000
Public Const Bit26 = &H2000000
Public Const Bit27 = &H4000000
Public Const Bit29 = &H8000000
Public Const Bit30 = &H10000000
Public Const Bit31 = &H20000000
Public Const Bit32 = &H40000000

Public Const Num_128 = &H80
Public Const Num_255 = &HFF&
Public Const Num_256 = &H100&
Public Const Num_32767 = &H7FFF&
Public Const Num_32768 = &H8000&
Public Const Num_Neg_32768 = &H8000
Public Const Num_65280 = &HFF00&
Public Const Num_65535 = &HFFFF&
Public Const Num_65536 = &H10000
Public Const Num_Neg_65536 = &HFFFF0000
Public Const Num_142606336 = &H8800000
Public Const Num_285212671 = &H10FFFFFF
Public Const Num_285212672 = &H11000000
Public Const Num_2147418112 = &H7FFF0000
Public Const Num_Neg_2147483648 = &H80000000
Public Const Num_Neg_285212672 = &HEF000000


'########################################################################################
Private Function FindBit(ByVal Index As Variant) As Variant

    If Index < ByteBound And Index > -1 Then
        Index = CByte(Index)
    ElseIf Index < IntBound And Index > -1 Then
        Index = CInt(Index)
    ElseIf Index < LongBound And Index > -1 Then
        Index = CLng(Index)
    Else
        Err.Raise 8, "Bit", "Invalid arguments."
    End If
    Dim check As Double
    Dim count As Long
    check = Index
    count = 1
    Do Until ((check = 0) Or ((check Mod 1) <> check))
        check = Index Mod 2
        count = count + 1
    Loop
    If check = 0 Then
        If count <= 24 Then
            FindBit = Index
        Else
            Select Case TypeName(Index)
                Case "Byte", "Integer", "Long"
                    FindBit = Index
                Case Else
                    Err.Raise 8, "Bit", "Invalid arguments."
            End Select
        End If
    ElseIf ((check Mod 1) <> check) Then
        If Index <= 24 Then
            count = 1
            Do Until Index = 0
                count = count * 2
                Index = Index - 1
            Loop
            count = count \ 2
            FindBit = count
        Else
            FindBit = Index
        End If
    Else
        Err.Raise 8, "Bit", "Invalid arguments."
    End If

End Function
'########################################################################################

Public Function ByteBound(Optional ByVal UnSigned As Boolean = False) As Byte
    ByteBound = Num_255
End Function
Public Function IntBound(Optional ByVal UnSigned As Boolean = False) As Long
    If UnSigned Then
        IntBound = Num_65535
    Else
        IntBound = 32767
    End If
End Function
Public Function LongBound(Optional ByVal UnSigned As Boolean = False) As Double
    If UnSigned Then
        LongBound = CDbl((VBA.Round(CDbl((CDbl(2147483647) * CDbl(2)) / Num_285212672), 0) * Num_285212672))
    Else
        LongBound = CDbl(2147483647)
    End If
End Function
Public Function HighBound(Optional ByVal UnSigned As Boolean = False) As Variant
    If UnSigned Then
        HighBound = "600000000000000"
    Else
        HighBound = CDec("92233723685477")
    End If
End Function

'########################################################################################
Public Property Let Bit(ByRef This As Variant, ByVal Index As Variant, ByRef Value As Boolean)
    Index = FindBit(Index)
    If (This And (Index)) And (Not Value) Then
        This = This - (Index)
    ElseIf (Not (This And (Index))) And Value Then
        This = This Or (Index)
    End If
End Property
Public Property Get Bit(ByRef This As Variant, ByVal Index As Variant) As Boolean
    Index = FindBit(Index)
    Bit = (This And Index)
End Property
'########################################################################################
Public Property Let BitByte(ByRef bThis As Byte, ByRef bBit As Byte, ByRef nValue As Boolean)
    If (bThis And bBit) And (Not nValue) Then
        bThis = bThis - bBit
    ElseIf (Not (bThis And bBit)) And nValue Then
        bThis = bThis Or bBit
    End If
End Property
Public Property Get BitByte(ByRef bThis As Byte, ByRef bBit As Byte) As Boolean
    BitByte = (bThis And bBit)
End Property
'########################################################################################
Public Property Let BitWord(ByRef iThis As Integer, ByRef bBit As Integer, ByRef nValue As Boolean)
    If (iThis And bBit) And (Not nValue) Then
        iThis = iThis - bBit
    ElseIf (Not (iThis And bBit)) And nValue Then
        iThis = iThis Or bBit
    End If
End Property
Public Property Get BitWord(ByRef iThis As Integer, ByRef bBit As Integer) As Boolean
    BitWord = (iThis And bBit)
End Property
'########################################################################################
Public Property Let BitLong(ByRef lThis As Long, ByRef Bit As Long, ByRef Value As Boolean)
    If (lThis And (Bit)) And (Not Value) Then
        lThis = lThis - (Bit)
    ElseIf (Not (lThis And (Bit))) And Value Then
        lThis = lThis Or (Bit)
    End If
End Property
Public Property Get BitLong(ByRef lThis As Long, ByRef Bit As Long) As Boolean
    BitLong = (lThis And (Bit))
End Property
'########################################################################################

Public Property Let LoByte(ByRef iThis As Integer, ByVal bLoByte As Byte)
    iThis = ((iThis Mod Num_256) * Num_256) + bLoByte
End Property
Public Property Get LoByte(ByRef iThis As Integer) As Byte
    LoByte = iThis And Num_255
End Property
Public Property Let HiByte(ByRef iThis As Integer, ByVal bHiByte As Byte)
    iThis = ((iThis \ Num_256) * Num_256) + bHiByte
End Property
Public Property Get HiByte(ByRef iThis As Integer) As Byte
    HiByte = iThis \ Num_256 And Num_255
End Property

'########################################################################################

Public Property Get LoWord(ByRef lThis As Long) As Long
   LoWord = (lThis And Num_65535)
End Property

Public Property Let LoWord(ByRef lThis As Long, ByVal lLoWord As Long)
   lThis = lThis And Not Num_65535 Or lLoWord
End Property

Public Property Get HiWord(ByRef lThis As Long) As Long
   If (lThis And Num_Neg_2147483648) = Num_Neg_2147483648 Then
      HiWord = ((lThis And Num_2147418112) \ Num_65536) Or Num_32768
   Else
      HiWord = (lThis And Num_Neg_65536) \ Num_65536
   End If
End Property

Public Property Let HiWord(ByRef lThis As Long, ByVal lHiWord As Long)
   If (lHiWord And Num_32768) = Num_32768 Then
      lThis = lThis And Not Num_Neg_65536 Or ((lHiWord And Num_32767) * Num_65536) Or Num_Neg_2147483648
   Else
      lThis = lThis And Not Num_Neg_65536 Or (lHiWord * Num_65536)
   End If
End Property


'Public Property Get LoWord(ByRef lThis As Long) As Integer
'    LoWord = (lThis And Num_32767) Or (Num_Neg_32768 And ((lThis And Num_Neg_32768) = Num_Neg_32768))
'End Property
'Public Property Get HiWord(ByRef lThis As Long) As Integer
'    HiWord = ((lThis And Num_2147418112) \ Num_65536) Or (Num_Neg_2147483648 And (lThis < 0))
'End Property
'Public Property Let LoWord(ByRef lThis As Long, ByVal lLoWord As Integer)
'    lThis = lThis And Not Num_65535 Or lLoWord
'End Property
'Public Property Let HiWord(ByRef lThis As Long, ByVal lHiWord As Integer)
'    If (lHiWord And Num_32768) = Num_32768 Then
'       lThis = lThis And Not Num_Neg_65536 Or ((lHiWord And Num_32767) * Num_65536) Or Num_Neg_2147483648
'    Else
'       lThis = lThis And Not Num_Neg_65536 Or (lHiWord * Num_65536)
'    End If
'End Property


'########################################################################################
Public Property Get LoLong(ByRef dThis As Double) As Long
    If (dThis And Num_285212671) > Num_142606336 Then
        LoLong = (dThis And Num_285212671) - (Num_285212672 * (dThis \ Num_285212672))
    Else
        LoLong = dThis And Num_285212671
    End If
End Property
Public Property Let LoLong(ByRef dThis As Double, ByVal lLoLong As Long)
    dThis = lLoLong + (((dThis And Num_Neg_65536) \ Num_65536) * Num_65536)
End Property
Public Property Get HiLong(ByRef dThis As Double) As Long
    If ((dThis And Num_Neg_285212672) \ Num_285212672) + (dThis Mod Num_285212672) < 0 Then
        HiLong = -(((dThis And Num_Neg_285212672) \ Num_285212672) + (dThis Mod Num_285212672))
    Else
        HiLong = ((dThis And Num_Neg_285212672) \ Num_285212672)
    End If
End Property
Public Property Let HiLong(ByRef dThis As Double, ByVal lHiLong As Long)
    dThis = ((lHiLong And Num_Neg_65536) \ Num_65536) * Num_65536
End Property


'Public Property Let LoLong(ByRef dThis As Double, ByVal lLoLong As Long)
'    dThis = (lLoLong And Num_Neg_65536) \ Num_65536
'End Property
'Public Property Get LoLong(ByRef dThis As Double) As Long
'    If (dThis And Num_285212671) > Num_142606336 Then
'        LoLong = (dThis And Num_285212671) - (Num_285212672 * (dThis \ Num_285212672))
'    Else
'        LoLong = dThis And Num_285212671
'    End If
'End Property
'Public Property Let HiLong(ByRef dThis As Double, ByVal lHiLong As Long)
'    dThis = ((lHiLong And Num_Neg_65536) \ Num_65536) * Num_65536
'End Property
'Public Property Get HiLong(ByRef dThis As Double) As Long
'    If ((dThis And Num_Neg_285212672) \ Num_285212672) + (dThis Mod Num_285212672) < 0 Then
'        HiLong = -(((dThis And Num_Neg_285212672) \ Num_285212672) + (dThis Mod Num_285212672))
'    Else
'        HiLong = ((dThis And Num_Neg_285212672) \ Num_285212672)
'    End If
'End Property
'########################################################################################


Public Property Get lo(ByVal Op As Variant) As Variant
    Select Case TypeName(Op)
        Case "Integer"
            'return lo byte of op
             lo = LoByte(CInt(Op))
        Case "Long"
            'return lo word of op
            lo = HiWord(CLng(Op))
        Case Else
            Err.Raise 8, "Lo", "Invalid arguments."
    End Select
End Property
Public Property Get hi(ByVal Op As Variant) As Variant
    Select Case TypeName(Op)
        Case "Integer"
            'return lo byte of op
             hi = HiByte(CInt(Op))
        Case "Long"
            'return lo word of op
            hi = LoWord(CLng(Op))
        Case Else
            Err.Raise 8, "Hi", "Invalid arguments."
    End Select
End Property

Public Property Get Wo(ByVal lo As Variant, ByVal hi As Variant) As Variant
    Select Case TypeName(lo)
        Case "Byte"
            Select Case TypeName(hi)
                Case "Byte"
                    If hi And Num_128 Then
                       Wo = ((hi * Num_256) Or lo) Or Num_Neg_65536
                    Else
                       Wo = (hi * Num_256) Or lo
                    End If
                Case Else
                    Err.Raise 8, "Hi", "Invalid arguments."
            End Select
        Case "Integer", "Long"
            Select Case TypeName(hi)
                Case "Integer", "Long"
'                    Dim ret As Long
'                    LoWord(ret) = Lo
'                    HiWord(ret) = Hi
'
'
'                    Wo = CLng(ret)

                    Wo = (hi * Num_65536) Or (lo And Num_65535)
                Case Else
                    Err.Raise 8, "Hi", "Invalid arguments."
            End Select
        Case Else
            Err.Raise 8, "Hi", "Invalid arguments."
    End Select
End Property

'Public Property Get Wo(ByRef Lo As Variant, ByVal Hi As Variant) As Variant
'    Select Case TypeName(Lo)
'        Case "Byte"
'            Select Case TypeName(Hi)
'                Case "Byte"
'                    If Hi And Num_128 Then
'                       Wo = ((Hi * Num_Num_256) Or Lo) Or Num_Neg_65536
'                    Else
'                       Wo = (Hi * Num_Num_256) Or Lo
'                    End If
'                Case Else
'                    Err.Raise 8, "Hi", "Invalid arguments."
'            End Select
'        Case "Integer", "Long"
'            Select Case TypeName(Hi)
'                Case "Integer", "Long"
'                    Wo = (Hi * Num_65536) Or (Lo And Num_65535)
'                Case Else
'                    Err.Raise 8, "Hi", "Invalid arguments."
'            End Select
'        Case Else
'            Err.Raise 8, "Hi", "Invalid arguments."
'    End Select
'End Property
