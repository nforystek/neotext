Attribute VB_Name = "modBitValue"

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

Public Const BitReg = 4

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

Private Type TwoLong
   lVal(0 To 1) As Long
End Type

Private Declare Sub GetMem8 Lib "MSVBVM60.DLL" (ByRef pSrc As Any, ByRef pDest As Any)

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
    Dim Check As Double
    Dim Count As Long
    Check = Index
    Count = 1
    Do Until ((Check = 0) Or ((Check Mod 1) <> Check))
        Check = Index Mod 2
        Count = Count + 1
    Loop
    If Check = 0 Then
        If Count <= 24 Then
            FindBit = Index
        Else
            Select Case VBA.TypeName(Index)
                Case "Byte", "Integer", "Long"
                    FindBit = Index
                Case Else
                    Err.Raise 8, "Bit", "Invalid arguments."
            End Select
        End If
    ElseIf ((Check Mod 1) <> Check) Then
        If Index <= 24 Then
            Count = 1
            Do Until Index = 0
                Count = Count * 2
                Index = Index - 1
            Loop
            Count = Count \ 2
            FindBit = Count
        Else
            FindBit = Index
        End If
    Else
        Err.Raise 8, "Bit", "Invalid arguments."
    End If

End Function

'########################################################################################
'#### Property ByteBound/IntBound/LongBound/HighBound/DoubleBound/CurBound(Currency) ####
'########################################################################################
'#### Returns the maximum numerical value the respective datatype can contain with ######
'#### the optional parameter to return the unsigned version, (i.e. int unsigned has #####
'#### no negative number spectrum 0 to 65535, while signed, it does -32767 to 32767) ####
'########################################################################################

Public Function ByteBound(Optional ByVal UnSigned As Boolean = False) As Variant

    ByteBound = (BitReg ^ BitReg)
    If Not UnSigned Then ByteBound = (ByteBound - 1)
    
    'ByteBound = Num_255
End Function
Public Function IntBound(Optional ByVal UnSigned As Boolean = False) As Variant
    If UnSigned Then
        IntBound = (((BitReg * BitReg) ^ BitReg) - 1)
    Else
        IntBound = ((((BitReg + BitReg) * BitReg) ^ 3) - 1)
    End If
    
'    If UnSigned Then
'        IntBound = Num_65535
'    Else
'        IntBound = 32767
'    End If
End Function
Public Function LongBound(Optional ByVal UnSigned As Boolean = False) As Variant

    If UnSigned Then
        LongBound = ((BitReg ^ BitReg ^ BitReg) - ((BitReg * BitReg * BitReg) ^ BitReg))
    Else
        LongBound = (((BitReg ^ BitReg ^ BitReg) / 2) - 1)
    End If

End Function
Public Function HighBound(Optional ByVal UnSigned As Boolean = False) As Variant

    If UnSigned Then
        HighBound = CStr(CDec(((((BitReg * 2) ^ (BitReg * 2) ^ (1 / 0.5)) * 3.2768) - 1)) * CDec(2) * CDec(3.2768))
    Else
        HighBound = ((((BitReg * 2) ^ (BitReg * 2) ^ (1 / 0.5)) * 3.2768) - 1) ' CStr(CDec(((((BitReg * 2) ^ (BitReg * 2) ^ (1 / 0.5)) * 3.2768) - 1)) * CDec(2) * CDec(3.2768))
    End If

End Function

Public Function DoubleBound(Optional ByVal UnSigned As Boolean = False) As Variant

    If Not UnSigned Then
        DoubleBound = -CCur(CCur(-CurBound(False) + 0.0001) / 10000)
    Else
        DoubleBound = CDbl(CStr(CDbl(((((BitReg * 2) ^ (BitReg * 2) ^ (1 / 0.5)) * 3.2768) - 1)) * CDbl(2) * CDbl(3.2768))) ' CStr(CDec(((((BitReg * 2) ^ (BitReg * 2) ^ (1 / 0.5)) * 3.2768) - 1)) * CDec(2) * CDec(3.2768))
    End If

End Function

Public Function CurBound(Optional ByVal UnSigned As Boolean = False) As Variant
    
    CurBound = CCur(CCur(HighBound(False)) + 0.9557)   '"922337203685477.5807")
    If UnSigned Then CurBound = HighBound(True) '-CCur(CCur(-CurBound + 0.0001) / 10000) '
  '  If UnSigned Then CurBound = CStr(CDec(((((BitReg * 2) ^ (BitReg * 2) ^ (1 / 0.5)) * 3.2768) - 1)) * CDec(2) * CDec(3.2768))
    
'    CurBound = ((((BitReg * BitReg) * ((BitReg ^ BitReg) ^ BitReg)) * (BitReg ^ BitReg)) ^ 2)
'    If UnSigned Then
'        CurBound = ((((BitReg ^ BitReg) * ((BitReg ^ BitReg) ^ BitReg)) * (BitReg ^ BitReg)) ^ 2)
'    Else
'        CurBound = ((((BitReg * BitReg) * ((BitReg ^ BitReg) ^ BitReg)) * (BitReg ^ BitReg)) ^ 2)
'    End If
'
End Function

'########################################################################################
'#### Property Bit ######################################################################
'########################################################################################
'#### A generic of the Bit* properties that follow, attempts to determine data type. ####
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
'#### Property BitByte ##################################################################
'########################################################################################
'#### Gets or sets a byte data type's (bthis) bit at bbit index (Bit01, ...) to value ###
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
'#### Property BitWord ##################################################################
'########################################################################################
'#### Gets or sets an int data type's (ithis) bit at bbit index (Bit01, ...) to value ###
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
'#### Property BitLong ##################################################################
'########################################################################################
'#### Gets or sets a long data type's (lthis) bit at bbit index (Bit01, ...) to value ###
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
'#### Property LoByte/HiByte ############################################################
'########################################################################################
'#### Gets or sets an int data type's (ithis) lo or hi order byte information value #####
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
'#### Property LoWord/HiWord ############################################################
'########################################################################################
'#### Gets or sets a long data type's (lthis) lo or hi order int information value ######
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


'########################################################################################
'#### Property LoLong/HiLong ############################################################
'########################################################################################
'#### Gets or sets a double data type's (lthis) lo or hi order long information value ###
'########################################################################################

Public Property Get LoLong(ByRef dThis As Double) As Long
    Dim rec As TwoLong
    GetMem8 dThis, rec
    LoLong = rec.lVal(0)
End Property
Public Property Let LoLong(ByRef dThis As Double, ByVal lLoLong As Long)
    Dim rec As TwoLong
    rec.lVal(1) = HiLong(dThis)
    GetMem8 rec, dThis
    rec.lVal(0) = lLoLong
    GetMem8 rec, dThis
End Property
Public Property Get HiLong(ByRef dThis As Double) As Long
    Dim rec As TwoLong
    GetMem8 dThis, rec
    HiLong = rec.lVal(1)
End Property

Public Property Let HiLong(ByRef dThis As Double, ByVal lHiLong As Long)
    Dim rec As TwoLong
    rec.lVal(0) = LoLong(dThis)
    GetMem8 rec, dThis
    rec.lVal(1) = lHiLong
    GetMem8 rec, dThis
End Property

'########################################################################################


