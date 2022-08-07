#Const [True] = -1
#Const [False] = 0

Attribute VB_Name = "modBits"
#Const modBits = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module

Public Const Bit1 = 1
Public Const Bit2 = 2
Public Const Bit3 = 4
Public Const Bit4 = 8
Public Const Bit5 = 16
Public Const Bit6 = 32
Public Const Bit7 = 64
Public Const Bit8 = 128

Public Const Bit01 = 1
Public Const Bit02 = 2
Public Const Bit03 = 4
Public Const Bit04 = 8
Public Const Bit05 = 16
Public Const Bit06 = 32
Public Const Bit07 = 64
Public Const Bit08 = 128

Public Const Bit09 = 256
Public Const Bit10 = 512
Public Const Bit11 = 1024
Public Const Bit12 = 2048
Public Const Bit13 = 4096
Public Const Bit14 = 8192
Public Const Bit15 = 16384
Public Const Bit16 = 32768

Public Const Bit17 = 65536
Public Const Bit18 = 131072
Public Const Bit19 = 262144
Public Const Bit20 = 524288
Public Const Bit21 = 1048576
Public Const Bit22 = 2097152
Public Const Bit23 = 4194304
Public Const Bit24 = 8388608

Public Const Bit25 = 16777216
Public Const Bit26 = 33554432
Public Const Bit27 = 67108864
Public Const Bit28 = 134217728
Public Const Bit29 = 268435456
Public Const Bit30 = 536870912
Public Const Bit31 = 1073741824
Public Const Bit32 = 8388608

Public Property Let BitByte(ByRef BWord As Byte, ByRef bBit As Byte, ByRef nValue As Boolean)
    If (BWord And bBit) And (Not nValue) Then
        BWord = BWord - bBit
    ElseIf (Not (BWord And bBit)) And nValue Then
        BWord = BWord Or bBit
    End If
End Property

Public Property Get BitByte(ByRef BWord As Byte, ByRef bBit As Byte) As Boolean
    BitByte = (BWord And bBit)
End Property

Public Property Let BitWord(ByRef BWord As Integer, ByRef bBit As Integer, ByRef nValue As Boolean)
    If (BWord And bBit) And (Not nValue) Then
        BWord = BWord - bBit
    ElseIf (Not (BWord And bBit)) And nValue Then
        BWord = BWord Or bBit
    End If
End Property

Public Property Get BitWord(ByRef BWord As Integer, ByRef bBit As Integer) As Boolean
    BitWord = (BWord And bBit)
End Property

Public Property Let BitLong(ByRef Word As Long, ByRef Bit As Long, ByRef Value As Boolean)
    If (Word And (Bit)) And (Not Value) Then
        Word = Word - (Bit)
    ElseIf (Not (Word And (Bit))) And Value Then
        Word = Word Or (Bit)
    End If
End Property

Public Property Get BitLong(ByRef Word As Long, ByRef Bit As Long) As Boolean
    BitLong = (Word And (Bit))
End Property

Public Property Let BitSingle(ByRef Word As Single, ByRef Bit As Double, ByRef Value As Boolean)
    If (Word And Bit) And (Not Value) Then
        Word = Word - (Bit)
    ElseIf (Not (Word And (Bit))) And Value Then
        Word = Word Or (Bit)
    End If
End Property

Public Property Get BitSingle(ByRef Word As Single, ByRef Bit As Double) As Boolean
    BitSingle = (Word And Bit)
End Property