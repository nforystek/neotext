VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWord"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary


Public Enum Bit
    IO1 = modBitValue.Bit1
    IO2 = modBitValue.Bit2
    IO3 = modBitValue.Bit3
    IO4 = modBitValue.Bit4
    IO5 = modBitValue.Bit5
    IO6 = modBitValue.Bit6
    IO7 = modBitValue.Bit7
    IO8 = modBitValue.Bit8
    
    IO9 = modBitValue.Bit09
    I10 = modBitValue.Bit10
    I11 = modBitValue.Bit11
    I12 = modBitValue.Bit12
    I13 = modBitValue.Bit13
    I14 = modBitValue.Bit14
    I15 = modBitValue.Bit15
    I16 = modBitValue.Bit16
    I17 = modBitValue.Bit17
    I18 = modBitValue.Bit18
    I19 = modBitValue.Bit19
    I20 = modBitValue.Bit20
    I21 = modBitValue.Bit21
    I22 = modBitValue.Bit22
    I23 = modBitValue.Bit23
    I24 = modBitValue.Bit24

End Enum

Public Property Let Value(ByRef BWord As Variant, ByRef bBit As Variant, ByRef nValue As Boolean)
    Select Case TypeName(BWord)
        Case "Byte"
            modBitValue.BitByte(CByte(BWord), CByte(bBit)) = nValue
        Case "Integer"
            modBitValue.BitWord(CInt(BWord), CInt(bBit)) = nValue
        Case "Long"
            modBitValue.BitLong(CLng(BWord), CLng(bBit)) = nValue
        Case Else
            Err.Raise 13, App.EXEName, "Type mismatch"
    End Select
End Property

Public Property Get Value(ByRef BWord As Variant, ByRef bBit As Variant) As Boolean
    Select Case TypeName(BWord)
        Case "Byte"
            Value = modBitValue.BitByte(CByte(BWord), CByte(bBit))
        Case "Integer"
            Value = modBitValue.BitWord(CInt(BWord), CInt(bBit))
        Case "Long"
            Value = modBitValue.BitLong(CLng(BWord), CLng(bBit))
        Case Else
            Err.Raise 13, App.EXEName, "Type mismatch"
    End Select
End Property

Public Property Let High(ByRef BWord As Variant, ByRef nValue As Variant)
    Select Case TypeName(BWord)
        Case "Byte"
            modBitValue.HiByte(CByte(BWord)) = nValue
        Case "Integer"
            modBitValue.HiWord(CInt(BWord)) = nValue
        Case "Long"
            modBitValue.HiLong(CLng(BWord)) = nValue
        Case Else
            Err.Raise 13, App.EXEName, "Type mismatch"
    End Select
End Property

Public Property Get High(ByRef BWord As Variant) As Variant
    Select Case TypeName(BWord)
        Case "Byte"
            High = modBitValue.HiByte(CByte(BWord))
        Case "Integer"
            High = modBitValue.HiWord(CInt(BWord))
        Case "Long"
            High = modBitValue.HiLong(CLng(BWord))
        Case Else
            Err.Raise 13, App.EXEName, "Type mismatch"
    End Select
End Property

Public Property Let Low(ByRef BWord As Variant, ByRef nValue As Variant)
    Select Case TypeName(BWord)
        Case "Byte"
            modBitValue.LoByte(CByte(BWord)) = nValue
        Case "Integer"
            modBitValue.LoWord(CInt(BWord)) = nValue
        Case "Long"
            modBitValue.LoLong(CLng(BWord)) = nValue
        Case Else
            Err.Raise 13, App.EXEName, "Type mismatch"
    End Select
End Property

Public Property Get Low(ByRef BWord As Variant) As Variant
    Select Case TypeName(BWord)
        Case "Byte"
            Low = modBitValue.LoByte(CByte(BWord))
        Case "Integer"
            Low = modBitValue.LoWord(CInt(BWord))
        Case "Long"
            Low = modBitValue.LoLong(CLng(BWord))
        Case Else
            Err.Raise 13, App.EXEName, "Type mismatch"
    End Select
End Property
