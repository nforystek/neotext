Attribute VB_Name = "modMathAPI"
Option Explicit
'TOP DOWN

Option Private Module
#If VBIDE = -1 Then
Public Declare Function DoubleToLowBitLong Lib "C:\Development\Neotext\Common\Binary\NTMathAPI.dll" (ByVal DoubleNumeric As Double) As Long
Public Declare Function DoubleToHighBitLong Lib "C:\Development\Neotext\Common\Binary\NTMathAPI.dll" (ByVal DoubleNumeric As Double) As Long
Public Declare Function LongToLowBitInt Lib "C:\Development\Neotext\Common\Binary\NTMathAPI.dll" (ByVal LongNumeric As Long) As Integer
Public Declare Function LongToHighBitInt Lib "C:\Development\Neotext\Common\Binary\NTMathAPI.dll" (ByVal LongNumeric As Long) As Integer
Public Declare Function IntToLowBitByte Lib "C:\Development\Neotext\Common\Binary\NTMathAPI.dll" (ByVal IntNumeric As Integer) As Byte
Public Declare Function IntToHighBitByte Lib "C:\Development\Neotext\Common\Binary\NTMathAPI.dll" (ByVal IntNumeric As Integer) As Byte
#Else
Public Declare Function DoubleToLowBitLong Lib "NTMathAPI.dll" (ByVal DoubleNumeric As Double) As Long
Public Declare Function DoubleToHighBitLong Lib "NTMathAPI.dll" (ByVal DoubleNumeric As Double) As Long
Public Declare Function LongToLowBitInt Lib "NTMathAPI.dll" (ByVal LongNumeric As Long) As Integer
Public Declare Function LongToHighBitInt Lib "NTMathAPI.dll" (ByVal LongNumeric As Long) As Integer
Public Declare Function IntToLowBitByte Lib "NTMathAPI.dll" (ByVal IntNumeric As Integer) As Byte
Public Declare Function IntToHighBitByte Lib "NTMathAPI.dll" (ByVal IntNumeric As Integer) As Byte
#End If
