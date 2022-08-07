#Const [True] = -1
#Const [False] = 0


Attribute VB_Name = "modNCode"
#Const modNCode = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module
#If Not modBitValue = -1 Then

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

#End If

#If Not modHexStr Then

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Function HesEncodeData(ByVal d As String) As String
    Dim s As String
    Dim l As Long
    Dim i As Long
    Dim n As String
    l = Len(d)
    If l > 0 Then
        For i = 1 To l
            n = Hex(Asc(Mid(d, i, 1)))
            If Len(n) < 2 Then
                s = s & "0" & n
            Else
                s = s & n
            End If
        Next
    End If
    HesEncodeData = s
End Function

Private Function HesDecodeData(ByVal d As String) As String
    Dim s As String
    Dim l As Long
    Dim i As Long
    l = Len(d)
    If l > 0 Then
        For i = 1 To l Step 2
            s = s & Chr(val("&H" & Mid(d, i, 2)))
        Next
    End If
    HesDecodeData = s
End Function

Private Function IsHexidecimal(ByVal Text As String) As Boolean
    Dim cnt As Integer
    Dim c2 As Integer
    Dim retVal As Boolean
    retVal = True
    If (Len(Text) Mod 2) = 0 Then
        If Len(Text) > 0 Then
            For cnt = 1 To Len(Text)
                If (Not IsNumeric(Mid(Text, cnt, 1))) And (Not (Asc(LCase(Mid(Text, cnt, 1))) >= 97 And Asc(LCase(Mid(Text, cnt, 1))) <= 122)) Then
                    retVal = False
                    Exit For
                End If
            Next
        Else
            retVal = False
        End If
    End If
    IsHexidecimal = retVal And (c2 <= 1)
End Function

Private Function IsNexidecimal(ByVal Text As String) As Boolean
    Dim cnt As Integer
    Dim c2 As Integer
    Dim retVal As Boolean
    retVal = True
    If Len(Text) > 0 Then
        For cnt = 1 To Len(Text)
            If (Asc(LCase(Mid(Text, cnt, 1))) >= 97 And Asc(LCase(Mid(Text, cnt, 1))) <= 122) Then
                retVal = False
                Exit For
            End If
        Next
    Else
        retVal = False
    End If
    IsNexidecimal = retVal And (c2 <= 1)
End Function


#Else

Public Function HexDecode(ByVal Data As String) As String
    HexDecode = HesDecodeData(Data)
End Function
Public Function HexEncode(ByVal Data As String) As String
    HexEncode = HesEncodeData(Data)
End Function

Public Function HexDecode(ByVal Data As String) As String
    HexDecode = modHexStr.HesDecodeData(Data)
End Function
Public Function HexEncode(ByVal Data As String) As String
    HexEncode = modHexStr.HesEncodeData(Data)
End Function

#End If

#If Not modBitValue Then

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
#End If

Public Function EncryptString(ByVal Text As String, ByVal Key As String, Optional ByVal OutputInHex As Boolean = True) As String
    If Len(Text) < 1 Or Len(Key) < 1 Then
        Err.Raise 8, "NTCipher10.NCode", "Both length of Text and Key in characters, must be non zero."
    Else
        EncryptString = StrConv(EncryptByte(StrConv(Text, vbFromUnicode), StrConv(Key, vbFromUnicode)), vbUnicode)
        If OutputInHex = True Then EncryptString = HesEncodeData(EncryptString)
    End If
End Function

Public Function DecryptString(ByVal Text As String, ByVal Key As String, Optional ByVal IsTextInHex As Boolean = True) As String
    If Len(Text) < 1 Or Len(Key) < 1 Then
        Err.Raise 8, "NTCipher10.NCode", "Both length of Text and Key in characters, must be non zero."
    Else
        If IsTextInHex = True Then Text = HesDecodeData(Text)
        DecryptString = StrConv(DecryptByte(StrConv(Text, vbFromUnicode), StrConv(Key, vbFromUnicode)), vbUnicode)
    End If
End Function

Public Function EncryptByte(Info() As Byte, Seed() As Byte) As Byte()
    
    Dim pin As Byte
    Dim swp As Byte
    Dim cap As Boolean
   
    Dim cnt1 As Long
    Dim cnt2 As Long
    
    Dim lbi As Long
    Dim ubi As Long
    Dim lbs As Long
    Dim ubs As Long
    
    lbi = LBound(Info)
    ubi = UBound(Info)
    lbs = LBound(Seed)
    ubs = UBound(Seed)
    
    BitByte(pin, Bit1) = BitByte(Seed(lbs), Bit6) Or BitByte(Seed(ubs), Bit2)
    BitByte(pin, Bit2) = BitByte(Seed(ubs), Bit4) Or BitByte(Seed(lbs), Bit1)
    BitByte(pin, Bit3) = BitByte(Seed(lbs), Bit8) Or BitByte(Seed(ubs), Bit5)
    BitByte(pin, Bit4) = BitByte(Seed(ubs), Bit7) Or BitByte(Seed(lbs), Bit3)
    
    cap = (BitByte(pin, Bit1) Or BitByte(pin, Bit2)) And (BitByte(pin, Bit3) Or BitByte(pin, Bit4))
    
    For cnt1 = lbs To ubs
    
        For cnt2 = lbi To ubi
        
            Select Case (-BitByte(Seed(cnt1), Bit1)) & (-BitByte(Seed(cnt1), Bit2)) & (-BitByte(Info(cnt2), Bit1)) & (-BitByte(Info(cnt2), Bit2))
                Case "0011"
                    BitByte(Info(cnt2), Bit1) = False
                    BitByte(Info(cnt2), Bit2) = False
                Case "0000"
                    BitByte(Info(cnt2), Bit1) = True
                    BitByte(Info(cnt2), Bit2) = True
                Case "0010"
                    BitByte(Info(cnt2), Bit1) = False
                    BitByte(Info(cnt2), Bit2) = True
                Case "0001"
                    BitByte(Info(cnt2), Bit1) = True
                    BitByte(Info(cnt2), Bit2) = False
    
                Case "1111"
                    BitByte(Info(cnt2), Bit1) = False
                    BitByte(Info(cnt2), Bit2) = False
                Case "1100"
                    BitByte(Info(cnt2), Bit1) = True
                    BitByte(Info(cnt2), Bit2) = False
                Case "1110"
                    BitByte(Info(cnt2), Bit1) = False
                    BitByte(Info(cnt2), Bit2) = True
                Case "1101"
                    BitByte(Info(cnt2), Bit1) = True
                    BitByte(Info(cnt2), Bit2) = True
    
                Case "1011"
                    BitByte(Info(cnt2), Bit1) = False
                    BitByte(Info(cnt2), Bit2) = True
                Case "1000"
                    BitByte(Info(cnt2), Bit1) = True
                    BitByte(Info(cnt2), Bit2) = True
                Case "1010"
                    BitByte(Info(cnt2), Bit1) = False
                    BitByte(Info(cnt2), Bit2) = False
                Case "1001"
                    BitByte(Info(cnt2), Bit1) = True
                    BitByte(Info(cnt2), Bit2) = False
    
                Case "0111"
                    BitByte(Info(cnt2), Bit1) = True
                    BitByte(Info(cnt2), Bit2) = False
                Case "0100"
                    BitByte(Info(cnt2), Bit1) = False
                    BitByte(Info(cnt2), Bit2) = True
                Case "0110"
                    BitByte(Info(cnt2), Bit1) = True
                    BitByte(Info(cnt2), Bit2) = True
                Case "0101"
                    BitByte(Info(cnt2), Bit1) = False
                    BitByte(Info(cnt2), Bit2) = False
            End Select
            
        Next
    Next
    
    For cnt1 = lbi To ubi

        Select Case (-BitByte(pin, Bit1)) & (-BitByte(pin, Bit2)) & (-BitByte(Info(cnt1), Bit3)) & (-BitByte(Info(cnt1), Bit4))
            Case "0011"
                BitByte(Info(cnt1), Bit3) = False
                BitByte(Info(cnt1), Bit4) = False
            Case "0000"
                BitByte(Info(cnt1), Bit3) = True
                BitByte(Info(cnt1), Bit4) = True
            Case "0010"
                BitByte(Info(cnt1), Bit3) = False
                BitByte(Info(cnt1), Bit4) = True
            Case "0001"
                BitByte(Info(cnt1), Bit3) = True
                BitByte(Info(cnt1), Bit4) = False

            Case "1111"
                BitByte(Info(cnt1), Bit3) = False
                BitByte(Info(cnt1), Bit4) = False
            Case "1100"
                BitByte(Info(cnt1), Bit3) = True
                BitByte(Info(cnt1), Bit4) = False
            Case "1110"
                BitByte(Info(cnt1), Bit3) = False
                BitByte(Info(cnt1), Bit4) = True
            Case "1101"
                BitByte(Info(cnt1), Bit3) = True
                BitByte(Info(cnt1), Bit4) = True

            Case "1011"
                BitByte(Info(cnt1), Bit3) = False
                BitByte(Info(cnt1), Bit4) = True
            Case "1000"
                BitByte(Info(cnt1), Bit3) = True
                BitByte(Info(cnt1), Bit4) = True
            Case "1010"
                BitByte(Info(cnt1), Bit3) = False
                BitByte(Info(cnt1), Bit4) = False
            Case "1001"
                BitByte(Info(cnt1), Bit3) = True
                BitByte(Info(cnt1), Bit4) = False

            Case "0111"
                BitByte(Info(cnt1), Bit3) = True
                BitByte(Info(cnt1), Bit4) = False
            Case "0100"
                BitByte(Info(cnt1), Bit3) = False
                BitByte(Info(cnt1), Bit4) = True
            Case "0110"
                BitByte(Info(cnt1), Bit3) = True
                BitByte(Info(cnt1), Bit4) = True
            Case "0101"
                BitByte(Info(cnt1), Bit3) = False
                BitByte(Info(cnt1), Bit4) = False
        End Select
       
    Next
    
    BitByte(swp, Bit1) = BitByte(Info(lbi), IIf(cap, Bit5, Bit6))
    BitByte(swp, Bit2) = BitByte(Info(lbi), IIf(cap, Bit6, Bit7))
    BitByte(swp, Bit3) = BitByte(Info(lbi), IIf(cap, Bit7, Bit8))
    BitByte(swp, Bit4) = BitByte(Info(lbi), IIf(cap, Bit8, Bit5))
    For cnt1 = lbi To ubi - 1
        BitByte(Info(cnt1), IIf(cap, Bit6, Bit5)) = BitByte(Info(cnt1 + 1), IIf(cap, Bit5, Bit6))
        BitByte(Info(cnt1), IIf(cap, Bit7, Bit6)) = BitByte(Info(cnt1 + 1), IIf(cap, Bit6, Bit7))
        BitByte(Info(cnt1), IIf(cap, Bit8, Bit7)) = BitByte(Info(cnt1 + 1), IIf(cap, Bit7, Bit8))
        BitByte(Info(cnt1), IIf(cap, Bit5, Bit8)) = BitByte(Info(cnt1 + 1), IIf(cap, Bit8, Bit5))
    Next
    BitByte(Info(ubi), IIf(cap, Bit6, Bit5)) = BitByte(swp, Bit1)
    BitByte(Info(ubi), IIf(cap, Bit7, Bit6)) = BitByte(swp, Bit2)
    BitByte(Info(ubi), IIf(cap, Bit8, Bit7)) = BitByte(swp, Bit3)
    BitByte(Info(ubi), IIf(cap, Bit5, Bit8)) = BitByte(swp, Bit4)
        
    EncryptByte = Info
        
End Function

Public Function DecryptByte(Info() As Byte, Seed() As Byte) As Byte()
    
    Dim pin As Byte
    Dim swp As Byte
    Dim cap As Boolean
    
    Dim cnt1 As Long
    Dim cnt2 As Long
    
    Dim lbi As Long
    Dim ubi As Long
    Dim lbs As Long
    Dim ubs As Long
        
    lbi = LBound(Info)
    ubi = UBound(Info)
    lbs = LBound(Seed)
    ubs = UBound(Seed)
    
    BitByte(pin, Bit1) = BitByte(Seed(lbs), Bit6) Or BitByte(Seed(ubs), Bit2)
    BitByte(pin, Bit2) = BitByte(Seed(ubs), Bit4) Or BitByte(Seed(lbs), Bit1)
    BitByte(pin, Bit3) = BitByte(Seed(lbs), Bit8) Or BitByte(Seed(ubs), Bit5)
    BitByte(pin, Bit4) = BitByte(Seed(ubs), Bit7) Or BitByte(Seed(lbs), Bit3)
    
    cap = (BitByte(pin, Bit1) Or BitByte(pin, Bit2)) And (BitByte(pin, Bit3) Or BitByte(pin, Bit4))

    BitByte(swp, Bit1) = BitByte(Info(ubi), IIf(cap, Bit6, Bit5))
    BitByte(swp, Bit2) = BitByte(Info(ubi), IIf(cap, Bit7, Bit6))
    BitByte(swp, Bit3) = BitByte(Info(ubi), IIf(cap, Bit8, Bit7))
    BitByte(swp, Bit4) = BitByte(Info(ubi), IIf(cap, Bit5, Bit8))
    For cnt1 = ubi To (lbi + 1) Step -1
        
        BitByte(Info(cnt1), IIf(cap, Bit5, Bit6)) = BitByte(Info(cnt1 - 1), IIf(cap, Bit6, Bit5))
        BitByte(Info(cnt1), IIf(cap, Bit6, Bit7)) = BitByte(Info(cnt1 - 1), IIf(cap, Bit7, Bit6))
        BitByte(Info(cnt1), IIf(cap, Bit7, Bit8)) = BitByte(Info(cnt1 - 1), IIf(cap, Bit8, Bit7))
        BitByte(Info(cnt1), IIf(cap, Bit8, Bit5)) = BitByte(Info(cnt1 - 1), IIf(cap, Bit5, Bit8))
    Next
    BitByte(Info(lbi), IIf(cap, Bit5, Bit6)) = BitByte(swp, Bit1)
    BitByte(Info(lbi), IIf(cap, Bit6, Bit7)) = BitByte(swp, Bit2)
    BitByte(Info(lbi), IIf(cap, Bit7, Bit8)) = BitByte(swp, Bit3)
    BitByte(Info(lbi), IIf(cap, Bit8, Bit5)) = BitByte(swp, Bit4)
    
    For cnt1 = lbi To ubi
    
        Select Case (-BitByte(Info(cnt1), Bit3)) & (-BitByte(Info(cnt1), Bit4)) & (-BitByte(pin, Bit1)) & (-BitByte(pin, Bit2))
            Case "0000"
                BitByte(Info(cnt1), Bit3) = True
                BitByte(Info(cnt1), Bit4) = True
            Case "1100"
                BitByte(Info(cnt1), Bit3) = False
                BitByte(Info(cnt1), Bit4) = False
            Case "0100"
                BitByte(Info(cnt1), Bit3) = True
                BitByte(Info(cnt1), Bit4) = False
            Case "1000"
                BitByte(Info(cnt1), Bit3) = False
                BitByte(Info(cnt1), Bit4) = True

            Case "0011"
                BitByte(Info(cnt1), Bit3) = True
                BitByte(Info(cnt1), Bit4) = True
            Case "1011"
                BitByte(Info(cnt1), Bit3) = False
                BitByte(Info(cnt1), Bit4) = False
            Case "0111"
                BitByte(Info(cnt1), Bit3) = True
                BitByte(Info(cnt1), Bit4) = False
            Case "1111"
                BitByte(Info(cnt1), Bit3) = False
                BitByte(Info(cnt1), Bit4) = True

            Case "0110"
                BitByte(Info(cnt1), Bit3) = True
                BitByte(Info(cnt1), Bit4) = True
            Case "1110"
                BitByte(Info(cnt1), Bit3) = False
                BitByte(Info(cnt1), Bit4) = False
            Case "0010"
                BitByte(Info(cnt1), Bit3) = True
                BitByte(Info(cnt1), Bit4) = False
            Case "1010"
                BitByte(Info(cnt1), Bit3) = False
                BitByte(Info(cnt1), Bit4) = True

            Case "1001"
                BitByte(Info(cnt1), Bit3) = True
                BitByte(Info(cnt1), Bit4) = True
            Case "0101"
                BitByte(Info(cnt1), Bit3) = False
                BitByte(Info(cnt1), Bit4) = False
            Case "1101"
                BitByte(Info(cnt1), Bit3) = True
                BitByte(Info(cnt1), Bit4) = False
            Case "0001"
                BitByte(Info(cnt1), Bit3) = False
                BitByte(Info(cnt1), Bit4) = True

        End Select
        
    Next

    For cnt1 = ubs To lbs Step -1
        
        For cnt2 = lbi To ubi
            
            Select Case (-BitByte(Info(cnt2), Bit1)) & (-BitByte(Info(cnt2), Bit2)) & (-BitByte(Seed(cnt1), Bit1)) & (-BitByte(Seed(cnt1), Bit2))
                Case "0000"
                    BitByte(Info(cnt2), Bit1) = True
                    BitByte(Info(cnt2), Bit2) = True
                Case "1100"
                    BitByte(Info(cnt2), Bit1) = False
                    BitByte(Info(cnt2), Bit2) = False
                Case "0100"
                    BitByte(Info(cnt2), Bit1) = True
                    BitByte(Info(cnt2), Bit2) = False
                Case "1000"
                    BitByte(Info(cnt2), Bit1) = False
                    BitByte(Info(cnt2), Bit2) = True

                Case "0011"
                    BitByte(Info(cnt2), Bit1) = True
                    BitByte(Info(cnt2), Bit2) = True
                Case "1011"
                    BitByte(Info(cnt2), Bit1) = False
                    BitByte(Info(cnt2), Bit2) = False
                Case "0111"
                    BitByte(Info(cnt2), Bit1) = True
                    BitByte(Info(cnt2), Bit2) = False
                Case "1111"
                    BitByte(Info(cnt2), Bit1) = False
                    BitByte(Info(cnt2), Bit2) = True

                Case "0110"
                    BitByte(Info(cnt2), Bit1) = True
                    BitByte(Info(cnt2), Bit2) = True
                Case "1110"
                    BitByte(Info(cnt2), Bit1) = False
                    BitByte(Info(cnt2), Bit2) = False
                Case "0010"
                    BitByte(Info(cnt2), Bit1) = True
                    BitByte(Info(cnt2), Bit2) = False
                Case "1010"
                    BitByte(Info(cnt2), Bit1) = False
                    BitByte(Info(cnt2), Bit2) = True

                Case "1001"
                    BitByte(Info(cnt2), Bit1) = True
                    BitByte(Info(cnt2), Bit2) = True
                Case "0101"
                    BitByte(Info(cnt2), Bit1) = False
                    BitByte(Info(cnt2), Bit2) = False
                Case "1101"
                    BitByte(Info(cnt2), Bit1) = True
                    BitByte(Info(cnt2), Bit2) = False
                Case "0001"
                    BitByte(Info(cnt2), Bit1) = False
                    BitByte(Info(cnt2), Bit2) = True
    
            End Select
        
        Next
    
    Next
    
    DecryptByte = Info

End Function



