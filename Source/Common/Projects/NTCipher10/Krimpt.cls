VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Krimpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary

'Public Function EnKrimptString(Info As String, Seed As String) As String
'    EnKrimptString = Convert(EnKrimpt(Convert(Info), Convert(Seed)))
'End Function
Public Function EnKrimpt(Info() As Byte, Seed() As Byte) As Byte()
    
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
        
    EnKrimpt = Info
        
End Function

'Public Function DeKrimptString(Info As String, Seed As String) As String
'    DeKrimptString = Convert(DeKrimpt(Convert(Info), Convert(Seed)))
'End Function
Public Function DeKrimpt(Info() As Byte, Seed() As Byte) As Byte()
    
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
    
    DeKrimpt = Info

End Function


