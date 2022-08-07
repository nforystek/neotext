VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary

Implements Krimpt

Private Sub Class_Initialize()
    Set Krimpt = New Krimpt
End Sub

Public Function EncryptString(ByVal Text As String, ByVal Key As String, Optional ByVal OutputInHex As Boolean = True) As String
    If Len(Text) < 1 Or Len(Key) < 1 Then
        Err.Raise 8, "NTCipher10.NCode", "Both length of Text and Key in characters, must be non zero."
    Else
        EncryptString = StrConv(Krimpt_EnKrimpt(StrConv(Text, vbFromUnicode), StrConv(Key, vbFromUnicode)), vbUnicode)
        If OutputInHex = True Then EncryptString = modHexStr.HesEncodeData(EncryptString)
    End If
End Function

Public Function DecryptString(ByVal Text As String, ByVal Key As String, Optional ByVal IsTextInHex As Boolean = True) As String
    If Len(Text) < 1 Or Len(Key) < 1 Then
        Err.Raise 8, "NTCipher10.NCode", "Both length of Text and Key in characters, must be non zero."
    Else
        If IsTextInHex = True Then Text = modHexStr.HesDecodeData(Text)
        DecryptString = StrConv(Krimpt_DeKrimpt(StrConv(Text, vbFromUnicode), StrConv(Key, vbFromUnicode)), vbUnicode)
    End If
End Function

Private Function EncryptByte(Info() As Byte, Seed() As Byte) As Byte()
    EncryptByte = Krimpt.EnKrimpt(Info, Seed)
End Function

Private Function DecryptByte(Info() As Byte, Seed() As Byte) As Byte()
    DecryptByte = Krimpt.DeKrimpt(Info, Seed)
End Function

   
Public Function HexDecode(ByVal Data As String) As String
    HexDecode = modHexStr.HesDecodeData(Data)
End Function
Public Function HexEncode(ByVal Data As String) As String
    HexEncode = modHexStr.HesEncodeData(Data)
End Function


Private Function Krimpt_DeKrimpt(Info() As Byte, Seed() As Byte) As Byte()
    Krimpt_DeKrimpt = Krimpt.DeKrimpt(Info, Seed)
End Function

Private Function Krimpt_EnKrimpt(Info() As Byte, Seed() As Byte) As Byte()
    Krimpt_EnKrimpt = Krimpt.EnKrimpt(Info, Seed)
End Function
