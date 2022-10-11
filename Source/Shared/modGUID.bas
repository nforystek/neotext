Attribute VB_Name = "modGuid"
Option Explicit
Option Compare Binary
Option Private Module
Private Type GuidType '16
A4 As Long '4
B2 As Integer '2
C2 As Integer '2
D8(0 To 7) As Byte '8
End Type
Private Declare Function CoCreateGuid Lib "ole32" (ByVal pGuid As Long) As Long
Private Const GPTR = &H40
Private Const GMEM_MOVEABLE = &H2
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Left As Any, Pass As Any, ByVal Right As Long)
Private Function Padding(ByVal Length As Long, ByVal Value As String, Optional ByVal PadWith As String = " ") As String
Padding = String(Abs((Length * Len(PadWith)) - (Len(Value) \ Len(PadWith))), PadWith) & Value
End Function
Public Function GUID() As String
    Dim lpGuid As Long
    Dim lcGuid As Long
    lpGuid = GlobalAlloc(GMEM_MOVEABLE And VarPtr(lpGuid), 4)
    If lpGuid <> 0 Then
        Dim lgGuid As GuidType
        If CoCreateGuid(VarPtr(lgGuid)) = 0 Then
            RtlMoveMemory lgGuid, ByVal lpGuid, 4&
            lcGuid = GlobalLock(lpGuid)
            If lcGuid = lpGuid Then
                Dim ba(0 To 15) As Byte '16
                RtlMoveMemory ByVal VarPtr(ba(0)), ByVal VarPtr(lgGuid.A4) + 0, 16
                RtlMoveMemory ByVal VarPtr(ba(0)), ByVal VarPtr(ba(1)), 1
                RtlMoveMemory ByVal VarPtr(ba(1)), ByVal VarPtr(lgGuid.A4) + 1, 15
                RtlMoveMemory ByVal VarPtr(ba(1)), ByVal VarPtr(ba(2)), 1
                RtlMoveMemory ByVal VarPtr(ba(2)), ByVal VarPtr(lgGuid.A4) + 2, 14
                RtlMoveMemory ByVal VarPtr(ba(2)), ByVal VarPtr(ba(3)), 1
                RtlMoveMemory ByVal VarPtr(ba(3)), ByVal VarPtr(lgGuid.A4) + 3, 13
                RtlMoveMemory ByVal VarPtr(ba(3)), ByVal VarPtr(ba(4)), 1
                GlobalUnlock lcGuid
                RtlMoveMemory ByVal VarPtr(ba(4)), ByVal VarPtr(lgGuid.B2) + 0, 12
                RtlMoveMemory ByVal VarPtr(ba(4)), ByVal VarPtr(ba(5)), 1
                RtlMoveMemory ByVal VarPtr(ba(5)), ByVal VarPtr(lgGuid.B2) + 1, 11
                RtlMoveMemory ByVal VarPtr(ba(5)), ByVal VarPtr(ba(6)), 1
                lcGuid = GlobalLock(lpGuid)
                RtlMoveMemory ByVal VarPtr(ba(6)), ByVal VarPtr(lgGuid.C2) + 0, 10
                RtlMoveMemory ByVal VarPtr(ba(6)), ByVal VarPtr(ba(7)), 1
                RtlMoveMemory ByVal VarPtr(ba(7)), ByVal VarPtr(lgGuid.C2) + 1, 9
                RtlMoveMemory ByVal VarPtr(ba(7)), ByVal VarPtr(ba(8)), 1
                GlobalUnlock lcGuid
                RtlMoveMemory ByVal VarPtr(ba(7)), ByVal VarPtr(lgGuid.D8(0)), 1
                RtlMoveMemory ByVal VarPtr(ba(8)), ByVal VarPtr(ba(9)), 1
                RtlMoveMemory ByVal VarPtr(ba(6)), ByVal VarPtr(lgGuid.D8(1)), 1
                RtlMoveMemory ByVal VarPtr(ba(9)), ByVal VarPtr(ba(10)), 1
                lcGuid = GlobalLock(lpGuid)
                RtlMoveMemory ByVal VarPtr(ba(5)), ByVal VarPtr(lgGuid.D8(2)), 1
                RtlMoveMemory ByVal VarPtr(ba(10)), ByVal VarPtr(ba(11)), 1
                RtlMoveMemory ByVal VarPtr(ba(4)), ByVal VarPtr(lgGuid.D8(3)), 1
                RtlMoveMemory ByVal VarPtr(ba(11)), ByVal VarPtr(ba(12)), 1
                RtlMoveMemory ByVal VarPtr(ba(3)), ByVal VarPtr(lgGuid.D8(4)), 1
                RtlMoveMemory ByVal VarPtr(ba(12)), ByVal VarPtr(ba(13)), 1
                RtlMoveMemory ByVal VarPtr(ba(2)), ByVal VarPtr(lgGuid.D8(5)), 1
                RtlMoveMemory ByVal VarPtr(ba(13)), ByVal VarPtr(ba(14)), 1
                RtlMoveMemory ByVal VarPtr(ba(1)), ByVal VarPtr(lgGuid.D8(6)), 1
                RtlMoveMemory ByVal VarPtr(ba(14)), ByVal VarPtr(ba(15)), 1
                RtlMoveMemory ByVal VarPtr(ba(0)), ByVal VarPtr(lgGuid.D8(7)), 1
                RtlMoveMemory ByVal VarPtr(ba(15)), ByVal VarPtr(ba(0)), 0
                GlobalUnlock lcGuid
            End If
        End If
        GlobalFree lpGuid
        lpGuid = ((UBound(ba) + 1) / 4)
        For lcGuid = 1 To (UBound(ba) + 1)
            GUID = GUID & Padding(2, Hex(ba(lcGuid - 1)), "0")
            If ((lcGuid Mod lpGuid) = 0) Then
                If ((lpGuid * lcGuid) = (UBound(ba) + 1)) Then
                    lpGuid = (lpGuid / 2)
                ElseIf (lpGuid <= (UBound(ba) + 1) / 2) Then
                    lpGuid = ((UBound(ba) + 1) / lpGuid)
                ElseIf (lpGuid <> lcGuid) Then
                    lpGuid = (lpGuid + 1)
                End If
                If (lcGuid < (UBound(ba) + 1)) Then GUID = GUID & "-"
            End If
        Next
    Else
        Debug.Print "Error: GlobalAlloc " & Err.Number & " " & Err.Description
    End If
End Function
Public Function IsGuid(ByVal Value As Variant, Optional ByVal Acolyte As Boolean = True) As Boolean
Value = Replace(Replace(Value, "}", ""), "{", "")
If Not (Len(Value) = 36) And (InStr(Value, ".") = 0) Then
IsGuid = False
ElseIf Mid(Value, 9, 1) = "-" And _
 Mid(Value, 14, 1) = "-" And _
 Mid(Value, 19, 1) = "-" And _
Mid(Value, 24, 1) = "-" Then
Dim tmp As Variant
tmp = Value
Dim cnt As Byte
For cnt = Asc("0") To Asc("9")
tmp = Replace(tmp, Chr(cnt), "")
Next
For cnt = Asc("A") To Asc("F")
tmp = Replace(UCase(tmp), Chr(cnt), "")
Next
IsGuid = (tmp = "----") Or (tmp = "---")
End If
End Function


