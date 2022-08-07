Attribute VB_Name = "modImage"
#Const modImage = -1
Option Explicit
'TOP DOWN
Option Compare Binary


Option Private Module
Public Type ImageInfoType
    Exists As Boolean
    Valid As Boolean
    
    Name As String
    Title As String
    Ext As String
    
    Height As Long
    Width As Long
    Dims As String
    
    Desc As String
End Type

Public Function GetWidth(ByVal Dimensions As String) As Long
    RemoveNextArg Dimensions, "x"
    GetWidth = CLng(Dimensions)
End Function
Public Function GetHeight(ByVal Dimensions As String) As Long
    GetHeight = CInt(RemoveNextArg(Dimensions, "x"))
End Function
Public Function GetImageInfo(ByVal FileName As String) As ImageInfoType
    On Error GoTo catch
    
    Dim handle As Integer
    Dim byteArr(255) As Byte, i As Long
    Dim imgInfo As ImageInfoType

    With imgInfo
        .Exists = PathExists(FileName, True)
        .Valid = False
        .Name = GetFileName(FileName)
        .Title = GetFileTitle(FileName)
        .Ext = UCase(Trim(GetFileExt(FileName, , True)))
        .Height = 0
        .Width = 0
        .Dims = "0x0"
        
        If .Exists Then

            handle = FreeFile
            Open FileName For Binary As #handle
            Get handle, , byteArr

            If byteArr(0) = &HFF And byteArr(1) = &HD8 Then
                .Valid = True
                
                GetHeaderJPG handle, imgInfo
                
                .Ext = "JPG"
                Close #handle
            Else
                Close #handle

                If byteArr(0) = &H47 And byteArr(1) = &H49 And byteArr(2) = &H46 _
                And byteArr(3) = &H38 Then
                    .Width = byteArr(7) * 256 + byteArr(6)
                    .Height = byteArr(9) * 256 + byteArr(8)
                    .Valid = True
                    .Ext = "GIF"
                Else

                    If byteArr(0) = 66 And byteArr(1) = 77 Then
                        .Valid = True

                        If byteArr(14) = 40 Then
    
                            .Width = byteArr(21) * 256 ^ 3 + byteArr(20) * 256 ^ 2 _
                                       + byteArr(19) * 256 + byteArr(18)
                          
                            .Height = byteArr(25) * 256 ^ 3 + byteArr(24) * 256 ^ 2 _
                                        + byteArr(23) * 256 + byteArr(22)

                        ElseIf byteArr(17) = 12 Then

                            .Width = byteArr(19) * 256 + byteArr(18)
                            .Height = byteArr(21) * 256 + byteArr(20)
                          
                        End If

                        .Ext = "BMP"
                    Else

                        If byteArr(0) = &H89 And byteArr(1) = &H50 And byteArr(2) = &H4E _
                        And byteArr(3) = &H47 Then
                            .Width = byteArr(18) * 256 + byteArr(19)
                            .Height = byteArr(22) * 256 + byteArr(23)
                            .Valid = True
                            .Ext = "PNG"
                        End If
                    End If
                End If
            End If
        End If
        
catch:
        .Dims = .Height & "x" & .Width
        
        If (Not Err.Number = 0) Or (Not .Valid) Then
            Err.Clear
            
            .Valid = False
            .Desc = "(" & IIf(.Exists, "Invalid Image File", "Image Not Found") & ")"
        
        ElseIf .Valid Then
            
            .Desc = "(" & .Ext & ", " & .Dims & ")"
        
        End If
    
    End With
    
    GetImageInfo = imgInfo
    On Error GoTo 0

End Function

Private Function GetHeaderJPG(ByVal FileNumber As Long, ByRef imgInfo As ImageInfoType) As Boolean
    
    On Error Resume Next
    
    Seek #FileNumber, 3
    
    Dim Buffer() As Byte

    ReDim Buffer(0 To 3)
    Get #FileNumber, , Buffer

    Do While Buffer(0) = &HFF
        Select Case Buffer(1)

        Case &HE0

            ReDim Buffer(0 To (Buffer(2) * (2 ^ 8) + Buffer(3)) - 3)
            Get #FileNumber, , Buffer

            If Buffer(7) = 1 Then

            ElseIf Buffer(7) = 2 Then

            End If
        
        Case &HC0
            ReDim Buffer(0 To (Buffer(2) * (2 ^ 8) + Buffer(3)) - 3)
            Get #FileNumber, , Buffer

            imgInfo.Height = Buffer(1) * (2 ^ 8) + Buffer(2)

            imgInfo.Width = Buffer(3) * (2 ^ 8) + Buffer(4)
        
        Case Else

            ReDim Buffer(0 To (Buffer(2) * (2 ^ 8) + Buffer(3)) - 3)
            Get #FileNumber, , Buffer
            
        End Select
        
        ReDim Buffer(0 To 3)
        Get #FileNumber, , Buffer
    Loop
    
    If Err Then Err.Clear
    On Error GoTo 0
    
    GetHeaderJPG = True
    
End Function
