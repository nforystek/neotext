Attribute VB_Name = "Module1"

Option Explicit


Public Sub Main()

    'Debug.Print SearchPath("embed*.*", -1, "C:\", FindAll)
    'End

   Kill "C:\Development\Neotext\Common\Projects\NTNodes10\Stream.cls"
    FileCopy "C:\Development\Neotext\Common\Projects\Copy of NTNodes10\Stream.cls", "C:\Development\Neotext\Common\Projects\NTNodes10\Stream.cls"
    
    
     BuildFileDescriptions "C:\Development\Neotext\Common\Projects\NTNodes10\Stream.cls", True
     End
     
'    Dim tmp As String
'    Dim head As String
'    Dim out As String
'
'    tmp = ReadFile("C:\Development\Neotext\Common\Projects\NTNodes10\Stream.cls")
'
'
'
'    Do Until tmp = ""
'        If InStr(txt, " ' _" & vbCrLf) > 0 Then
'
'            head = StrReverse(NextArg(StrReverse(NextArg(txt, " ' _" & vbCrLf)), vbLf & vbCr)) & " ' _" & vbCrLf
'            FindNextHeader = StrReverse(RemoveArg(StrReverse(RemoveNextArg(txt, " ' _" & vbCrLf)), vbLf & vbCr))
'            Do While NextArg(txt, vbCrLf) = ""
'                RemoveNextArg txt, vbCrLf
'            Loop
'            If InStr(txt, vbCrLf & "Attribute ") <= InStr(txt, vbCrLf) And InStr(txt, """" & vbCrLf) > 0 Then
'                head = head & RemoveNextArg(txt, """" & vbCrLf) & """" & vbCrLf
'            Else
'                head = head & RemoveNextArg(txt, vbCrLf) & vbCrLf
'            End If
'            Debug.Print "HEAD:"
'            Debug.Print head
'            Debug.Print
'
'        Else
'            FindNextHeader = txt
'            txt = ""
'        End If
'    Loop
    
    End
End Sub





