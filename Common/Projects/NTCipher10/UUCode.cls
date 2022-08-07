VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UUCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary
'==================================================================================
' UUencode/decode in VB (c) 2/5/2000 by Alien3333@yahoo.com
' A small utility that uuencode and decode with a easy standard M$ GUI
'
' Well, I put this together in my 2 day vacation to celebrate Chinese New Year !
' All codes are standard as possible, ignore some of my clumsy C-style codes =)
' This small application show how to use filelistbox, listbox, drivelistbox, dirlistbox,
' Reading/Writing binary files, Reading text file line by line, Reading file in small portions,
' uuencode/decode with VB way, not C way of bit shifting, character and string manipulation,
' and all you can name it ...
'
' I learn VB through the Internet and MSDN, so some of the codes can look very familiar.
' I also read through uuencode.c and uudecode.c in LINUX to verify correctness.
'
' Enjoy !!!
'==================================================================================
' modified by Nicholas Forystek
'==================================================================================
Private FileName As String

Private Function plain_filename(ByVal FileName As String) As String
    Dim Length As Long
    Dim temp_str As String
    Dim i As Long
    Dim done As Boolean
    Dim left_str As String
    Dim right_str As String
    
    done = False
    
    temp_str = FileName
    
    Length = Len(temp_str)
    
    i = Length
    
    While done = False And i <> 0
        right_str = Right(temp_str, 1)
        If right_str = "\" Then
            done = True
        End If
        temp_str = Left(temp_str, i - 1)
        i = i - 1
    Wend
    
    plain_filename = Right(FileName, Length - Len(temp_str) - 1)

End Function

' 3 character into 4 characters
Private Function Encode(ByVal instring As String) As String
    On Error GoTo encodedone
    
    Dim outstring As String
    Dim i As Integer
    
    Dim y0 As Integer
    Dim y1 As Integer
    Dim y2 As Integer
    Dim y3 As Integer
    
    Dim x0 As Integer
    Dim x1 As Integer
    Dim x2 As Integer
    
    ' Very Important pad 3 byte to make 3 multiple
    ' This can add 1 or 2 extra NULL character to the end of the file
    ' Resulting a different file size, but no harm, for easier implementation
    
    If Len(instring) Mod 3 <> 0 Then
        instring = instring & String(3 - Len(instring) Mod 3, Chr$(0))
    End If
    
    For i = 1 To Len(instring) Step 3
        x0 = Asc(Mid(instring, i, 1))
        x1 = Asc(Mid(instring, i + 1, 1))
        x2 = Asc(Mid(instring, i + 2, 1))
        
        'MsgBox "x0=" + CStr(x0) + ", " + "x1=" + CStr(x1) + ", " + "x2=" + CStr(x2)
        
        y0 = (x0 \ 4 + 32)
        y1 = ((x0 Mod 4) * 16) + (x1 \ 16 + 32)
        y2 = ((x1 Mod 16) * 4) + (x2 \ 64 + 32)
        y3 = (x2 Mod 64) + 32
        
        If (y0 = 32) Then y0 = 96
        If (y1 = 32) Then y1 = 96
        If (y2 = 32) Then y2 = 96
        If (y3 = 32) Then y3 = 96
        
        'MsgBox "y0=" + CStr(y0) + ", " + "y1=" + CStr(y1) + ", " + "y2=" + CStr(y2) + ", " + "y3=" + CStr(y3)
        
        outstring = outstring + Chr(y0) + Chr(y1) + Chr(y2) + Chr(y3)
    
    Next i
encodedone:
    If Err Then Err.Clear
    Encode = outstring

End Function

' 4 character into 3 characters
Private Function Decode(ByVal instring As String) As String
    On Error GoTo decodedone
    
    Dim outstring As String
    
    Dim i As Integer
    
    Dim x0 As Integer
    Dim x1 As Integer
    Dim x2 As Integer
    
    Dim y0 As Integer
    Dim y1 As Integer
    Dim y2 As Integer
    Dim y3 As Integer
    
    For i = 1 To Len(instring) Step 4
        y0 = Asc(Mid(instring, i, 1))
        y1 = Asc(Mid(instring, i + 1, 1))
        y2 = Asc(Mid(instring, i + 2, 1))
        y3 = Asc(Mid(instring, i + 3, 1))
        
        If (y0 = 96) Then y0 = 32
        If (y1 = 96) Then y1 = 32
        If (y2 = 96) Then y2 = 32
        If (y3 = 96) Then y3 = 32
        
        'MsgBox "y0=" + CStr(y0) + ", " + "y1=" + CStr(y1) + ", " + "y2=" + CStr(y2) + ", " + "y3=" + CStr(y3)
        
        x0 = ((y0 - 32) * 4) + ((y1 - 32) \ 16)
        x1 = ((y1 Mod 16) * 16) + ((y2 - 32) \ 4)
        x2 = ((y2 Mod 4) * 64) + (y3 - 32)
        
        'MsgBox "x0=" + CStr(x0) + ", " + "x1=" + CStr(x1) + ", " + "x2=" + CStr(x2)
        
        outstring = outstring + Chr(x0) + Chr(x1) + Chr(x2)
    Next i
decodedone:

    If Err Then Err.Clear
    Decode = outstring

End Function

Public Sub UUEncode(ByVal FileName As String)
    On Error GoTo errencode
    Dim errmsg As String
    Dim outputfile As String
    Dim total As Long
    Dim remain As Long
    Dim current As Long
    Dim i As Long
    
    Dim portion_size As Long
    Dim instring As String
    Dim filenum1 As Integer
    Dim filenum2 As Integer
    
    portion_size = 45
    
    outputfile = FileName & ".uue"
   
        filenum1 = FreeFile
'open the original file as binary read
Open (FileName) For Binary Access Read Shared As #filenum1
    filenum2 = FreeFile
    
    'If PathExists(outputfile, True) Then Kill outputfile
'open the target file as binary write
Open (outputfile) For Binary Access Write As #filenum2
If LOF(filenum2) > 0 Then Seek #filenum2, LOF(filenum2)
'for standard uuencode compatibility
Put #filenum2, , "begin 644 " + plain_filename(FileName) + vbCrLf

'total number of full sized portion with "portion_size" bytes
        total = LOF(filenum1) \ portion_size

'remain hold the remaining bytes toward end of file
        remain = LOF(filenum1) Mod portion_size

'prepare instring to read "portion_size" bytes at a time
        instring$ = String(portion_size, Chr(0))

'current file position
        current = 1
    
    'for loop to read the portion one by one

        For i = 1 To total
          Get #filenum1, current, instring$

'use the ENC() for standard uuencode compatibility, pad "M"
              
                Put #filenum2, , ENC(portion_size) + Encode(instring$) + vbCrLf
                  current = current + portion_size
        Next
        
        instring = String(remain, Chr(0))
       
'get the remaining bytes toward end of the file
        Get #filenum1, current, instring$
        
'get the remaining bytes size and calculate ENC() for the last line
        
        Put #filenum2, , ENC(LOF(filenum1) - current + 1) + Encode(instring$) + vbCrLf
                
        Close #filenum1
 
'put "end" for standard uuencode compatibility
       
        Put #filenum2, , ENC(0) + vbCrLf + "end" + vbCrLf
        Close #filenum2
        
    Exit Sub
errencode:
    errmsg = Err.Description
    Err.Clear
    If filenum1 > 0 Then Close #filenum1
    If filenum2 > 0 Then Close #filenum2
    On Error GoTo 0
    Err.Raise 8, "UUEncode", "Error UUEncoding: " & errmsg
End Sub

Public Sub UUDecode(ByVal FileName As String)
    On Error GoTo errencode
    Dim errmsg As String
    
    Dim instring As String
    Dim outstring As String
    Dim has_begin As Boolean
    Dim filename2 As String
    
    Dim filenum1 As Long
    Dim filenum2 As Long
    Dim linsize As Long
    
    has_begin = False
    filenum1 = FreeFile

    Open (FileName) For Input As #filenum1& ' file opened for reading
    
    While Not EOF(filenum1&) And has_begin = False
        Line Input #filenum1&, instring
        If Left(instring, 6) = "begin " Then
            filename2 = GetFilePath(FileName) + "\" + RemoveArg(RemoveArg(instring, " "), " ")
            has_begin = True
        End If
    Wend
'    linsize = 4096
    If Not (has_begin = True) Then Err.Raise 8, "UUDecode", "No begin line found for uucoded file."
                                    
    If has_begin = True Then
        
        filenum2 = FreeFile
        If PathExists(filename2, True) Then Kill filename2
        Open (filename2) For Binary Access Write As #filenum2&
        
        While (Not EOF(filenum1&)) And (instring <> "end")
'            instring = String(linsize, Chr(0))
            Line Input #filenum1&, instring
'            linsize = Len(NextArg(instring, vbCrLf))
            outstring = Right(instring, Len(instring) - 1)
            
            If instring <> "end" Then
                'If Len(outstring) > 2 Then
                Put #filenum2&, , Decode(outstring)
                'End If
            End If
                        
        Wend
        Close #filenum2&
    End If
    Close #filenum1&

    Exit Sub
errencode:
    errmsg = Err.Description
    Err.Clear
    If filenum1 > 0 Then Close #filenum1&
    If filenum2 > 0 Then Close #filenum2&
    On Error GoTo 0
    Err.Raise 8, "UUEncode", "Error UUEncoding: " & errmsg
End Sub

Private Function ENC(ByVal c As Integer) As String
    If c = 0 Then
        ENC = "`"
    Else
        c = c + 32
        ENC = Chr(c)
    End If
End Function


