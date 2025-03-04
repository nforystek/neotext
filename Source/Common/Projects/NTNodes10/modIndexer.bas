Attribute VB_Name = "modIndexer"

Option Explicit

Option Compare Binary

Public Type IndexData
    NextLoc As Long
    StrSize As Long
End Type

Public Type FileData
    StartLoc As Long
    Indecies() As IndexData
End Type


'Public Function DebugPrint(ByRef FileCount As Long, ByRef FileIndex() As FileData, ByVal Resource As String) As String
'    If FileCount > 0 Then
'
'        Dim Handle As Long
'        Dim Index As Long
'        Dim tmp As String
'        Dim tmp2 As String
'        Dim Ret As String
'        Dim inc As Byte
'        Dim num As Integer
'        Dim Max As Long
'        Dim pos As Long
'        Dim tmp3 As Long
'
'        For Handle = 1 To FileCount
'            Ret = Ret & "FILE:"
'            tmp = String(FileLen(Resource), " ")
'
'            tmp2 = Mid(tmp, FileIndex(Handle).StartLoc)
'            tmp = Left(tmp, FileIndex(Handle).StartLoc - 1) & "@"
'
'            For Index = LBound(FileIndex(Handle).Indecies) To UBound(FileIndex(Handle).Indecies)
'                If FileIndex(Handle).Indecies(Index).StrSize > 0 Then
'                    tmp = tmp & String(FileIndex(Handle).Indecies(Index).StrSize - 1, "-")
'                    tmp2 = Mid(tmp2, FileIndex(Handle).Indecies(Index).StrSize)
'                End If
'
'                tmp = tmp & tmp2
'                If FileIndex(Handle).Indecies(Index).NextLoc > 0 Then
'                    tmp2 = Mid(tmp, FileIndex(Handle).Indecies(Index).NextLoc + 1)
'                    tmp = Left(tmp, FileIndex(Handle).Indecies(Index).NextLoc - 1) & "@"
'                End If
'            Next
'            tmp = RTrimStrip(tmp, " ")
'            If Len(tmp) > Max Then Max = Len(tmp)
'            Ret = Ret & tmp & vbCrLf
'
'        Next
'
'        num = FreeFile
'        Open Resource For Binary Lock Read As #num
'        tmp = "DATA:"
'        Do Until EOF(num)
'            Get #num, , inc
'            tmp = tmp & Chr(inc)
'        Loop
'        tmp = Left(tmp, Max + 5) & vbCrLf
'        Ret = tmp & Ret
'
'        Close #num
'
'        DebugPrint = Ret
'    End If
'End Function

Private Function SizeFileIndex(ByRef FileCount As Long, ByRef FileIndex() As FileData) As Long
    If FileCount > 0 Then
        Dim handle As Long
        Dim Index As Long
        Dim pos As Long
        For handle = 1 To FileCount
            For Index = LBound(FileIndex(handle).Indecies) To UBound(FileIndex(handle).Indecies)
                If Index = 1 Then
                    If (FileIndex(handle).StartLoc + FileIndex(handle).Indecies(Index).StrSize > pos) Then
                        pos = FileIndex(handle).StartLoc + FileIndex(handle).Indecies(Index).StrSize
                    End If
                Else
                    If (FileIndex(handle).Indecies(Index - 1).NextLoc + FileIndex(handle).Indecies(Index).StrSize > pos) Then
                        pos = FileIndex(handle).Indecies(Index - 1).NextLoc + FileIndex(handle).Indecies(Index).StrSize
                    End If
                End If
            Next
        Next
        SizeFileIndex = pos
    End If
End Function

Private Function SizeIndecies(ByRef FileCount As Long, ByRef FileIndex() As FileData) As Long
    If FileCount > 0 Then
        Dim handle As Long
        Dim Index As Long
        Dim pos As Long
        pos = 4
        For handle = 1 To FileCount
            pos = pos + 4
            For Index = LBound(FileIndex(handle).Indecies) To UBound(FileIndex(handle).Indecies)
                pos = pos + 8
            Next
        Next
    End If
    SizeIndecies = pos
End Function

Public Sub LoadIndecies(ByRef FileCount As Long, ByRef FileIndex() As FileData, ByVal Resource As String)
    Dim num As Integer
    Dim handle As Long
    Dim Index As Long
    Dim tmp As Long
    Dim pos As Long
    num = FreeFile
    Open Resource For Binary Lock Read As #num
        pos = LOF(num) - 3
        If pos > 0 Then
            Get #num, pos, tmp
            FileCount = tmp
            If FileCount > 0 Then
                ReDim FileIndex(1 To FileCount) As FileData
                For handle = 1 To FileCount
                    pos = pos - 4
                    Get #num, pos, tmp
                    FileIndex(handle).StartLoc = tmp
                    Index = 0
                    Do
                        Index = Index + 1
                        ReDim Preserve FileIndex(handle).Indecies(1 To Index) As IndexData
                        pos = pos - 4
                        Get #num, pos, tmp
                        FileIndex(handle).Indecies(Index).StrSize = tmp
                        pos = pos - 4
                        Get #num, pos, tmp
                        FileIndex(handle).Indecies(Index).NextLoc = tmp
                    Loop Until FileIndex(handle).Indecies(Index).NextLoc = 0
                Next
            Else
                Erase FileIndex
            End If
        Else
            FileCount = 0
        End If
    Close #num
End Sub

Public Sub SaveIndecies(ByRef FileCount As Long, ByRef FileIndex() As FileData, ByVal Resource As String)
    Dim num As Integer
    Dim handle As Long
    Dim pos As Long
    Dim tmp As Long
    num = FreeFile
    Open Resource For Append As #num
    Close #num
    Open Resource For Binary Lock Read As #num
    handle = SizeFileIndex(FileCount, FileIndex)
    pos = SizeIndecies(FileCount, FileIndex)
    tmp = handle
    If LOF(num) - handle > 0 Then
        If pos < LOF(num) - handle Then
            tmp = (LOF(num) - pos)
        End If
    End If
    tmp = tmp + 1
    If FileCount > 0 Then
        For handle = FileCount To 1 Step -1
            For pos = UBound(FileIndex(handle).Indecies) To LBound(FileIndex(handle).Indecies) Step -1
                If UBound(FileIndex(handle).Indecies) = pos Then
                    Put #num, tmp, CLng(0)
                Else
                    Put #num, tmp, FileIndex(handle).Indecies(pos).NextLoc
                End If
                tmp = tmp + 4
                Put #num, tmp, FileIndex(handle).Indecies(pos).StrSize
                tmp = tmp + 4
            Next
            Put #num, tmp, FileIndex(handle).StartLoc
            tmp = tmp + 4
        Next
    End If
    Put #num, tmp, FileCount
    Close #num
End Sub

Private Function FindFreeSpace(ByRef FileCount As Long, ByRef FileIndex() As FileData, ByRef StartPos As Long, ByVal SeekingSize As Long) As Long
    Dim handle As Long
    Dim Index As Long
    Dim Redo As Boolean
    Dim TotSize As Long
    If FileCount > 0 Then
        Do
            If TotSize <= 0 Then TotSize = SeekingSize
            Do
                Redo = False
                For handle = LBound(FileIndex) To UBound(FileIndex)
                    If StartPos >= FileIndex(handle).StartLoc And (StartPos <= FileIndex(handle).StartLoc + (FileIndex(handle).Indecies(LBound(FileIndex(handle).Indecies)).StrSize - 1)) Then
                        StartPos = FileIndex(handle).StartLoc + FileIndex(handle).Indecies(LBound(FileIndex(handle).Indecies)).StrSize
                        Redo = True
                    End If
                    If Not Redo Then
                        For Index = LBound(FileIndex(handle).Indecies) + 1 To UBound(FileIndex(handle).Indecies)
                            If StartPos >= FileIndex(handle).Indecies(Index - 1).NextLoc And (StartPos <= FileIndex(handle).Indecies(Index - 1).NextLoc + (FileIndex(handle).Indecies(Index).StrSize - 1)) Then
                                StartPos = FileIndex(handle).Indecies(Index - 1).NextLoc + FileIndex(handle).Indecies(Index).StrSize
                                Redo = True
                            End If
                        Next
                    End If
                Next
            Loop While Redo
            For handle = LBound(FileIndex) To UBound(FileIndex)
                If StartPos < FileIndex(handle).StartLoc Then
                    If (StartPos + (TotSize - 1)) >= FileIndex(handle).StartLoc Then
                        TotSize = (FileIndex(handle).StartLoc - StartPos)
                        If TotSize <= 0 Then Redo = True
                    End If
                End If
                If Not Redo Then
                    For Index = LBound(FileIndex(handle).Indecies) + 1 To UBound(FileIndex(handle).Indecies)
                        If StartPos < FileIndex(handle).Indecies(Index - 1).NextLoc Then
                            If (StartPos + (TotSize - 1)) >= FileIndex(handle).Indecies(Index - 1).NextLoc Then
                                TotSize = (FileIndex(handle).Indecies(Index - 1).NextLoc - StartPos)
                                If TotSize <= 0 Then Redo = True
                            End If
                        End If
                    Next
                End If
            Next
        Loop While Redo
    Else
        TotSize = SeekingSize
    End If
    FindFreeSpace = TotSize
End Function

Public Function SizeOfAlloc(ByRef FileCount As Long, ByRef FileIndex() As FileData, ByVal StartPos As Long) As Long
    If FileCount > 0 Then
        Dim handle As Long
        handle = GetHandleByLocation(FileCount, FileIndex, StartPos)
        If (handle >= 1) And (handle <= FileCount) Then
            Dim Index As Long
            For Index = LBound(FileIndex(handle).Indecies) To UBound(FileIndex(handle).Indecies)
                SizeOfAlloc = SizeOfAlloc + FileIndex(handle).Indecies(Index).StrSize
            Next
        End If
    End If
End Function

Public Function Allocate(ByRef FileCount As Long, ByRef FileIndex() As FileData, ByVal Size As Long) As Long
    If (Size > 0) Then
        Dim Amount As Long
        Dim StartPos As Long
        StartPos = 1
        FileCount = FileCount + 1
        Dim Index As Long
        Do Until Size <= 0
            Amount = FindFreeSpace(FileCount - 1, FileIndex, StartPos, Size)
            If Index = 0 Then ReDim Preserve FileIndex(1 To FileCount) As FileData
            Index = Index + 1
            ReDim Preserve FileIndex(FileCount).Indecies(1 To Index) As IndexData
            If Index = 1 Then
                FileIndex(FileCount).StartLoc = StartPos
            Else
                FileIndex(FileCount).Indecies(Index - 1).NextLoc = StartPos
            End If
            FileIndex(FileCount).Indecies(Index).NextLoc = 0
            FileIndex(FileCount).Indecies(Index).StrSize = Amount
            Size = Size - Amount
        Loop
        Allocate = FileIndex(FileCount).StartLoc
    Else
        Allocate = 0
    End If
End Function

Public Sub Dealloc(ByRef FileCount As Long, ByRef FileIndex() As FileData, ByVal StartLoc As Long)
    If FileCount > 0 Then
        Dim handle As Long
        handle = GetHandleByLocation(FileCount, FileIndex, StartLoc)
        If (handle >= 1) And (handle <= FileCount) Then
            If (FileCount = 1) Then
                Erase FileIndex
                FileCount = 0
            ElseIf (FileCount > 0) Then
                Dim Index As Long
                FileIndex(handle).StartLoc = FileIndex(FileCount).StartLoc
                ReDim FileIndex(handle).Indecies(LBound(FileIndex(FileCount).Indecies) To UBound(FileIndex(FileCount).Indecies)) As IndexData
                For Index = LBound(FileIndex(FileCount).Indecies) To UBound(FileIndex(FileCount).Indecies)
                    FileIndex(handle).Indecies(Index).NextLoc = FileIndex(FileCount).Indecies(Index).NextLoc
                    FileIndex(handle).Indecies(Index).StrSize = FileIndex(FileCount).Indecies(Index).StrSize
                Next
                ReDim Preserve FileIndex(1 To FileCount - 1) As FileData
                FileCount = FileCount - 1
            End If
        End If
    End If
End Sub

Public Sub Realloc(ByRef FileCount As Long, ByRef FileIndex() As FileData, ByVal StartLoc As Long, ByVal Size As Long)
    If FileCount > 0 Then
        If (Size > 0) Then
            Dim Data As String
            Dim handle As Long
            Dim Amount As Long
            Dim Index As Long
            handle = GetHandleByLocation(FileCount, FileIndex, StartLoc)
            If (handle >= 1) And (handle <= FileCount) Then
                Amount = SizeOfAlloc(FileCount, FileIndex, StartLoc)
                If Size < Amount Then
                    Amount = 0
                    Do Until Amount + FileIndex(handle).Indecies(Index + 1).StrSize > Size
                        Index = Index + 1
                        Amount = Amount + FileIndex(handle).Indecies(Index).StrSize
                    Loop
                    FileIndex(handle).Indecies(Index + 1).StrSize = Size - Amount
                    ReDim Preserve FileIndex(handle).Indecies(LBound(FileIndex(handle).Indecies) To Index + 1) As IndexData
                    FileIndex(handle).Indecies(Index + 1).NextLoc = 0
                ElseIf Size > Amount Then
                    Size = (Size - Amount)
                    Dim newHandle As Long
                    newHandle = GetHandleByLocation(FileCount, FileIndex, Allocate(FileCount, FileIndex, Size))
                    Amount = UBound(FileIndex(handle).Indecies)
                    ReDim Preserve FileIndex(handle).Indecies(LBound(FileIndex(handle).Indecies) To UBound(FileIndex(handle).Indecies) + UBound(FileIndex(newHandle).Indecies)) As IndexData
                    FileIndex(handle).Indecies(Amount).NextLoc = FileIndex(newHandle).StartLoc
                    For Index = LBound(FileIndex(newHandle).Indecies) To UBound(FileIndex(newHandle).Indecies)
                        FileIndex(handle).Indecies(Amount + Index).NextLoc = FileIndex(newHandle).Indecies(Index).NextLoc
                        FileIndex(handle).Indecies(Amount + Index).StrSize = FileIndex(newHandle).Indecies(Index).StrSize
                    Next
                    Dealloc FileCount, FileIndex, FileIndex(newHandle).StartLoc
                End If
            End If
        End If
    End If
End Sub

Public Sub SetAlloc(ByRef FileCount As Long, ByRef FileIndex() As FileData, ByVal StartLoc As Long, ByVal Data As String, ByVal Resource As String)
    If FileCount > 0 Then
        Dim handle As Long
        handle = GetHandleByLocation(FileCount, FileIndex, StartLoc)
        If (handle >= 1) And (handle <= FileCount) Then
            Dim num As Integer
            Dim pos As Long
            Dim Index As Long
            num = FreeFile
            Open Resource For Append As #num
            Close #num
            Open Resource For Binary Lock Write As #num
                With FileIndex(handle)
                    pos = .StartLoc
                    For Index = LBound(.Indecies) To UBound(.Indecies)
                        Put #num, pos, CStr(Left(Data, .Indecies(Index).StrSize))
                        Data = Mid(Data, .Indecies(Index).StrSize + 1)
                        If (Data = "") Then
                            Exit For
                        ElseIf (.Indecies(Index).NextLoc > 0) Then
                            pos = .Indecies(Index).NextLoc
                        End If
                    Next
                End With
            Close #num
        End If
    End If
End Sub

Public Function GetAlloc(ByRef FileCount As Long, ByRef FileIndex() As FileData, ByVal StartLoc As Long, ByVal Resource As String) As String
    If FileCount > 0 Then
        Dim handle As Long
        handle = GetHandleByLocation(FileCount, FileIndex, StartLoc)
        If (handle >= 1) And (handle <= FileCount) Then
            Dim num As Integer
            Dim Index As Long
            Dim pos As Long
            Dim tmp As String
            Dim dat As String
            num = FreeFile
            Open Resource For Binary Lock Read As #num
                With FileIndex(handle)
                    pos = .StartLoc
                    For Index = LBound(.Indecies) To UBound(.Indecies)
                        tmp = String(.Indecies(Index).StrSize, Chr(0))
                        Get #num, pos, tmp
                        dat = dat & Replace(tmp, Chr(0), "")
                        If (.Indecies(Index).NextLoc > 0) Then
                            pos = .Indecies(Index).NextLoc
                        End If
                    Next
                End With
            Close #num
            GetAlloc = dat
        End If
    End If
End Function

Private Function GetHandleByLocation(ByRef FileCount As Long, ByRef FileIndex() As FileData, ByVal StartLoc As Long) As Long
    If FileCount > 0 Then
        GetHandleByLocation = -1
        Dim handle As Long
        For handle = 1 To FileCount
             If FileIndex(handle).StartLoc = StartLoc Then GetHandleByLocation = handle
        Next
    End If
End Function

