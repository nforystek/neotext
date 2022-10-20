Attribute VB_Name = "modMain"
#Const modMain = True
Option Explicit
'TOP DOWN
Option Private Module
Private Const Base36$ = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"

Public Sub Main()

    Dim cmdLine As String
    cmdLine = Command
    
    If Left(Trim(cmdLine), 2) = "/?" Then
        MsgBox "Usage: DATABASE.EXE ""<MDB File path>"" ""<Source File Path>"" [""<Source File Path>""]", vbInformation, App.Title
    Else
    
        Dim mdbFile As String
        Dim clsFile As String
        Dim nowPassword As String
        Dim newPassword As String
        Dim SourceCode As String
        
        Randomize
        
        mdbFile = RemoveQuotedArg(cmdLine)
        newPassword = Left(Sessa2$(Int(1 + ((LongBound - 1) * Rnd))) & Sessa2$(Int(1 + ((LongBound - 1) * Rnd))) & Sessa2$(Int(1 + ((LongBound - 1) * Rnd))) & Sessa2$(Int(1 + ((LongBound - 1) * Rnd))), 12)
        
        If PathExists(mdbFile, True) Then

            If PathExists(Replace(LCase(mdbFile), ".mdb", ".pwd"), True) Then
                nowPassword = ReadFile(Replace(LCase(mdbFile), ".mdb", ".pwd"))
            End If
            WriteFile Replace(LCase(mdbFile), ".mdb", ".pwd"), newPassword
            
            Do
            
                clsFile = Trim(RemoveQuotedArg(cmdLine))
        
                If PathExists(clsFile, True) Then
            
                    SourceCode = ReadFile(clsFile)
                    SourceCode = Replace(SourceCode, nowPassword, newPassword)
                    WriteFile clsFile, SourceCode
            
                Else
                    MsgBox "Source Code file not found - [" & clsFile & "]", vbExclamation, App.Title
                End If
            
            Loop Until (Trim(cmdLine) = "")
            
            CompactDatabase mdbFile, nowPassword, newPassword
        Else
            MsgBox "Database file not found - [" & mdbFile & "]", vbExclamation, App.Title
        End If
        
    End If
    
End Sub

Private Function Decim&(Sessa2$)
    Dim Posiz%, CifraT&, ValC&
    CifraT& = 0
    For Posiz% = 1 To Len(Sessa2$)
        ValC& = (InStr(Base36$, Mid$(Sessa2$, Len(Sessa2$) - Posiz% + 1, 1)) - 1) * 36 ^ (Posiz% - 1)
        CifraT& = CifraT& + ValC&
    Next Posiz%
    Decim& = CifraT&
End Function

Private Function Sessa2$(Decim&)
    Dim DeScr&, CifraT$, Cifra%
    DeScr& = Decim&
    CifraT$ = ""
    Do
        Cifra% = DeScr& Mod 36
        DeScr& = DeScr& \ 36
        CifraT$ = Mid$(Base36$, Cifra% + 1, 1) + CifraT$
    Loop Until DeScr& = 0
    Sessa2$ = CifraT$
End Function

Public Function CompactDatabase(ByVal dbFile As String, ByVal nowPassword As String, ByVal newPassword As String) As String
    On Error Resume Next
    
    Dim JRO As New JRO.JetEngine
    Dim sPath As String
    Dim dPath As String

    sPath = dbFile
    dPath = Replace(dbFile, ".mdb", ".bak")
    
    If PathExists(dPath) Then Kill dPath
    JRO.CompactDatabase _
            "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sPath & ";Jet OLEDB:Database Password=" & nowPassword & ";", _
            "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dPath & ";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Password=" & newPassword & ";"

    If PathExists(dPath) Then
        Kill sPath
        FileCopy dPath, sPath
        Kill dPath
    End If

    Set JRO = Nothing
    
    CompactDatabase = Err.Description
    If (Err.Number <> 0) Then
        MsgBox "Error: " & Err.Number & ", " & Err.Description, vbCritical, App.Title
        Err.Clear
    End If
    On Error GoTo 0
End Function

