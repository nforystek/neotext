Attribute VB_Name = "modMain"
#Const [True] = -1
#Const [False] = 0
#Const modMain = -1
Option Explicit

Option Private Module

Private HyperLists As New Nodes


Public TestAction As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Type HalfLong
    bit1 As Byte
    int1 As Integer
    int2 As Integer
    bit2 As Byte
End Type

Type HalfBit
    bit1 As Byte
    int1 As Integer
    int2 As Integer
    bit2 As Byte
End Type

'the following functions are to allocate the memory of the nodes, you can change local to global for it
Private Declare Function LocalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LocalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LocalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LocalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef xDest As Long, ByRef xSource As Long, ByVal nbytes As Long)
Private Declare Sub RtlMoveMemoryAny Lib "kernel32" Alias "RtlMoveMemory" (ByRef xDest As Any, ByRef xSource As Any, ByVal nbytes As Long)

'Public Sub TestVariant()
'    Dim c As Currency
'    c = 23
'    HyperLists.Insert
'    HyperLists.Value = c
'    c = 0
'    c = 35
'    HyperLists.Insert
'    HyperLists.Value = c
'    c = HyperLists.Value
'
'
'    Debug.Print TypeName(c) & " " & c
'    HyperLists.forward
'
'    c = HyperLists.Value
'    Debug.Print TypeName(c) & " " & c
'End Sub
'
'Public Sub TestObject()
'    Dim fso As Object
'    Set fso = CreateObject("Scripting.FileSystemObject")
'
'    HyperLists.Insert
'
'
'    Set HyperLists.Object = fso.getfolder("C:\")
'
'
'    HyperLists.Insert
'    Set HyperLists.Object = fso.getfolder("D:\")
'    Set fso = Nothing
'
'    Set fso = HyperLists.Object
'
'
'
'    Debug.Print TypeName(HyperLists.Value) & " " & fso.path
'    HyperLists.forward
'
'    Set fso = HyperLists.Object
'    Debug.Print TypeName(fso) & " " & fso.path
'End Sub

Public Function DebugNodes(ByVal n As iNode, Optional ByVal FromPoint As Boolean = False) As String
    If n.Count = 0 Then
        DebugNodes = "@INVALID"
    Else
        
        Dim ptr As Long
        Dim eol As Long
        Dim bol As Long
        Dim chk As Long
        eol = n.Final
        bol = n.First
        ptr = n.Point
        chk = n.check
    
        If FromPoint Then
            Do Until n.Point = bol
                n.forward
            Loop
        End If
        
        Do
            DebugNodes = DebugNodes & " "
            If n.Point = ptr Then
                DebugNodes = DebugNodes & "^"
            End If
            If n.Point = chk Then
                DebugNodes = DebugNodes & "v"
            End If
            If n.Point = eol Then
                DebugNodes = DebugNodes & "E"
            End If
            If n.Point = bol Then
                DebugNodes = DebugNodes & "B"
            End If
            DebugNodes = DebugNodes & "@" & n.Point & " "
        
            n.forward
        Loop Until (n.Point = ptr And (Not FromPoint)) Or (n.Point = bol And FromPoint)
    
        If FromPoint Then
            Do Until n.Point = ptr
                n.forward
            Loop
        End If
        
        DebugNodes = Trim(DebugNodes)
    
    End If
    
End Function


Public Sub Main()

'
'   TestObject
''
''
'    HyperLists.Clear
''
'End
'
  '  TestVariant

  '  HyperLists.Clear

'End
'    HyperLists.append
'    HyperLists.append
'    HyperLists.append
'    HyperLists.append
'    HyperLists.append
'
'    DebugNodes HyperLists
'
'    HyperLists.prior
'    HyperLists.prior
'
'
'
'    DebugNodes HyperLists
'
'
'    HyperLists.Clear
'    End
    
  '  Set HyperLists = HyperLists.CreateList


' TestVariant
'    TestObject
''
' HyperLists.Clear
''
'    End

'    HyperLists.append
'    HyperLists.append
'    HyperLists.append
'    HyperLists.append
'    HyperLists.append
'
'    DebugNodes HyperLists
'
'    HyperLists.prior
'    HyperLists.prior
'
'
'
'    DebugNodes HyperLists
'
'
'    HyperLists.Clear
'    End
    
        
    ReadyInput
    frmMain.Show

    Dim lRGB As Long
    Dim bDir As Boolean
    Dim elapse As Long
    Dim latency As Long
    Dim printlag As Long
    bDir = True
    
    Do
        If InputLoop() Then
            elapse = GetTickCount
            If Pressed(VK_UP) Then
                TestAction = 1
            End If
            If Pressed(VK_DOWN) Then
                TestAction = 2
            End If
            If Pressed(VK_LEFT) Then
                TestAction = 3
            End If
            If Pressed(VK_RIGHT) Then
                TestAction = 4
            End If
            If Pressed(VK_INSERT) Then
                TestAction = 5
            End If
            If Pressed(VK_DELETE) Then
                TestAction = 6
            End If
            If (Toggled(VK_RCONTROL) And (TestAction > 0)) Then
                TestAction = -TestAction
            ElseIf (Not Toggled(VK_RCONTROL) And (TestAction < 0)) Then
                TestAction = -TestAction
            End If
            frmMain.Cls

            frmMain.Print "Insert/Delete to add or remove nodes at current location;"
            frmMain.Print "Up/Down to append to last of, or remove first of the list;"
            frmMain.Print "Right/Left to move the position of current point of list;"
            frmMain.Print "Green's current point, Blue's first Node, Red's Last node;"
            frmMain.Print "Home=Save, End=Clear, Return=Load; ";
            If HyperLists.Exists() Then
                frmMain.Print "Point(" & HyperLists.Handle & "); ";
            Else
                frmMain.Print "Point(nill); ";
            End If
            frmMain.Print
            frmMain.Print "Action: ";
            If TestAction = 1 Then
                frmMain.Print "Append; ";
            ElseIf TestAction = 2 Then
                frmMain.Print "Remove; ";
            ElseIf TestAction = 5 Then
                frmMain.Print "Insert; ";
            ElseIf TestAction = 6 Then
                frmMain.Print "Delete; ";
            Else
                frmMain.Print "None; ";
            End If
            frmMain.Print "Count: ";
            If (Not HyperLists.Exists()) Or (HyperLists.Count < 1) Then
                frmMain.Print "None; ";
            Else
                frmMain.Print HyperLists.Count & "; ";
            End If
            frmMain.Print "Motion: ";
            If TestAction = 3 Then
                frmMain.Print "Backward; ";
            ElseIf TestAction = 4 Then
                frmMain.Print "Forward; ";
            Else
                frmMain.Print "None; ";
            End If
            
            If latency <> printlag And latency <> 0 Then
                printlag = latency
            End If
            frmMain.Print "Latency: " & printlag
            
            If TestAction = 1 Then
                HyperLists.Insert
                If lRGB = 1 Then
                    bDir = True
                End If
                lRGB = lRGB + IIf(bDir, 1, -1)
                If lRGB = 255 Then
                    bDir = False
                End If
                HyperLists.Value = RGB(lRGB, lRGB, lRGB)
            ElseIf TestAction = 2 Then
                HyperLists.Delete
            ElseIf TestAction = 3 Then
                HyperLists.Backward
            ElseIf TestAction = 4 Then
                HyperLists.forward
            ElseIf TestAction = 5 Then
                HyperLists.Insert
                If lRGB = 1 Then
                    bDir = True
                End If
                lRGB = lRGB + IIf(bDir, 1, -1)
                If lRGB = 255 Then
                    bDir = False
                End If
                HyperLists.Value = RGB(lRGB, lRGB, lRGB)
            ElseIf TestAction = 6 Then
                HyperLists.Delete
            End If
            
           ' If Pressed(VK_HOME) Then
           '     HyperLists.SaveList HyperLists
           ' End If
            If Pressed(VK_END) Then
                HyperLists.Clear
            End If
           ' If Pressed(VK_RETURN) Then
           '     HyperLists.loadList HyperLists
           ' End If
            
            If Not HyperLists.Exists Then
                frmMain.Print "First: @INVALID Final: @INVALID Check: @INVALID"
            Else
                frmMain.Print "First: " & HyperLists.First & " Final: " & HyperLists.Final & " Check: " & HyperLists.check
            End If
            frmMain.Print DebugNodes(HyperLists, False)

            frmMain.DrawNodes HyperLists
                        
            TestAction = 0
            
            latency = CLng(GetTickCount - elapse)
            
        End If

        DoEvents
     
    Loop Until EndOfInput Or frmMain.Visible = False

    HyperLists.Clear
    Unload frmMain

End Sub


Public Property Get LoWord(ByRef lThis As Long) As Long
   LoWord = (lThis And &HFFFF&)
End Property

Public Property Let LoWord(ByRef lThis As Long, ByVal lLoWord As Long)
   lThis = lThis And Not &HFFFF& Or lLoWord
End Property

Public Property Get HiWord(ByRef lThis As Long) As Long
   If (lThis And &H80000000) = &H80000000 Then
      HiWord = ((lThis And &H7FFF0000) \ &H10000) Or &H8000&
   Else
      HiWord = (lThis And &HFFFF0000) \ &H10000
   End If
End Property

Public Property Let HiWord(ByRef lThis As Long, ByVal lHiWord As Long)
   If (lHiWord And &H8000&) = &H8000& Then
      lThis = lThis And Not &HFFFF0000 Or ((lHiWord And &H7FFF&) * &H10000) Or &H80000000
   Else
      lThis = lThis And Not &HFFFF0000 Or (lHiWord * &H10000)
   End If
End Property


