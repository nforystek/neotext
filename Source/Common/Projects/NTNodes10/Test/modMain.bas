Attribute VB_Name = "modMain"
#Const modMain = -1
Option Explicit
'TOP DOWN

Option Private Module

Private HyperLists As New NTNodes10.Nodes
Private MyNodeList As NTNodes10.inode

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

Public Declare Function vbaObjSetAddref Lib "MSVBVM60.DLL" Alias "__vbaObjSetAddref" (dstObject As Any, ByVal srcObjPtr As Long) As Long
Public Declare Function vbaObjSet Lib "MSVBVM60.DLL" Alias "__vbaObjSet" (dstObject As Any, ByVal srcObjPtr As Long) As Long

Public Sub TestVariant()
    Dim c As Currency
    c = 23
    MyNodeList.Append
    MyNodeList.Value = c
    MyNodeList.Append
    c = 0
    c = 35
    MyNodeList.Value = c
    c = MyNodeList.Value
    Debug.Print TypeName(c) & " " & c
    
    MyNodeList.Forward
    MyNodeList.Forward
    MyNodeList.Forward
    c = MyNodeList.Value
    Debug.Print TypeName(c) & " " & c
End Sub

'Public Sub TestRecord()
'    Dim c As HalfLong
'    c.bit1 = 23
'    c.bit2 = 8
'    c.int1 = 44
'    c.int2 = 48
'
'    MyNodeList.append
'    MyNodeList.record = HyperLists.VarToArray(VarPtr(c), 4)
'    'c = 0
'   ' c = 35
'    MyNodeList.append
'    MyNodeList.record = HyperLists.VarToArray(VarPtr(c), 4)
'
'    RtlMoveMemory VarPtr(c), ByVal VarPtr(MyNodeList.record), 4
'
'   ' Debug.Print TypeName(c) & " " & c
'    MyNodeList.Forward
'
'    RtlMoveMemory VarPtr(c), ByVal VarPtr(MyNodeList.record), 4
'   ' c = MyNodeList.record
'   ' Debug.Print TypeName(c) & " " & c
'End Sub
'
Public Sub TestObject()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    MyNodeList.Append
    Set MyNodeList.Object = fso.getfolder("C:\")
    Debug.Print TypeName(MyNodeList.Object)
    
    MyNodeList.Append
    Set MyNodeList.Object = fso.getfolder("D:\")

    Set fso = Nothing
    Set fso = MyNodeList.Object
    Debug.Print TypeName(fso)
    Set fso = MyNodeList.Object
    Debug.Print TypeName(fso)
End Sub

Public Sub Testvalue()

    MyNodeList.Append
    MyNodeList.Value = 34
    
   ' Debug.Print MyNodeList.DataType
    
    MyNodeList.Append
    

    MyNodeList.Value = 45
    

    Debug.Print MyNodeList.Value
    MyNodeList.backward
    Debug.Print MyNodeList.Value
   
    
End Sub
''Public Sub NewObject(ByRef FromObj As Object, ByRef Ptr2 As Object)
''    Dim Ptr3 As Object
''    Dim Ptr4 As Long
''    Dim Ptr5 As Long
''    Dim Ptr6 As Object
''
''    vbaObjSet ByVal VarPtr(test3), ObjPtr(test1)
''    vbaObjSetAddref ByVal VarPtr(test4), ObjPtr(test2)
''    vbaObjSetAddref test2, test3
''    vbaObjSetAddref test1, test4
''
''    '#########SWAP OBJECT MEMORY
''    vbaObjSet ByVal VarPtr(Ptr4), ObjPtr(Ptr3)
''    vbaObjSetAddref Ptr3, ObjPtr(Ptr1)
''    vbaObjSetAddref ByVal VarPtr(Ptr2), ObjPtr(Ptr1)
''    vbaObjSetAddref Ptr5, ObjPtr(Ptr1)
''    vbaObjSetAddref Ptr4, ObjPtr(Ptr6)
''    vbaObjSetAddref ByVal VarPtr(Ptr3), Ptr4
''    vbaObjSetAddref ByVal VarPtr(Ptr1), ObjPtr(Ptr6)
''    vbaObjSet Ptr5, Ptr4
''End Sub
'
'Public Sub SwapPointer(ByRef Ptr1 As Object, ByRef Ptr2 As Object)
'    Dim Ptr3 As Long
'    Dim Ptr4 As Long
'    vbaObjSet ByVal VarPtr(Ptr3), ObjPtr(Ptr1)
'    vbaObjSetAddref ByVal VarPtr(Ptr4), ObjPtr(Ptr2)
'    vbaObjSetAddref Ptr2, Ptr3
'    vbaObjSetAddref Ptr1, Ptr4
'End Sub
'
'Public Sub NewPointer(ByRef Internal As Object, ByRef NewSpec As Object)
'    Dim test3 As Long
'    Dim test4 As Long
'    Dim test6 As Object
'    Dim External As Object
'    vbaObjSet ByVal VarPtr(test3), ObjPtr(NewSpec)
'    vbaObjSetAddref NewSpec, ObjPtr(Internal)
'    vbaObjSetAddref ByVal VarPtr(Internal), ObjPtr(External)
'    vbaObjSetAddref test4, ObjPtr(Internal)
'    vbaObjSetAddref test3, ObjPtr(test6)
'    vbaObjSetAddref ByVal VarPtr(Internal), ObjPtr(test6)
'    vbaObjSet test4, test3
'End Sub

Public Sub Main()


'  Set MyNodeList = HyperLists.CreateList
'
'
'  MyNodeList.Append
'  MyNodeList.Append
'  MyNodeList.Append
'  MyNodeList.Append
'
'
' MyNodeList.Forward
'
' MyNodeList.Forward
' MyNodeList.Forward
'   MyNodeList.Forward
'
'  MyNodeList.Remove
'  MyNodeList.Remove
'  MyNodeList.Remove
'  MyNodeList.Remove
''
'TestObject
'Testvalue
'TestRecord
''
'''   ' TestStrings
'''
'   MyNodeList.Clear
'''
'    End
'
'    MyNodeList.append
'    MyNodeList.append
'    MyNodeList.append
'    MyNodeList.append
'    MyNodeList.append
'
'    DebugNodes MyNodeList
'
'    MyNodeList.prior
'    MyNodeList.prior
'
'
'
'    DebugNodes MyNodeList
'
'
'    MyNodeList.Clear
  '  End


    Set MyNodeList = New NTNodes10.Nodes

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
            
            End If
            frmMain.Cls

            frmMain.Print "Insert/Delete to add or remove nodes at current location;"
            frmMain.Print "Up/Down to append to last of, or remove first of the list;"
            frmMain.Print "Right/Left to move the position of current point of list;"
            frmMain.Print "Green's current point, Blue's first Node, Red's Last node;"
            frmMain.Print "Home=Save, End=Clear, Return=Load; ";
            If MyNodeList.Exists Then
                frmMain.Print "Point(" & MyNodeList.Point & "); ";
            Else
                frmMain.Print "Point(nill); ";
            End If
            frmMain.Print
            frmMain.Print "Action: ";
            If Abs(TestAction) = 1 Then
                frmMain.Print "Append; ";
            ElseIf Abs(TestAction) = 2 Then
                frmMain.Print "Remove; ";
            ElseIf Abs(TestAction) = 5 Then
                frmMain.Print "Insert; ";
            ElseIf Abs(TestAction) = 6 Then
                frmMain.Print "Delete; ";
            Else
                frmMain.Print "None; ";
            End If
            frmMain.Print "Count: ";
            If (Not MyNodeList.Exists) Or (MyNodeList.Count < 1) Then
                frmMain.Print "None; ";
            Else
                frmMain.Print MyNodeList.Count & "; ";
            End If
            frmMain.Print "Motion: ";
            If Abs(TestAction) = 3 Then
                frmMain.Print "Backward; ";
            ElseIf Abs(TestAction) = 4 Then
                frmMain.Print "Forward; ";
            Else
                frmMain.Print "None; ";
            End If

            If latency <> printlag And latency <> 0 Then
                printlag = latency
            End If
            frmMain.Print "Latency: " & printlag

            If Abs(TestAction) = 1 Then
                MyNodeList.Append
                If lRGB = 1 Then
                    bDir = True
                End If
                lRGB = lRGB + IIf(bDir, 1, -1)
                If lRGB = 255 Then
                    bDir = False
                End If
                MyNodeList.Value = RGB(lRGB, lRGB, lRGB)
            ElseIf Abs(TestAction) = 2 Then
                MyNodeList.Remove
            ElseIf Abs(TestAction) = 3 Then
                MyNodeList.backward
            ElseIf Abs(TestAction) = 4 Then
                MyNodeList.Forward
            ElseIf Abs(TestAction) = 5 Then
                MyNodeList.Insert
                If lRGB = 1 Then
                    bDir = True
                End If
                lRGB = lRGB + IIf(bDir, 1, -1)
                If lRGB = 255 Then
                    bDir = False
                End If
                MyNodeList.Value = RGB(lRGB, lRGB, lRGB)
            ElseIf Abs(TestAction) = 6 Then
                MyNodeList.Delete
            End If

            If Pressed(VK_HOME) Then
                MyNodeList.save App.path & "\Nodes.bin"
            End If
            If Pressed(VK_END) Then
                MyNodeList.Clear
            End If
            If Pressed(VK_RETURN) Then
                MyNodeList.Load App.path & "\Nodes.bin"
            End If

            If Not MyNodeList.Exists Then
                frmMain.Print "First: @INVALID Final: @INVALID Check: @INVALID"
            Else
                frmMain.Print "First: " & MyNodeList.First & " Final: " & MyNodeList.Final & " Check: " & MyNodeList.Check
            End If

                frmMain.Print DebugNodes(MyNodeList, False)
                frmMain.DrawNodes MyNodeList

            
            If (Not Toggled(VK_RCONTROL)) Then
                TestAction = 0
            End If

            latency = CLng(GetTickCount - elapse)

        End If

        DoEvents

    Loop Until EndOfInput Or frmMain.Visible = False

    MyNodeList.Clear
    Set MyNodeList = Nothing
    Set HyperLists = Nothing
    Unload frmMain

End Sub

Public Function DebugNodes(ByVal N As inode, Optional ByVal FromPoint As Boolean = False) As String
    If N.Count = 0 Then
        DebugNodes = "@INVALID"
    Else
        Dim tot As Long
        
        Dim ptr As Long
        Dim eol As Long
        Dim bol As Long
        Dim chk As Long
        eol = N.Final
        bol = N.First
        ptr = N.Point
        chk = N.Check
    
        If FromPoint Then
            Do Until N.Point = bol
                N.Forward
            Loop
        End If
        
        Do
            DebugNodes = DebugNodes & " "
            If N.Point = ptr Then
                DebugNodes = DebugNodes & "^"
            End If
            If N.Point = chk Then
                DebugNodes = DebugNodes & "v"
            End If
            If N.Point = eol Then
                DebugNodes = DebugNodes & "E"
            End If
            If N.Point = bol Then
                DebugNodes = DebugNodes & "B"
            End If
            DebugNodes = DebugNodes & "@" & N.Point & " "
         N.Forward
            tot = tot + 1
            If tot > 400 Then Exit Do
        Loop Until (N.Point = ptr And (Not FromPoint)) Or (N.Point = bol And FromPoint)
    
        If FromPoint Then
            Do Until N.Point = ptr
                N.Forward
            Loop
        End If
        
        DebugNodes = Trim(DebugNodes)
    
    End If
    
End Function

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


