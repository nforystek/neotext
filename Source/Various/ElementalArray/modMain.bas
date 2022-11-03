Attribute VB_Name = "modMain"
Option Explicit
Option Base 0
Option Compare Binary
Option Private Module

#Const Locking = False

Public Const GPTR = &H40
Public Const GHND = &H42
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_FIXED = &H0
Public Const G32BIT_SIZE = (GHND - GPTR + GMEM_MOVEABLE)
Public Const GHILOW_SIZE = (GHND - GPTR + GMEM_FIXED)

Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub PutLong Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, num As Any, Optional ByVal Size As Long = G32BIT_SIZE)
Private Declare Sub GetLong Lib "kernel32" Alias "RtlMoveMemory" (num As Any, ByVal ptr As Long, Optional ByVal Size As Long = G32BIT_SIZE)
Private Declare Sub RtlMoveMemory Lib "kernel32" (nDest As Any, Source As Any, ByVal Length As Long)

'###########################################################################
'###########################################################################
'###########################################################################
'###########################################################################
'###########################################################################
    
Public Function ArrayInsert(ByRef Handle As Long) As Long ' _
    allocates a new element of the array at the beginning of it
Attribute ArrayInsert.VB_Description = "    allocates a new element of the array at the beginning of it"
    If Handle = 0 Then
        Handle = GlobalAlloc(GHND And VarPtr(Handle), G32BIT_SIZE)
    Else
        #If Locking Then
            Dim lHandle As Long
            lHandle = GlobalLock(Handle)
        #End If
        ArrayInsert = GlobalReAlloc(Handle, GlobalSize(Handle) + G32BIT_SIZE, GHND)
        RtlMoveMemory ByVal ArrayInsert + G32BIT_SIZE, ByVal ArrayInsert, (-GlobalSize(ArrayInsert) + -G32BIT_SIZE + (GlobalSize(ArrayInsert) * GHILOW_SIZE))
        #If Locking Then
            GlobalUnlock lHandle
        #End If
        If Handle <> ArrayInsert Then
            GlobalFree Handle
            Handle = ArrayInsert
        End If
    End If
    ArrayInsert = Handle
End Function
Public Function ArrayRemove(ByRef Handle As Long) As Long ' _
    re-moves from the array, disposing of the first element of it
Attribute ArrayRemove.VB_Description = "    re-moves from the array, disposing of the first element of it"
    If Handle > 0 Then
        If ArrayCount(Handle) > 1 Then
            #If Locking Then
                Dim lHandle As Long
                lHandle = GlobalLock(Handle)
            #End If
            RtlMoveMemory ByVal Handle, ByVal Handle + G32BIT_SIZE, (-GlobalSize(Handle) + -G32BIT_SIZE + (GlobalSize(Handle) * GHILOW_SIZE))
            ArrayRemove = GlobalReAlloc(Handle, GlobalSize(Handle) - G32BIT_SIZE, GHND)
            #If Locking Then
                GlobalUnlock lHandle
            #End If
            If Handle <> ArrayRemove Then
                GlobalFree Handle
                Handle = ArrayRemove
            End If
        Else
            GlobalFree Handle
            Handle = 0
        End If
    End If
End Function

Public Function ArrayAppend(ByRef Handle As Long) As Long ' _
    allocates a new element of the array at the end of it
Attribute ArrayAppend.VB_Description = "    allocates a new element of the array at the end of it"
    If Handle = 0 Then
        Handle = GlobalAlloc(GHND And VarPtr(Handle), G32BIT_SIZE)
    Else
        #If Locking Then
            Dim lHandle As Long
            lHandle = GlobalLock(Handle)
        #End If
        ArrayAppend = GlobalReAlloc(Handle, GlobalSize(Handle) + G32BIT_SIZE, GHND)
        #If Locking Then
            GlobalUnlock lHandle
        #End If
        If Handle <> ArrayAppend Then
            GlobalFree Handle
            Handle = ArrayAppend
        End If
    End If
    ArrayAppend = Handle + ((ArrayCount(Handle) - 1) * G32BIT_SIZE)
End Function
Public Function ArrayDelete(ByRef Handle As Long) As Long ' _
    depletes the array, disposing of the last element of it
Attribute ArrayDelete.VB_Description = "    depletes the array, disposing of the last element of it"
    If Handle > 0 Then
        If ArrayCount(Handle) > 1 Then
            #If Locking Then
                Dim lHandle As Long
                lHandle = GlobalLock(Handle)
            #End If
            ArrayDelete = GlobalReAlloc(Handle, GlobalSize(Handle) - G32BIT_SIZE, GHND)
            #If Locking Then
                GlobalUnlock lHandle
            #End If
            If Handle <> ArrayDelete Then
                GlobalFree Handle
                Handle = ArrayDelete
            End If
        Else
            GlobalFree Handle
            Handle = 0
        End If
    End If
End Function
Public Property Get ArrayIndex(ByVal Handle As Long, ByVal Addr As Long) As Long
    'gets the index of the element at addr in the array of handle
    ArrayIndex = ((Addr - Handle) \ G32BIT_SIZE)
End Property
Public Property Get ArrayAddr(ByVal Handle As Long, ByVal Index As Long) As Long
    'gets the addr of the element at index in the array of handle
    ArrayAddr = (Handle + (Index * G32BIT_SIZE))
End Property
Public Property Get ArrayItem(ByVal Addr As Long, Optional ByVal Index As Long = 0) As Long
    'gets the contents (long value) of the element at the optional index in the array/element at addr
    GetLong ArrayItem, (Addr + (Index * G32BIT_SIZE)), G32BIT_SIZE
End Property
Public Property Let ArrayItem(ByVal Addr As Long, Optional ByVal Index As Long = 0, ByVal RHS As Long)
    'sets the contents (long value) of the element at the optional index in the array/element at addr
    PutLong (Addr + (Index * G32BIT_SIZE)), RHS, G32BIT_SIZE
End Property

Public Property Get ArrayCount(ByVal Handle As Long)
    'gets the number of elements in the array of handle
    ArrayCount = (GlobalSize(Handle) \ G32BIT_SIZE)
End Property

'###########################################################################
'###########################################################################
'###########################################################################
'###########################################################################
'###########################################################################

Private Sub DebugArray(ByVal Arry As Long)
    Debug.Print
    Debug.Print "ArrayCount(" & Arry & ")=" & ArrayCount(Arry)
    If ArrayCount(Arry) > 0 Then
        Dim cnt As Long
        For cnt = 0 To ArrayCount(Arry) - 1
            Debug.Print "ArrayItem(" & ArrayAddr(Arry, cnt) & ")=" & ArrayItem(Arry, cnt)
        Next
    End If
    Debug.Print
End Sub

Private Sub Main()

    Dim Arry As Long 'the array

    Debug.Print ArrayAppend(Arry) & "=ArrayAppend(" & Arry & ")"
    ArrayItem(Arry, 0) = 0
    Debug.Print "ArrayItem(" & Arry & ", 0)=0"

    Debug.Print ArrayAppend(Arry) & "=ArrayAppend(" & Arry & ")"
    ArrayItem(Arry, 1) = 1
    Debug.Print "ArrayItem(" & Arry & ", 1)=1"


    Debug.Print ArrayAppend(Arry) & "=ArrayAppend(" & Arry & ")"
    ArrayItem(Arry, 2) = 2
    Debug.Print "ArrayItem(" & Arry & ", 2)=2"


    Debug.Print ArrayInsert(Arry) & "=ArrayInsert(" & Arry & ")"
    ArrayItem(Arry, 0) = 3
    Debug.Print "ArrayItem(" & Arry & ", 0)=3"

    DebugArray Arry

    Debug.Print ArrayRemove(Arry) & "=ArrayRemove(" & Arry & ")"

    Debug.Print ArrayDelete(Arry) & "=ArrayDelete(" & Arry & ")"

    Debug.Print ArrayDelete(Arry) & "=ArrayDelete(" & Arry & ")"

    Debug.Print ArrayDelete(Arry) & "=ArrayDelete(" & Arry & ")"

    DebugArray Arry

    Debug.Print CBool(Arry = 0) & "=CBool(Arry = 0)"

    Debug.Print


End Sub
