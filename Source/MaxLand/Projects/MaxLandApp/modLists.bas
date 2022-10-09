Attribute VB_Name = "modLists"
Option Explicit
'TOP DOWN

Option Private Module

Public Type SAFEARRAYBOUND ' 8 bytes
    cElements As Long
    lLbound   As Long
End Type

Public Type SAFEARRAYHEADER ' 20 bytes (for one dimensional arrays
    Dimensions    As Integer
    fFeatures     As Integer
    DataSize      As Long
    cLocks        As Long
    dataPointer   As Long
    sab(1)        As SAFEARRAYBOUND
End Type

Public Type MethodInfo
    Pointer As Long
    Address As Long
End Type

Public Reference As Long
Public Programs As Program
Public References As VBA.Collection

Public Property Get ProgramCount() As Long
    If Not References Is Nothing Then
        ProgramCount = References.Count
    Else
        ProgramCount = 0
    End If
End Property

Public Sub ClearPrograms()
    Do Until Programs Is Nothing
        DeleteProgram
    Loop
End Sub

Public Sub InsertProgram()
    If References Is Nothing Then Set References = New Collection
    
    Dim tmpObj As Program
    If Programs Is Nothing Then
        Set Programs = New Program
        Set tmpObj = Programs
    Else
        Set tmpObj = New Program
    End If
    References.Add ObjPtr(tmpObj), "h" & tmpObj.Handle
    Set tmpObj.Prior = Programs.Prior
    Set Programs.Prior.Forth = tmpObj
    Set tmpObj.Forth = Programs
    Set Programs.Prior = tmpObj
End Sub

Public Sub DeleteProgram()
    If Not Programs Is Nothing Then
        Dim tmpObj As Program
        References.Remove "h" & Programs.Handle
        If (Programs.Prior.Handle = Programs.Handle) Then
            Set Programs = Nothing
        Else
            Set Programs.Prior.Forth = Programs.Forth
            Set Programs.Forth.Prior = Programs.Prior
            Set Programs = Programs.Prior.Forth
        End If
    End If
End Sub

Public Static Sub ShiftPrograms()
    If Not Programs Is Nothing Then
        Set Programs = Programs.Forth
    End If
End Sub

Public Function RedimArray(ByVal DataSize As Long, ByVal lNumElements As Long, ByRef sa As SAFEARRAYHEADER, ByVal lDataPointer As Long, ByVal lArrayPointer As Long) As Long
    If lNumElements > 0 And lDataPointer <> 0 And lArrayPointer <> 0 Then
        With sa
            .DataSize = DataSize                              ' byte = 1 byte, integer = 2 bytes etc
            .Dimensions = 1 '2                                ' one dimensional
            .dataPointer = lDataPointer                       ' to unicode string data (or other?)
            .sab(0).lLbound = 0
            .sab(0).cElements = lNumElements
            RtlMoveMemory ByVal lArrayPointer, VarPtr(sa), 4& ' fake VB out
            RedimArray = True
        End With
    End If
End Function

Public Sub DestroyArray(ByVal lArrayPointer As Long)
    Dim lZero As Long
    RtlMoveMemory ByVal lArrayPointer, lZero, 4         ' put the array back to its original state
End Sub

Public Function GetObjectFunctionsPointers(obj As Object, ByVal NumberOfMethods As Long, Optional ByVal PublicVarNumber As Long, Optional ByVal PublicObjVariantNumber As Long) As MethodInfo()
    Dim FPS() As MethodInfo
    ReDim FPS(NumberOfMethods - 1)
    Dim obj1 As Long
    obj1 = ObjPtr(obj)
    Dim VTable As Long
    RtlMoveMemory VTable, ByVal obj1, 4
    Dim PTX As Long
    Dim cnt As Long
    For cnt = 0 To NumberOfMethods - 1
        PTX = VTable + 28 + (PublicVarNumber * 2 * 4) + (PublicObjVariantNumber * 3 * 4) + cnt * 4
        RtlMoveMemory FPS(cnt).Pointer, PTX, 4
        RtlMoveMemory FPS(cnt).Address, ByVal PTX, 4
    Next
    GetObjectFunctionsPointers = FPS
End Function

