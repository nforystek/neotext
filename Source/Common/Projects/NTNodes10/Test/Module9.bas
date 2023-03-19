Attribute VB_Name = "Module1"
Option Explicit

Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Dest As Any, ByRef Source As Any, ByVal Length As Long)

Private Function PtrObj(ByRef lPtr As Long) As Object
    Dim lZero As Long
    Dim NewObj As Object
    RtlMoveMemory NewObj, lPtr, 4&
    Set PtrObj = NewObj
    RtlMoveMemory NewObj, lZero, 4&
End Function

Private Function PtrVar(ByRef lPtr As Long) As Variant
    Dim lZero As Long
    Dim NewObj As Variant
    RtlMoveMemory NewObj, ByVal lPtr, 4&
    RtlMoveMemory PtrVar, NewObj, 4&
   ' PtrVar = NewObj
'    RtlMoveMemory NewObj, lZero, 4&
End Function

Private Function PtrUnk(ByRef lPtr As Long) As IUnknown
    Dim lZero As Long
    Dim NewObj As IUnknown
    RtlMoveMemory PtrUnk, lPtr, 4&
'    PtrUnk = NewObj
'    RtlMoveMemory NewObj, lZero, 4&
End Function

Public Sub Main()

    Dim cls As New Class1
'    Dim Obj As Object
'    Dim unk As IUnknown
    cls.Var2 = "hello"
    
    
'    Set unk = cls
'    Set cls = Nothing
'    Set cls = unk
'    Set unk = Nothing
'
'
'
'
'    Set Obj = cls
'    Set cls = Nothing
'    Set unk = Obj
'    Set Obj = Nothing
'    Set cls = unk
'    Set unk = Nothing
'
'
'    Set Obj = Nothing
'    Set Obj = PtrObj(ObjPtr(cls))
'
'    Set unk = PtrUnk(ObjPtr(Obj))
'
'    Debug.Print Obj.Var2
'
'
'    Dim ptr As Variant
'    ptr = PtrVar(VarPtr(Obj))
'
'    Set Obj = Nothing
'    Set unk = Nothing
'
'
'
'   ' Debug.Print TypeName(PtrObj(ptr))
'
'
'    Debug.Print cls.Var2
'
'    End
'    Dim cls As New Class1

    cls.var8.sol2.Add "thius"

    Dim serial As String

    cls.Var1 = 83
    cls.Var2 = "some string"



    serial = Serialize(cls)

    Debug.Print serial

    Set cls = Deserialize(serial)

   Debug.Print cls.Var1

    Dim cls2 As Class2
    Dim col As NTNodes10.Collection

    Set cls2 = cls.var8

    Set col = cls2.sol2

   Debug.Print col.Count

    Debug.Print cls.Var2
    
    
    
End Sub
