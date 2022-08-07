VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Program"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 Option Explicit
'TOP DOWN

'###############################
'the ordering of public memory in
'this object class is very sensitive
'to the allocation ability to pointer
'call the functions in swapping them
'##################################

Private pHandle As Long
Private pAccepts As Long
Private pForth As Program
Private pPrior As Program
Private pMethods() As MethodInfo

Public Function Address()
    Debug.Print "2-";
    RtlMoveMemory ByVal pMethods(0).Pointer, pMethods(2).Address, 4&
    Address = Address()
End Function

Public Function Instruction()
    Debug.Print "1-";
    RtlMoveMemory ByVal pMethods(1).Pointer, pMethods(0).Address, 4&
    Instruction = Instruction()
End Function

Public Function Location()
    Debug.Print "3-";
    RtlMoveMemory ByVal pMethods(2).Pointer, pMethods(1).Address, 4&
    Location = 1
End Function
Public Function Routine(Optional ByVal i As Long) As Long
    RtlMoveMemory ByVal pMethods(1).Pointer, pMethods(Instruction).Address, 4&
    MessageQueue(i).Proc = HandleWindowProc(MessageQueue(i).Proc, MessageQueue(i).hwnd, _
        MessageQueue(i).uMsg, MessageQueue(i).wArg, MessageQueue(i).lArg, MessageQueue(i).addr)
    MovePaging pMethods
    RtlMoveMemory ByVal pMethods(2).Pointer, pMethods(2).Address, 4&
End Function

Public Function NewEnum() As IUnknown
    Set NewEnum = References.[_NewEnum]
End Function

Public Property Get Handle() As Long
    Handle = pHandle
End Property

Public Property Get Prior() As Program
    If pPrior Is Nothing Then
        Set Prior = Me
    Else
        Set Prior = pPrior
    End If
End Property
Friend Property Set Prior(ByRef newval As Program)
    Set pPrior = newval
End Property
Public Property Get Forth() As Program
    If pForth Is Nothing Then
        Set Forth = Me
    Else
        Set Forth = pForth
    End If
End Property
Friend Property Set Forth(ByRef newval As Program)
    Set pForth = newval
End Property

Private Sub Class_Initialize()
    Reference = Reference + 1
    pHandle = Reference
    pMethods = GetObjectFunctionsPointers(Me, 5, 0, 0)
    pAccepts = -1
End Sub

Private Sub Class_Terminate()
    Erase pMethods
End Sub

Public Property Get Accepts(Optional ByVal Flags As Long = 0)
    If (Flags > 0) Then
        Accepts = BitLong(pAccepts, Flags)
    ElseIf (Flags = 0) Then
        Accepts = pAccepts
    Else
        Accepts = pAccepts And Flags
    End If
End Property
Public Property Let Accepts(Optional ByVal Flags As Long = -1, ByVal newval)
    If (TypeName(newval) = "Long" And Flags = -1) Or TypeName(newval) = "Boolean" Then
        pAccepts = newval
    ElseIf (TypeName(newval) = "Boolean" Or (Flags > -1)) And (Flags > 0) Then
        BitLong(pAccepts, Flags) = newval
    ElseIf (Flags > -1) Then
        If TypeName(newval) = "Boolean" Then
            If newval Then
                pAccepts = pAccepts + Flags
            Else
                pAccepts = pAccepts - Flags
            End If
        Else
            pAccepts = Flags Xor newval
        End If
    Else
        pAccepts = Flags
    End If
End Property



