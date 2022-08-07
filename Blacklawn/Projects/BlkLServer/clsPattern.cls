VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPattern"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary


Private pName As String

Private pPeek As Long
Private pList As String

Private pPass As String

Public Property Get Name() As String
    Name = pName
End Property
Public Property Let Name(ByVal newVal As String)
    pName = newVal
End Property

Public Property Get Peek() As Long
    Peek = pPeek
End Property
Public Property Let Peek(ByVal newVal As Long)
    pPeek = newVal
End Property
Public Property Get List() As String
    List = pList
End Property
Public Property Let List(ByVal newVal As String)
    pList = newVal
End Property

Public Property Get Pass() As String
    Pass = pPass
End Property
Public Property Let Pass(ByVal newVal As String)
    pPass = newVal
End Property

