VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Private pVar As Variant
Private pObj As Object
Private pRef As Boolean

Public Property Get obj() As Object
    Set obj = pObj
End Property

Public Property Get Var() As Variant
    Var = pVar
End Property

Public Sub AddObjByref(ByRef obj As Object)
    pRef = True
    vbaObjSet pObj, ObjPtr(obj)
    'Set pObj = Obj
End Sub

Public Sub AddObjByval(ByVal obj As Object)
    pRef = False
    vbaObjSetAddref pObj, ObjPtr(obj)
    'Set pObj = Obj
End Sub

Public Sub AddVarByref(ByRef Var As Variant)
    pRef = True
    'vbaObjSet pVar, VarPtr(Var)
    pVar = Var
End Sub

Public Sub AddVarByval(ByVal Var As Variant)
    pRef = False
    'vbaObjSetAddref pVar, VarPtr(Var)
    pVar = Var
End Sub

