VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QueueCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Private fpList As NTNodes10.Collection
Private ftpURI As String

Public Property Get List() As NTNodes10.Collection
    Set List = fpList
End Property
Public Property Set List(ByVal RHS As NTNodes10.Collection)
    Set fpList = RHS
End Property

Public Property Get URI() As String
    URI = ftpURI
End Property
Public Property Let URI(ByVal RHS As String)
    ftpURI = RHS
End Property

