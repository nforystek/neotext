VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "objEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public Function MessageQueueLog(ByVal MsgTo As String) As Long
    MessageQueueLog = modMsgQueue.MessageQueueLog(MsgTo)
End Function

Public Function MessageQueueGet(ByVal MsgTo As String) As Collection
    Set MessageQueueGet = modMsgQueue.MessageQueueGet(MsgTo)
End Function

Public Sub MessageQueueAdd(ByVal MsgTo As String, ByVal Msg As String)
    modMsgQueue.MessageQueueAdd MsgTo, Msg
End Sub

Public Function GetEvents() As Collection
    Set GetEvents = MaxEvents.GetEvents()
End Function

Public Sub AddEvent(ByVal Source As String, ByVal Location As String, ByVal Message As String)
    MaxEvents.AddEvent Nothing, Source, Location, Message
End Sub

Public Sub ClearEvents()
    MaxEvents.ClearEvents
End Sub

Attribute 