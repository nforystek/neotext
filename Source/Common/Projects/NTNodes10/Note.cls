VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Note"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pHandle As Long
Private pForth As Note
Private pPrior As Note
Private pRefer As VBA.Collection

Friend Property Get Notes() As VBA.Collection
    Static NotesObj As VBA.Collection
    If NotesObj Is Nothing Then
        If pRefer Is Nothing Then
            Set NotesObj = New VBA.Collection
        Else
            Set NotesObj = pRefer
        End If
    ElseIf pRefer Is Nothing Then
        Set pRefer = NotesObj
    End If
    Set Notes = NotesObj
End Property
Friend Property Set Notes(ByRef newval As VBA.Collection)
    Set pRefer = newval
End Property
Private Sub ObjSwap(ByRef n1 As Note, ByRef n2 As Note)
    Dim Obj As Note
    Set Obj = n1
    Set n1 = n2
    Set n2 = Obj
End Sub

Public Property Get Handle() As Long
    Static iHandle As Long
    If Not (iHandle = 0) Then
        iHandle = iHandle + 1
        If pHandle = 0 Then pHandle = iHandle
    End If
    Handle = pHandle
End Property
Friend Property Let Handle(ByVal newval As Long)
    pHandle = newval
End Property

Public Property Get Prior() As Note
    If pPrior Is Nothing Then
        Set Prior = Me
    Else
        Set Prior = pPrior
    End If
End Property
Friend Property Set Prior(ByRef newval As Note)
    Set pPrior = newval
End Property
Public Property Get Forth() As Note
    If pForth Is Nothing Then
        Set Forth = Me
    Else
        Set Forth = pForth
    End If
End Property
Friend Property Set Forth(ByRef newval As Note)
    Set pForth = newval
End Property

Private Sub Class_Terminate()
    DeleteNote Me.Handle
End Sub

Private Sub Class_Initialize()
    Static iHandle As Long

    If Instances Is Nothing Then
        Set Instances = Me
    Else
        pHandle = iHandle
        iHandle = iHandle + 1
    End If

    Set Notes = Instances.Notes
End Sub

Private Sub InsertNote()
    Dim tmpObj As Note
    If Instances Is Nothing Then
        Set Instances = Me
        Set tmpObj = Instances
    Else
        Set tmpObj = New Note
        tmpObj.Handle = Handle
        tmpObj.Handle = Handle + 1
        Handle = Handle + 1
    End If
    tmpObj.Notes.Add tmpObj, "h" & tmpObj.Handle
    Set tmpObj.Prior = Instances.Prior
    Set Instances.Prior.Forth = tmpObj
    Set tmpObj.Forth = Instances
    Set Instances.Prior = tmpObj
End Sub

Private Sub DeleteNote(Optional ByVal Handle As Long = 0)
    If Not Instances Is Nothing Then
        If Handle > 0 Then
            Do Until Instances.Handle = Handle
                Set Instances = Instances.Forth
            Loop
        End If
        If (Instances.Prior.Handle = Instances.Handle) Then
            Set Instances = Nothing
        Else
            Set Instances.Prior.Forth = Instances.Forth
            Set Instances.Forth.Prior = Instances.Prior
            Set Instances = Instances.Prior.Forth
            If Handle = -1 Then DeleteNote -1
        End If
    End If
End Sub

Public Static Function Insert() As Object
    InsertNote
End Function

Public Static Sub Delete()
    Notes.Remove "h" & Handle
    DeleteNote Me.Handle
End Sub

Public Sub Clear()
    DeleteNote -1
End Sub

Public Function NewEnum() As IUnknown
    Set NewEnum = pRefer.[_NewEnum]
End Function




Private Sub Collection_Add(Anything As Variant, Optional Key As Variant, Optional Before As Variant, Optional After As Variant)
'
End Sub

Private Sub Collection_Clear()
'
End Sub

Private Function Collection_Clone() As Collection
'
End Function

Private Property Get Collection_Count() As Long
'
End Property

Private Property Get Collection_Exists(ByVal KeyOrIndex As Variant) As Boolean
'
End Property

Private Property Get Collection_Item(ByVal KeyOrIndex As Variant) As Variant
'
End Property

Private Property Let Collection_Item(ByVal KeyOrIndex As Variant, RHS As Variant)
'
End Property

Private Property Set Collection_Item(ByVal KeyOrIndex As Variant, RHS As Variant)
'
End Property

Private Property Get Collection_Key(ByVal Index As Variant) As Variant
'
End Property

Private Function Collection_NewEnum() As stdole.IUnknown
'
End Function

Private Sub Collection_Remove(Optional KeyOrIndex As Variant)
'
End Sub
