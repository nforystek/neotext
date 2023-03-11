VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Billboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pKey As String

Private pTransparent As Boolean
Private pTranslucent As Boolean

Private pHeight As Single
Private pWidth As Single

Private pFileIndex() As Long

Private pSerialize As Boolean
Public Property Get Serialize() As Boolean
Attribute Serialize.VB_MemberFlags = "40"
    Serialize = pSerialize
End Property
Public Property Let Serialize(ByVal RHS As Boolean)
    pSerialize = RHS
End Property

Public Property Get Key() As String
Attribute Key.VB_UserMemId = 0
Attribute Key.VB_MemberFlags = "600"
    Key = pKey
End Property
Public Property Let Key(ByVal RHS As String)
    pKey = RHS
End Property

Public Property Get Transparent() As Boolean
    Transparent = pTransparent
End Property
Public Property Let Transparent(ByVal RHS As Boolean)
    pTransparent = RHS
End Property
Public Property Get Translucent() As Boolean
    Translucent = pTranslucent
End Property
Public Property Let Translucent(ByVal RHS As Boolean)
    pTranslucent = RHS
End Property

Public Property Get FileName(Optional ByVal ImageIndex As Long = 0) As String
    If ImageIndex >= LBound(pFileIndex) And ImageIndex <= UBound(pFileIndex) Then
        If (pFileIndex(ImageIndex) > 0) Then
            FileName = GetIndexFile(pFileIndex(ImageIndex))
        End If
    Else
        Err.Raise 8, , "ImageIndex does not exist."
    End If
End Property
Public Property Let FileName(Optional ByVal ImageIndex As Long = 0, ByVal RHS As String)
    If ImageIndex > UBound(pFileIndex) Then
        ReDim Preserve pFileIndex(0 To ImageIndex) As Long
    End If
    
    If RHS = "" And pFileIndex(ImageIndex) > 0 Then
        pFileIndex(ImageIndex) = 0
    ElseIf PathExists(RHS, True) Then
        pFileIndex(ImageIndex) = GetFileIndex(RHS)
        Set Files(pFileIndex(ImageIndex)).Data = LoadTexture(RHS)
        ImageDimensions RHS, Files(pFileIndex(ImageIndex)).Size
    End If
End Property

Public Property Get Width() As Single
    Width = pWidth
End Property
Public Property Let Width(ByVal RHS As Single)
    pWidth = RHS
End Property
Public Property Get Height() As Single
    Height = pHeight
End Property
Public Property Let Height(ByVal RHS As Single)
    pHeight = RHS
End Property

Private Sub Class_Initialize()
    ReDim pFileIndex(0 To 0) As Long
    Key = "K" & ObjPtr(Me)
End Sub

Private Sub Class_Terminate()
    If Key <> "" Then
        If Billboards.Exists(Key) Then Billboards.Remove Key
    End If
End Sub
