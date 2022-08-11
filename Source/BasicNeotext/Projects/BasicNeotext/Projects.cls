VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Projects"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN
Option Compare Binary


Private cCol As Collection

Public Property Get Count() As Long
    Count = cCol.Count
End Property

Public Property Get Item(ByVal cKey As Variant) As Project
    Set Item = cCol.Item(SafeKey(cKey))
End Property
Public Property Get Items() As Collection
    Set Items = cCol
End Property
Friend Function Exists(ByVal cKey As Variant) As Boolean
    If cCol.Count > 0 Then
        Dim Test As Project
        On Error Resume Next
        Set Test = cCol(SafeKey(cKey))
        If Err.Number <> 0 Then
            Err.Clear
            Exists = False
        Else
            Exists = True
        End If
    Else
        Exists = False
    End If
End Function

Friend Function Add(ByRef OBJ As Project) As Variant
    If OBJ.Key = "" Then
        Dim p As Project
        Dim cnt As Long
        For Each p In cCol
            If p.FileGroup = OBJ.FileGroup Then
                Set OBJ = p
                Exit Function
            ElseIf p.FileName = OBJ.FileName Then
                Set OBJ = p
                Exit Function
            End If
        Next
        Add = cCol.Count + 1
    Else
        Add = OBJ.Key
    End If
    If OBJ.Key = "" Then OBJ.Key = GetNewKey(cCol)
    cCol.Add OBJ, OBJ.Key
End Function

Friend Sub Remove(ByVal cKey As Variant)
    If IsNumeric(cKey) Then
        If val(cKey) > 0 And val(cKey) <= cCol.Count Then
            cCol.Remove cKey
        End If
    Else
        cCol.Remove SafeKey(cKey)
    End If
End Sub

Private Sub Class_Initialize()
    Set cCol = New Collection
End Sub

Private Sub Class_Terminate()
    ClearCollection cCol, True, True
    Set cCol = Nothing
End Sub

Public Function SafeKey(ByVal mKey As Variant) As Variant
    If IsNumeric(mKey) Then
        SafeKey = CLng(mKey)
    Else
        Dim ret As String
        ret = Trim(mKey)
        
        ret = Replace(ret, "/", CStr(Asc("/")))
        ret = Replace(ret, "\", CStr(Asc("\")))
        ret = Replace(ret, "|", CStr(Asc("|")))
        ret = Replace(ret, "[", CStr(Asc("[")))
        ret = Replace(ret, "]", CStr(Asc("]")))
        ret = Replace(ret, " ", CStr(Asc(" ")))
        ret = Replace(ret, "&", CStr(Asc("&")))
        ret = Replace(ret, "%", CStr(Asc("%")))
        ret = Replace(ret, "$", CStr(Asc("$")))
        ret = Replace(ret, ",", CStr(Asc(",")))
        ret = Replace(ret, ".", CStr(Asc(".")))
        ret = Replace(ret, "-", CStr(Asc("-")))
        ret = Replace(ret, "_", CStr(Asc("_")))
        ret = Replace(ret, "+", CStr(Asc("+")))
        ret = Replace(ret, "=", CStr(Asc("=")))
        ret = Replace(ret, "!", CStr(Asc("!")))
        ret = Replace(ret, "@", CStr(Asc("@")))
        ret = Replace(ret, "#", CStr(Asc("#")))
        ret = Replace(ret, "^", CStr(Asc("^")))
        ret = Replace(ret, "*", CStr(Asc("*")))
        ret = Replace(ret, "(", CStr(Asc(")")))
        ret = Replace(ret, "`", CStr(Asc("`")))
        ret = Replace(ret, "~", CStr(Asc("~")))
        ret = Replace(ret, "<", CStr(Asc("<")))
        ret = Replace(ret, ">", CStr(Asc(">")))
        ret = Replace(ret, "'", CStr(Asc("'")))
        ret = Replace(ret, """", CStr(Asc("""")))
        
        SafeKey = UCase(ret)
    End If
End Function

