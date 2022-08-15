VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Serializer"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'This object attempts to Serialize any object passed to the Serialize function into XML
'and Deserialize it to an object again with the Deserialize function, optionally specifying a
'Library if know, it makes things easier for debugging.  It handles collections and nesting too.

Private Const Dent As String = "  "
Private Indent As String

Private Enum ImportantMembers
    Count = 1
    Key = 2
    Add = 4
    Item = 8
    Clear = 64
    SerialFlag = 128
    SerialObj = 256
End Enum


Private Function KeysExist(ByRef col) As Boolean
    On Error Resume Next
    Dim Key As String
    Key = col.Key(1)
    If Not Err Then
        KeysExist = True
    Else
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function CreateObjectEx(ByRef classname As String)
    On Error Resume Next
    Set CreateObjectEx = CreateObject(classname)
    If Err.Number <> 0 Then
        Err.Clear
        Set CreateObjectEx = modFactory.CreateObjectPrivate(classname)
        If Err.Number <> 0 Then
            Err.Clear
            Set CreateObjectEx = CreateObject(RemoveArg(classname, "."))
            If Err.Number <> 0 Then
                Err.Clear
                Set CreateObjectEx = modFactory.CreateObjectPrivate(RemoveArg(classname, "."))
            End If
        End If
    End If
    On Error GoTo 0
End Function

Private Function SerializeObjectEx(ByVal ObjToSerial, Optional ByVal member As String = "", Optional ByVal LibFilePath As String = "") As String
    Dim frm As New frmTemp
    Dim hlp As New Helper
    frm.ScriptControl1.AddObject "Serializer", Me, True
    frm.ScriptControl1.AddObject "ObjHelper", hlp, True
    frm.ScriptControl1.AddObject "ObjToSerial", ObjToSerial, True 'use the public method and just replace the name/xml tag defaults with nothing
    On Error Resume Next
    frm.ScriptControl1.ExecuteStatement "Set ObjHelper.Holder = ObjToSerial" & IIf(member = "", "", "." & member)
    If Err.Number = 0 Then
        SerializeObjectEx = Replace(Replace(Serialize(hlp.Holder, LibFilePath), "<?xml version=""1.0""?>" & vbCrLf & "<Serial>" & vbCrLf, ""), "</Serial>" & vbCrLf, "")
    Else
        SerializeObjectEx = ""
        Err.Clear
    End If
    On Error GoTo 0
    Unload frm
    Set frm = Nothing
End Function

Private Function IsRedundant(ByRef ObjToSerial, ByVal member As String) As Boolean
    Dim frm As New frmTemp
    Dim hlp As New Helper
    frm.ScriptControl1.AddObject "ObjToSerial", ObjToSerial
    frm.ScriptControl1.AddObject "ObjHelper", hlp
    On Error Resume Next
    IsRedundant = frm.ScriptControl1.Eval("(ObjHelper.ObjPtr(ObjToSerial)=ObjHelper.ObjPtr(ObjToSerial." & member & "))")
    If Err.Number <> 0 Then
        Err.Clear
        IsRedundant = True
    End If
    On Error GoTo 0
    Unload frm
    Set frm = Nothing
End Function

Private Function SerializeValueEx(ByRef ObjToSerial, Optional ByVal member As String = "") As String
    Dim frm As New frmTemp
    frm.ScriptControl1.AddObject "ObjToSerial", ObjToSerial
    SerializeValueEx = frm.ScriptControl1.Eval("Typename(ObjToSerial" & IIf(member = "", "", "." & member) & ")")
    SerializeValueEx = "<" & SerializeValueEx & ">" & URLEncode(frm.ScriptControl1.Eval("ObjToSerial" & IIf(member = "", "", "." & member))) & "</" & SerializeValueEx & ">"
    Unload frm
    Set frm = Nothing
End Function

Private Function IsObjectEx(ByRef ObjToSerial, ByVal member As String) As Boolean
    Dim frm As New frmTemp
    frm.ScriptControl1.AddObject "ObjToSerial", ObjToSerial
    IsObjectEx = frm.ScriptControl1.Eval("IsObject(ObjToSerial." & member & ")")
    Unload frm
    Set frm = Nothing
End Function

Private Function TypeNameEx(ByRef ObjToSerial, ByVal member As String) As String
    Dim frm As New frmTemp
    frm.ScriptControl1.AddObject "ObjToSerial", ObjToSerial, True
    TypeNameEx = frm.ScriptControl1.Eval("TypeName(ObjToSerial." & member & ")")
    Unload frm
    Set frm = Nothing
End Function

Private Sub PropertySetEx(ByRef ObjToSerial, Optional ByRef member As Variant, Optional ByRef SetToObject, Optional ByRef Key As Variant)
    Dim frm As New frmTemp
    frm.ScriptControl1.AddObject "ObjToSerial", ObjToSerial, True
    frm.ScriptControl1.AddObject "SetToObject", SetToObject, True
    If IsMissing(Key) And (Not IsMissing(member)) Then
        frm.ScriptControl1.ExecuteStatement "Set ObjToSerial." & member & " = SetToObject"
    ElseIf (Not IsMissing(Key)) And IsMissing(member) Then
        frm.ScriptControl1.ExecuteStatement "ObjToSerial.Add SetToObject" & Key
    End If
    Unload frm
    Set frm = Nothing
End Sub

Private Sub PropertyLetEx(ByVal ObjToSerial, Optional ByRef member As Variant, Optional ByVal SetToValue As String = "", Optional ByRef Key As Variant)
    Dim frm As New frmTemp
    frm.ScriptControl1.AddObject "ObjToSerial", ObjToSerial, True
    If IsMissing(Key) And (Not IsMissing(member)) Then
        Select Case frm.ScriptControl1.Eval("TypeName(ObjToSerial." & member & ")")
            Case "String"
                frm.ScriptControl1.ExecuteStatement "ObjToSerial." & member & " = " & SetToValue
            Case "Integer", "Long", "Single", "Byte", "Double"
                frm.ScriptControl1.ExecuteStatement "ObjToSerial." & member & " = " & SetToValue
        End Select
    ElseIf (Not IsMissing(Key)) And IsMissing(member) Then
        frm.ScriptControl1.ExecuteStatement "ObjToSerial.Add " & SetToValue & Key
    End If
    Unload frm
    Set frm = Nothing
End Sub

'Private Function IsCollection(ByVal lib As TypeLibInfo, ByVal classname As String) As Boolean
'    Dim cls As CoClasses
'    Dim fce As Interfaces
'    Dim mem As Members
'    Dim inf As MemberInfo
'    Dim cnt As Long
'    Dim cnt2 As Long
'    Dim cnt3 As Long
'    Dim hasItem As Boolean
'    Dim hasCount As Boolean
'    Dim hasAdd As Boolean
'    Set cls = lib.CoClasses
'    If InStr(classname, ".") > 0 Then classname = RemoveArg(classname, ".")
'    If cls.Count > 0 Then
'        For cnt = 1 To cls.Count
'            If Trim(LCase(classname)) = Trim(LCase(cls.Item(cnt).Name)) Then
'                Set fce = cls.Item(cnt).Interfaces
'                For cnt2 = 1 To fce.Count
'                    Set mem = fce.Item(cnt2).Members
'                    If mem.Count > 0 Then
'                        For cnt3 = 1 To mem.Count
'                            Set inf = mem.Item(cnt3)
'                            If LCase(inf.Name) = "count" Then hasCount = True
'                            If LCase(inf.Name) = "item" Then hasItem = True
'                            If LCase(inf.Name) = "add" Then hasAdd = True
'                        Next
'                    End If
'                Next
'            End If
'        Next
'    End If
'    IsCollection = (hasCount And hasItem And hasAdd)
'End Function

Private Function GetImportantMembers(ByVal lib As TypeLibInfo, ByVal classname As String) As ImportantMembers
    Dim cls As CoClasses
    Dim fce As Interfaces
    Dim mem As Members
    Dim inf As MemberInfo
    Dim cnt As Long
    Dim cnt2 As Long
    Dim cnt3 As Long
    Set cls = lib.CoClasses
    If InStr(classname, ".") > 0 Then classname = RemoveArg(classname, ".")
    If cls.Count > 0 Then
        For cnt = 1 To cls.Count
            If Trim(LCase(classname)) = Trim(LCase(cls.Item(cnt).Name)) Then
                Set fce = cls.Item(cnt).Interfaces
                For cnt2 = 1 To fce.Count
                    Set mem = fce.Item(cnt2).Members
                    If mem.Count > 0 Then
                        For cnt3 = 1 To mem.Count
                            Set inf = mem.Item(cnt3)
                            If LCase(inf.Name) = "count" Then GetImportantMembers = GetImportantMembers Or ImportantMembers.Count
                            If LCase(inf.Name) = "item" Then GetImportantMembers = GetImportantMembers Or ImportantMembers.Item
                            If LCase(inf.Name) = "add" Then GetImportantMembers = GetImportantMembers Or ImportantMembers.Add
                            If LCase(inf.Name) = "key" Then GetImportantMembers = GetImportantMembers Or ImportantMembers.Key
                            If LCase(inf.Name) = "clear" Then GetImportantMembers = GetImportantMembers Or ImportantMembers.Clear
                            If LCase(inf.Name) = "serialize" Then
                                Select Case inf.ReturnType
                                    Case TliVarType.VT_BOOL
                                        GetImportantMembers = GetImportantMembers Or ImportantMembers.SerialFlag
                                    Case Else
                                        GetImportantMembers = GetImportantMembers Or ImportantMembers.SerialObj
                                End Select
                            End If
                        Next
                    End If
                Next
            End If
        Next
    End If
End Function

Private Function IsCollection(ByVal ColMems As ImportantMembers) As Boolean
    IsCollection = (((ColMems And ImportantMembers.Count) = ImportantMembers.Count) And _
                    ((ColMems And ImportantMembers.Item) = ImportantMembers.Item))
End Function

Private Function GetLibraryObject(ByVal ObjToSerial, Optional ByVal LibFilePath As String = "") As TypeLibInfo
    If VBA.IsObject(ObjToSerial) Then
        Set GetLibraryObject = TLI.InterfaceInfoFromObject(ObjToSerial).Parent
    ElseIf VBA.TypeName(ObjToSerial) = "String" Then
        Dim seekName As String
        Dim findName As String
        seekName = ObjToSerial
        Dim col As New NTNodes10.Collection
        If IsCompiled Then
            col.Add modCommon.AppEXE(False, True)
            col.Add modCommon.AppEXE(False, False)
        Else
            col.Add "C:\Development\Neotext\Common\Projects\NTNodes10\Test\Project8.exe"
            col.Add "C:\Development\Neotext\Common\Binary\NTDirectX.ocx"
            col.Add "C:\Development\Neotext\Common\Binary\NTNodes10.dll"
        End If
        If LibFilePath <> "" Then col.Add LibFilePath
        If InStr(seekName, ".") > 0 Then
            findName = NextArg(seekName, ".")
            seekName = RemoveArg(seekName, ".")
        End If
        Dim cls As CoClasses
        Dim cnt As Long
        Do While (col.Count > 0)
            If PathExists(col.Item(1), True) Then
                Set GetLibraryObject = TypeLibInfoFromFile(col.Item(1))
                If Not (GetLibraryObject Is Nothing) Then
                    If LCase(Trim(GetLibraryObject.Name)) = LCase(Trim(findName)) Or Trim(findName) = "" Then
                        Set cls = GetLibraryObject.CoClasses
                        If cls.Count > 0 Then
                            For cnt = 1 To cls.Count
                                If Trim(LCase(seekName)) = Trim(LCase(cls.Item(cnt).Name)) Then
                                    col.Clear
                                    Set col = Nothing
                                    Exit Function
                                End If
                            Next
                        End If
                    End If
                End If
            End If
            col.Remove 1
        Loop
        Dim test As Object
        Set test = modFactory.CreateObjectPrivate(ObjToSerial)
        If Not test Is Nothing Then
            Set GetLibraryObject = TLI.InterfaceInfoFromObject(ObjToSerial).Parent
        Else
            Set GetLibraryObject = Nothing
        End If
    End If
End Function

Private Function Serialize2(ByRef ObjToSerial, Optional ByVal LibFilePath As String = "") As String
    Dim lib As TypeLibInfo
    Dim cls As CoClasses
    Dim fce As Interfaces
    Dim mem As Members
    Dim inf As MemberInfo
    Dim cnt As Long
    Dim cnt2 As Long
    Dim cnt3 As Long
    Dim gets As String
    Dim tmp As String
    Indent = Indent & Dent
    Serialize2 = Serialize2 & Indent & "<" & VBA.TypeName(ObjToSerial) & ">"
    Set lib = GetLibraryObject(ObjToSerial, LibFilePath)
    Dim colMem As ImportantMembers
    colMem = GetImportantMembers(lib, VBA.TypeName(ObjToSerial))
    If ((colMem And SerialObj) = SerialObj) Then
         Serialize2 = Serialize2 & URLEncode(ObjToSerial.Serialize)
    Else
        If ((colMem And SerialFlag) = SerialFlag) Then
            If Not ObjToSerial.Serialize Then
                cnt = -1 'skip it
            End If
        End If
        If cnt = 0 Then
            Serialize2 = Serialize2 & vbCrLf
            Set cls = lib.CoClasses
            If cls.Count > 0 Then
                For cnt = 1 To cls.Count
                    If Trim(LCase(VBA.TypeName(ObjToSerial))) = Trim(LCase(cls.Item(cnt).Name)) Then
                        Set fce = cls.Item(cnt).Interfaces
                        For cnt2 = 1 To fce.Count
                            Set mem = fce.Item(cnt2).Members
                            If mem.Count > 0 Then
                                For cnt3 = 1 To mem.Count
                                    Set inf = mem.Item(cnt3)
                                    If (inf.InvokeKind = INVOKE_PROPERTYGET) Then
                                        If inf.Parameters.Count = 0 Then
                                            If LCase(inf.Name) <> "serialize" Then
                                                Select Case inf.ReturnType 'gather gets to ensure we only serialize
                                                'those with put and get and in event of object, exclude circular ref
                                                    Case TliVarType.VT_BSTR, TliVarType.VT_I4, TliVarType.VT_R4, _
                                                        TliVarType.VT_UI1, TliVarType.VT_I2, TliVarType.VT_R8, TliVarType.VT_R8, TliVarType.VT_BOOL
                                                        gets = gets & inf.Name & ","
                                                    Case TliVarType.VT_DISPATCH, 0
                                                        If Not IsRedundant(ObjToSerial, inf.Name) Then
                                                            gets = gets & inf.Name & ","
                                                        End If
                                                    Case TliVarType.VT_VARIANT
                                                End Select
                                            End If
                                        End If
                                    End If
                                Next
                                Indent = Indent & Dent
                                For cnt3 = 1 To mem.Count
                                    Set inf = mem.Item(cnt3)
                                    If (inf.InvokeKind = INVOKE_PROPERTYPUT Or inf.InvokeKind = INVOKE_PROPERTYPUTREF) And (InStr(gets, inf.Name & ",") > 0) Then
                                        Select Case inf.ReturnType
                                            Case 24, 19 'not user defined so far
                                               ' Debug.Print inf.ReturnType
                                            Case TliVarType.VT_BOOL
                                                'Debug.Print "String " & inf.Name & " " & inf.InvokeKind
                                                Serialize2 = Serialize2 & Indent & "<" & inf.Name & ">" & SerializeValueEx(ObjToSerial, inf.Name) & "</" & inf.Name & ">" & vbCrLf
                                            Case TliVarType.VT_BSTR
                                                'Debug.Print "String " & inf.Name & " " & inf.InvokeKind
                                                Serialize2 = Serialize2 & Indent & "<" & inf.Name & ">" & SerializeValueEx(ObjToSerial, inf.Name) & "</" & inf.Name & ">" & vbCrLf
                                            Case TliVarType.VT_I4
                                                'Debug.Print "Long " & inf.Name & " " & inf.InvokeKind
                                                Serialize2 = Serialize2 & Indent & "<" & inf.Name & ">" & SerializeValueEx(ObjToSerial, inf.Name) & "</" & inf.Name & ">" & vbCrLf
                                            Case TliVarType.VT_R4
                                                'Debug.Print "Single " & inf.Name & " " & inf.InvokeKind
                                                Serialize2 = Serialize2 & Indent & "<" & inf.Name & ">" & SerializeValueEx(ObjToSerial, inf.Name) & "</" & inf.Name & ">" & vbCrLf
                                            Case TliVarType.VT_UI1
                                                'Debug.Print "Byte " & inf.Name & " " & inf.InvokeKind
                                                Serialize2 = Serialize2 & Indent & "<" & inf.Name & ">" & SerializeValueEx(ObjToSerial, inf.Name) & "</" & inf.Name & ">" & vbCrLf
                                            Case TliVarType.VT_I2
                                                'Debug.Print "Integer " & inf.Name & " " & inf.InvokeKind
                                                Serialize2 = Serialize2 & Indent & "<" & inf.Name & ">" & SerializeValueEx(ObjToSerial, inf.Name) & "</" & inf.Name & ">" & vbCrLf
                                            Case TliVarType.VT_R8
                                                'Debug.Print "Double " & inf.Name & " " & inf.InvokeKind
                                                Serialize2 = Serialize2 & Indent & "<" & inf.Name & ">" & SerializeValueEx(ObjToSerial, inf.Name) & "</" & inf.Name & ">" & vbCrLf
                                            Case TliVarType.VT_DISPATCH, 0
                                                'If (Not (LCase(inf.Name) = "forth")) Or ((LCase(inf.Name) = "forth") And (Not (LCase(TypeNameEx(ObjToSerial, inf.Name)) = LCase(VBA.TypeName(ObjToSerial))))) Then
                                                    tmp = SerializeObjectEx(ObjToSerial, inf.Name, LibFilePath)
                                                    If tmp <> "" Then Serialize2 = Serialize2 & Indent & "<" & inf.Name & ">" & vbCrLf & tmp & Indent & "</" & inf.Name & ">" & vbCrLf
                                                'End If
                                            Case TliVarType.VT_VARIANT
                                               ' Debug.Print "VT_VARIANT"
                                            Case Else
                                               ' Debug.Print inf.ReturnType
                                        End Select
                                    End If
                                Next
                                Indent = Mid(Indent, Len(Dent) + 1)
                            End If
                        Next
                    End If
                Next
            End If
            If IsCollection(colMem) Then
                If ObjToSerial.Count > 0 Then
                    Indent = Indent & Dent
                    For cnt = 1 To ObjToSerial.Count
                        Select Case VBA.TypeName(ObjToSerial.Item(cnt))
                            Case "String", "Long", "Single", "Double", "Byte", "Integer"
                                Serialize2 = Serialize2 & Indent & "<Item>" & vbCrLf
                                Indent = Indent & Dent
                                If ((colMem And ImportantMembers.Key) = ImportantMembers.Key) Then  'KeysExist(ObjToSerial) Then
                                    If ObjToSerial.Key(cnt) <> "" Then Serialize2 = Serialize2 & Indent & "<Key>" & URLEncode(ObjToSerial.Key(cnt)) & "</Key>" & vbCrLf
                                End If
                                Indent = Mid(Indent, Len(Dent) + 1)
                                Serialize2 = Serialize2 & Indent & "<" & VBA.TypeName(ObjToSerial.Item(cnt)) & ">" & URLEncode(ObjToSerial.Item(cnt)) & "</" & VBA.TypeName(ObjToSerial.Item(cnt)) & ">" & vbCrLf
                                Serialize2 = Serialize2 & Indent & "</Item>" & vbCrLf
                            Case Else
                                'Debug.Print VBA.TypeName(ObjToSerial.Item(cnt))
                                tmp = SerializeObjectEx(ObjToSerial.Item(cnt), , LibFilePath)
                                If Replace(Replace(tmp, "<" & VBA.TypeName(ObjToSerial.Item(cnt)) & "></" & VBA.TypeName(ObjToSerial.Item(cnt)) & ">" & vbCrLf, ""), " ", "") <> "" Then
                                    Serialize2 = Serialize2 & Indent & "<Item>" & vbCrLf
                                    Indent = Indent & Dent
                                
                                    If ((colMem And ImportantMembers.Key) = ImportantMembers.Key) Then  'KeysExist(ObjToSerial) Then
                                        If ObjToSerial.Key(cnt) <> "" Then Serialize2 = Serialize2 & Indent & "<Key>" & URLEncode(ObjToSerial.Key(cnt)) & "</Key>" & vbCrLf
                                    End If
                                    Indent = Mid(Indent, Len(Dent) + 1)
                                    Serialize2 = Serialize2 & tmp
                                    Serialize2 = Serialize2 & Indent & "</Item>" & vbCrLf
                                End If
                        End Select
                        
                    Next
                    Indent = Mid(Indent, Len(Dent) + 1)
                End If
            End If
            Serialize2 = Serialize2 & Indent
        End If
    End If
    Serialize2 = Serialize2 & "</" & VBA.TypeName(ObjToSerial) & ">" & vbCrLf
    Indent = Mid(Indent, Len(Dent) + 1)
End Function

Public Function Serialize(ByVal ObjToSerial, Optional ByVal LibFilePath As String = "") As String ' _
Converts ActiveX objects into text for read and write properties of the following data types; Boolean, Long, Integer, Double, Single, Byte, String, Object and Collection (explicit type library definition may be supplied by file path) If a boolean flag named Serialize of an objects properties is present, then it mustbe set true or it will be skipped. It Serialize is a string, then it will use it as the objects serial text.
Attribute Serialize.VB_Description = "Converts ActiveX objects into text for read and write properties of the following data types; Boolean, Long, Integer, Double, Single, Byte, String, Object and Collection (explicit type library definition may be supplied by file path) If a boolean flag named Serialize of an objects properties is present, then it mustbe set true or it will be skipped. It Serialize is a string, then it will use it as the objects serial text."
    Static objptrs As String 'returns an object xml serialized
    'format is name encloses type encloses properties recursive
    'special cases for collection and key collections with item
    If VBA.IsObject(ObjToSerial) And (Not InStr(objptrs, ObjPtr(ObjToSerial) & ",") > 0) And (Not (ObjToSerial Is Nothing)) Then
        objptrs = objptrs & ObjPtr(ObjToSerial) & "," 'avoid circular references
        Dim objText As String
        objText = Serialize2(ObjToSerial, LibFilePath)
        If objText <> "" Then
            Serialize = "<?xml version=""1.0""?>" & vbCrLf & "<Serial>" & vbCrLf
            Serialize = Serialize & objText
            Serialize = Serialize & "</Serial>" & vbCrLf
        End If
        objptrs = Replace(objptrs, ObjPtr(ObjToSerial) & ",", "")
    End If
End Function

Private Function Deserialize2(ByRef xml As MSXML.IXMLDOMNode, Optional ByVal LibFilePath As String = "")
    Dim lib As TypeLibInfo
    Dim Key As Variant
    Set lib = GetLibraryObject(xml.nodeName, LibFilePath)
    Dim colMem As ImportantMembers
    Dim child As MSXML.IXMLDOMNode
    Dim child2 As MSXML.IXMLDOMNode
    Set Deserialize2 = CreateObjectEx(lib.Name & "." & xml.nodeName)
    For Each child In xml.childNodes
        colMem = GetImportantMembers(lib, VBA.TypeName(Deserialize2))
        
        If (colMem And SerialObj) = SerialObj Then
            PropertyLetEx Deserialize2, "Serialize", """" & Replace(URLDecode(child.Text), """", """""") & """"
        ElseIf IsCollection(colMem) And (child.nodeName = "Item") Then

            Key = ""
            For Each child2 In child.childNodes
                If child2.nodeName = "Key" Then
                    Key = child2.Text
                    If Key <> "" Then
                        Key = ", """ & URLDecode(Key) & """"
                    End If
                Else
                    Select Case child2.nodeName
                        Case "String"
                            PropertyLetEx Deserialize2, , """" & Replace(URLDecode(child2.Text), """", """""") & """", Key
                        Case "Boolean"
                            PropertyLetEx Deserialize2, , "CBool(" & URLDecode(child2.Text) & ")", Key
                        Case "Integer"
                            PropertyLetEx Deserialize2, , "CInt(" & URLDecode(child2.Text) & ")", Key
                        Case "Long"
                            PropertyLetEx Deserialize2, , "CLng(" & URLDecode(child2.Text) & ")", Key
                        Case "Single"
                            PropertyLetEx Deserialize2, , "CSng(" & URLDecode(child2.Text) & ")", Key
                        Case "Byte"
                            PropertyLetEx Deserialize2, , "CByte(" & URLDecode(child2.Text) & ")", Key
                        Case "Double"
                            PropertyLetEx Deserialize2, , "CDbl(" & URLDecode(child2.Text) & ")", Key
                        Case Else
                            'Debug.Print child2.nodeName
                            PropertySetEx Deserialize2, , Deserialize2(child2, LibFilePath), Key
                    End Select
                End If
            Next

        Else
            For Each child2 In child.childNodes
                Select Case child2.nodeName
                    Case "String"
                        PropertyLetEx Deserialize2, child.nodeName, """" & Replace(URLDecode(child2.Text), """", """""") & """"
                    Case "Integer", "Long", "Single", "Byte", "Double", "Boolean"
                        PropertyLetEx Deserialize2, child.nodeName, URLDecode(child2.Text)
                    Case Else
                        PropertySetEx Deserialize2, child.nodeName, Deserialize2(child2, LibFilePath)
                End Select
            Next
        End If
    Next
End Function

Public Function Deserialize(ByVal ObjInSerial As String, Optional ByVal LibFilePath As String = "") ' _
Converts text retruned by the Serialize() function into program memory returning the ActiveX object (explicit type library definition may be supplied by file path) If a boolean flag named Serialize of an objects properties is present, then it mustbe set true or it will be skipped. It Serialize is a string, then it will use only the Serialize property as the object serial.
Attribute Deserialize.VB_Description = "Converts text retruned by the Serialize() function into program memory returning the ActiveX object (explicit type library definition may be supplied by file path) If a boolean flag named Serialize of an objects properties is present, then it mustbe set true or it will be skipped. It Serialize is a string, then it will use only the Serialize property as the object serial."
    Dim xml As New MSXML.DOMDocument
    Dim child As MSXML.IXMLDOMNode
    Dim child2 As MSXML.IXMLDOMNode
    xml.async = "false"
    xml.loadXML ObjInSerial
    Set Deserialize = Nothing
    If xml.parseError.errorCode = 0 Then
        For Each child In xml.childNodes
            If Deserialize Is Nothing And child.nodeName <> "xml" And child.nodeName <> "" Then
                If child.nodeName = "Serial" Then
                    For Each child2 In child.childNodes
                        If Deserialize Is Nothing And child2.nodeName <> "xml" And child2.nodeName <> "" Then
                            Set Deserialize = Deserialize2(child2, LibFilePath)
                        End If
                    Next
                ElseIf child.nodeName <> "Nothing" Then
                    Set Deserialize = Deserialize2(child, LibFilePath)
                End If
            End If
        Next
    Else
        Err.Raise 8, , "Invalid formatted."
    End If
End Function

Public Function URLDecode(ByVal encodedString As String) As String
    Dim ReturnString As String
    Dim currentChar As String
    Dim i As Long
    i = 1
    Do Until i > Len(encodedString)
        currentChar = Mid(encodedString, i, 1)
        If currentChar = "+" Then
            ReturnString = ReturnString + " "
            i = i + 1
        ElseIf currentChar = "%" Then
            currentChar = Mid(encodedString, i + 1, 2)
            ReturnString = ReturnString + Chr(Val("&H" & currentChar))
            i = i + 3
        Else
            ReturnString = ReturnString + currentChar
            i = i + 1
        End If
    Loop
    URLDecode = ReturnString
End Function

Public Function URLEncode(ByVal encodeString As String) As String
    Dim ReturnString As String
    Dim currentChar As String
    Dim i As Long
    For i = 1 To Len(encodeString)
        currentChar = Mid(encodeString, i, 1)
        If Asc(currentChar) < 91 And Asc(currentChar) > 64 Then
            ReturnString = ReturnString + currentChar
        ElseIf Asc(currentChar) < 123 And Asc(currentChar) > 96 Then
            ReturnString = ReturnString + currentChar
        ElseIf Asc(currentChar) < 58 And Asc(currentChar) > 47 Then
            ReturnString = ReturnString + currentChar
        ElseIf Asc(currentChar) = 32 Then
            ReturnString = ReturnString + "+"
        Else
            If Len(Hex(Asc(currentChar))) = 1 Then
                ReturnString = ReturnString + "%0" + Hex(Asc(currentChar))
            Else
                ReturnString = ReturnString + "%" + Hex(Asc(currentChar))
            End If
        End If
    Next
    URLEncode = ReturnString
End Function
