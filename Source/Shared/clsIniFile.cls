VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private nNames As Collection
Private nValues As Collection

Public Function Exists(ByVal sName As Variant) As Boolean
    If nNames.Count > 0 Then
        Exists = True
        Dim Test As Variant
        On Error Resume Next
        Test = nNames(sName)
        If Err Then
            Err.Clear
            Exists = False
        End If
        On Error GoTo 0
    Else
        Exists = False
    End If
End Function

Public Function Add(ByVal sName As String, Optional ByVal sValue As Variant = vbNullString)
    If Exists(sName) Then
        Remove sName
    End If
    nNames.Add sName, sName
    nValues.Add CStr(sValue), sName
End Function

Public Function Remove(ByVal sName As String)
    If Exists(sName) Then
        nNames.Remove sName
        nValues.Remove sName
    End If
End Function

Public Property Get Setting(ByVal sName As String) As Variant
    sName = LCase(Trim(sName))
    If Exists(sName) Then
        Select Case LCase(Trim(nValues(sName)))
            Case "true", "false"
                Setting = CBool(nValues(sName))
            Case Else
                Setting = CStr(nValues(sName))
        End Select
    Else
        Add sName
    End If
End Property
Public Property Let Setting(ByVal sName As String, ByVal sValue As Variant)
    sName = LCase(Trim(sName))
    Add sName, sValue
End Property
Public Property Let IniText(ByVal NewValue As String)
    Dim inCmd As String
    Dim inParam As String
    ClearCollection nNames, False, False
    ClearCollection nValues, False, False
    NewValue = Replace(NewValue, vbLf, vbNullString)
    Do Until NewValue = ""
        inParam = RemoveNextArg(NewValue, vbCr)
        inCmd = LCase(Trim(RemoveNextArg(inParam, "=")))
        If Not (Trim(inCmd) = "") Then
            Add inCmd, inParam
        End If
    Loop
End Property
Public Property Get IniText() As String
    Dim outText As String
    Dim outVar As Variant
    For Each outVar In nNames
        If Not (Trim(outVar) = "") Then
            outText = outText & outVar & "=" & nValues(outVar) & vbCrLf
        End If
    Next
    IniText = outText
End Property
Public Function LoadIniFile(ByVal FileName As String, Optional ByVal DefaultIniText As String = vbNullString, Optional ByVal Silent As Boolean = False) As Boolean
On Error GoTo Catch
    
    IniText = DefaultIniText & vbCrLf & ReadFile(FileName)

Catch:
    If Err Then
        If Not Silent Then
            MsgBox "Error reading INI file: " & Err.Description, vbCritical
        End If
        Err.Clear
        LoadIniFile = False
    Else
        LoadIniFile = True
    End If
    
On Error GoTo 0
End Function
Public Function SaveIniFile(ByVal FileName As String, Optional ByVal Silent As Boolean = False) As Boolean
On Error GoTo Catch
    
    WriteFile FileName, IniText
    
Catch:
    If Err Then
        If Not Silent Then
            MsgBox "Error writing INI file: " & Err.Description, vbCritical
        End If
        Err.Clear
        SaveIniFile = False
    Else
        SaveIniFile = True
    End If
    
On Error GoTo 0
End Function

Private Sub Class_Initialize()
    Set nNames = New Collection
    Set nValues = New Collection
End Sub

Private Sub Class_Terminate()
    ClearCollection nNames, False, True
    ClearCollection nValues, False, True
End Sub
