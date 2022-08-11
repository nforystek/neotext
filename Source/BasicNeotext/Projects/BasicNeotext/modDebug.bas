#Const [True] = -1
#Const [False] = 0

Attribute VB_Name = "modDebug"
Option Explicit

Public Sub DebugFilenames(vbp() As String)
    DebugPrint "Filenames"
    Dim cnt As Long
    For cnt = LBound(vbp) To UBound(vbp)
        DebugPrint vbTab & vbp(cnt)
    Next
End Sub
Public Sub DebugCodepane(vbp As CodePane)
    DebugPrint "Codepane"
    DebugPrint vbTab & vbp.Window.Caption
    DebugPrint vbTab & vbp.Window.Type
    vbp.Window.Visible = True
End Sub
Public Sub DebugCodeModules(vbp As CodeModule)
    DebugPrint "CodeModules"
    DebugPrint vbTab & vbp.Parent.Description
    DebugPrint vbTab & vbp.Parent.Name
    DebugCodepane vbp.CodePane
End Sub
Public Sub DebugProperties(vbp As Property)
    DebugPrint vbTab & vbp.Name
    DebugPrint vbTab & vbp.Value
End Sub
Public Sub DebugComponent(vbp As VBComponent)
    DebugPrint "Component"
    DebugPrint vbTab & vbp.Name
    
End Sub
Public Sub DebugReferences(vbp As Reference)
    DebugPrint "Reference"
    DebugPrint vbTab & vbp.BuiltIn
    DebugPrint vbTab & vbp.Description
    DebugPrint vbTab & vbp.FullPath
    DebugPrint vbTab & vbp.GUID
    DebugPrint vbTab & vbp.IsBroken
    DebugPrint vbTab & vbp.Major
    DebugPrint vbTab & vbp.Minor
    DebugPrint vbTab & vbp.Name
    DebugPrint vbTab & vbp.Type
End Sub
Public Sub DebugProject(vbp As VBProject)
    DebugPrint "Project"
    DebugPrint vbTab & vbp.BuildFileName
    DebugPrint vbTab & vbp.CompatibleOleServer
    DebugPrint vbTab & vbp.Description
    DebugPrint vbTab & vbp.FileName
    DebugPrint vbTab & vbp.HelpContextID
    DebugPrint vbTab & vbp.HelpFile
    DebugPrint vbTab & vbp.IconState
    DebugPrint vbTab & vbp.IsDirty
    DebugPrint vbTab & vbp.Name
    DebugPrint vbTab & vbp.Saved
    DebugPrint vbTab & vbp.StartMode
    DebugPrint vbTab & vbp.Type
    If Not vbp.VBComponents Is Nothing Then
        If vbp.VBComponents.Count > 0 Then
            Dim vbc As VBComponent
            For Each vbc In vbp.VBComponents
                DebugComponent vbc
            Next
        End If
    End If
    If Not vbp.References Is Nothing Then
        If vbp.References.Count > 0 Then
            Dim vbr As Reference
            For Each vbr In vbp.References
                DebugReferences vbr
            Next
        End If
    End If
End Sub
Public Sub DebugAddIn(Add As AddIn)
    DebugPrint "AddIn"
    DebugPrint vbTab & Add.Description
    DebugPrint vbTab & Add.GUID
    DebugPrint vbTab & Add.ProgId
End Sub

