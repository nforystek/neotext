Attribute VB_Name = "modInit"
#Const modInit = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Sub InitMainGlobals()
    TreeFileExt = ".tree"
    ItemFileExt = ".item"
    
    EngineFolder = AppPath & "Engine\"
        MediaFolder = EngineFolder & "Media\"
        TemplateFolder = EngineFolder & "Templates\"
            BlankTreeFile = TemplateFolder & "Blank Tree" & TreeFileExt
            BlankItemFile = TemplateFolder & "Blank" & ItemFileExt
        
    MyTreeFolder = AppPath & "My Trees\"
        ExampleFolder = MyTreeFolder & "Examples\"
        ExportFolder = MyTreeFolder & "Export\"

    DefaultIni = "ViewPreviewPane = True" & vbCrLf & _
                "ViewGraphicalEdit = False" & vbCrLf & _
                "ForceDimensions = False" & vbCrLf & _
                "PromptForTemplate = True" & vbCrLf & _
                "IncludeCustom = False" & vbCrLf & _
                "IncludeCode = True" & vbCrLf & _
                "DialogDir=" & vbCrLf & _
                "DefaultDir=" & vbCrLf & _
                "Recent3=" & vbCrLf & _
                "Recent2=" & vbCrLf & _
                "Recent1=" & vbCrLf & _
                "Recent0=" & vbCrLf
End Sub

Public Sub InitFontFamily(ByRef cmbObject As ComboBox)
    With cmbObject
        .Clear
        .AddItem "Serif"
        .AddItem "Sans-Serif"
        .AddItem "Cursive"
        .AddItem "Fantasy"
        .AddItem "Monospace"
        .ListIndex = 0
    End With
End Sub
Public Sub InitFontColor(ByRef cmbObject As ComboBox)
    With cmbObject
        .Clear
        .AddItem "Black"
        .AddItem "White"
        .AddItem "Gray"
        .AddItem "Silver"
        .AddItem "Red"
        .AddItem "Green"
        .AddItem "Blue"
        .AddItem "Yellow"
        .AddItem "Purple"
        .AddItem "Olive"
        .AddItem "Navy"
        .AddItem "Aqua"
        .AddItem "Lime"
        .AddItem "Maroon"
        .AddItem "Teal"
        .AddItem "Fuchsia"
        .ListIndex = 0
    End With
End Sub

Public Sub InitFontSize(ByRef cmbObject As ComboBox)
    With cmbObject
        .Clear
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        .AddItem "+1"
        .AddItem "-1"
        .ListIndex = 1
    End With
End Sub

Public Sub InitItemHeight(ByRef cmbObject As ComboBox)
    With cmbObject
        .Clear
        .AddItem "16"
        .AddItem "20"
        .AddItem "24"
        .AddItem "28"
        .AddItem "32"
        .ListIndex = 1
    End With
End Sub
Public Sub InitFolderColor(ByVal SubFolder As String, ByRef cmbObject As ComboBox)
    
    With cmbObject
        cmbObject.Clear
    
        Dim nDir As String
        
        On Error Resume Next
        nDir = Dir(EngineFolder & "Media\" & SubFolder & "\", VbFileAttribute.vbDirectory)
        If Err Then Err.Clear
        On Error GoTo 0
        
        Do Until nDir = vbNullString
            If InStr(nDir, ".") = 0 Then
                cmbObject.AddItem FormalWord(nDir)
            End If
            nDir = Dir(, VbFileAttribute.vbDirectory)
        Loop
        cmbObject.ListIndex = 0
    End With
End Sub

