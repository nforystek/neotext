#Const [True] = -1
#Const [False] = 0
Attribute VB_Name = "modGraphics"



#Const modGraphics = -1
Option Explicit
'TOP DOWN
Option Compare Binary
'Option Private Module
Private SkinVars As NTNodes10.Collection

Private Const GraphicsINI = "skin.ini"

Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As Any) As Long

Private Function RetrunUseFolder(ByVal Folder1 As String, ByVal Folder2 As String, ByVal FileName As String) As String
    RetrunUseFolder = IIf(Not PathExists(Folder1 + FileName), Folder2 + FileName, Folder1 + FileName)
End Function
Private Function SkinVarToResource(ByVal SkinVar As String) As Long
    Dim retVal As Long
    
    Select Case LCase(SkinVar)
    
        Case "backout"
            retVal = 101
        Case "back"
            retVal = 102
        Case "browseout"
            retVal = 103
        Case "browse"
            retVal = 104
        Case "closeout"
            retVal = 105
        Case "close"
            retVal = 106
        Case "copyout"
            retVal = 107
        Case "copy"
            retVal = 108
        Case "cutout"
            retVal = 109
        Case "cut"
            retVal = 110
        Case "deleteout"
            retVal = 111
        Case "delete"
            retVal = 112
        Case "forwardout"
            retVal = 113
        Case "forward"
            retVal = 114
        Case "goout"
            retVal = 115
        Case "go"
            retVal = 116
        Case "leftout"
            retVal = 117
        Case "left"
            retVal = 118
        Case "newfolderout"
            retVal = 119
        Case "newfolder"
            retVal = 120
        Case "pasteout"
            retVal = 121
        Case "paste"
            retVal = 122
        Case "refreshout"
            retVal = 123
        Case "refresh"
            retVal = 124
        Case "rightout"
            retVal = 125
        Case "right"
            retVal = 126
        Case "stopout"
            retVal = 127
        Case "stop"
            retVal = 128
        Case "uplevelout"
            retVal = 129
        Case "uplevel"
            retVal = 130
            
            

        Case "schedule_addout"
                    retVal = 131
        Case "schedule_add"
                    retVal = 132
        Case "schedule_deleteout"
                    retVal = 133
        Case "schedule_delete"
                    retVal = 134
        Case "schedule_downout"
                    retVal = 135
        Case "schedule_down"
                    retVal = 136
        Case "schedule_editout"
                    retVal = 137
        Case "schedule_edit"
                    retVal = 138
        Case "schedule_eventsout"
                    retVal = 139
        Case "schedule_events"
                    retVal = 140
        Case "schedule_loadout"
                    retVal = 141
        Case "schedule_load"
                    retVal = 142
        Case "schedule_runselectedout", "schedule_runallout", "schedule_runout"
                    retVal = 143
        Case "schedule_runselected", "schedule_runall", "schedule_run"
                    retVal = 144

        Case "schedule_saveout"
                    retVal = 145
        Case "schedule_save"
                    retVal = 146
        Case "schedule_servicestartout"
                    retVal = 147
        Case "schedule_servicestart"
                    retVal = 148
        Case "schedule_servicestopout"
                    retVal = 149
        Case "schedule_servicestop"
                    retVal = 150
        Case "schedule_stopout"
                    retVal = 151
        Case "schedule_stop"
                    retVal = 152
        Case "schedule_upout"
                    retVal = 153
        Case "schedule_up"
                    retVal = 154
                    
        Case "favorites_graphic"
            retVal = 155
                    
Case "script_addfileout"
                    retVal = 101
Case "script_addfile"
                    retVal = 102
Case "script_copyout"
                    retVal = 103
Case "script_copy"
                    retVal = 104
Case "script_cutout"
                    retVal = 105
Case "script_cut"
                    retVal = 106
Case "script_findout"
                    retVal = 107
Case "script_find"
                    retVal = 108
Case "script_newprojectout"
                    retVal = 109
Case "script_newproject"
                    retVal = 110
Case "script_openprojectout"
                    retVal = 111
Case "script_openproject"
                    retVal = 112
Case "script_pasteout"
                    retVal = 113
Case "script_paste"
                    retVal = 114
Case "script_redoout"
                    retVal = 115
Case "script_redo"
                    retVal = 116
Case "script_removefileout"
                    retVal = 117
Case "script_removefile"
                    retVal = 118
Case "script_runout"
                    retVal = 119
Case "script_run"
                    retVal = 120
Case "script_saveprojectout"
                    retVal = 121
Case "script_saveproject"
                    retVal = 122
Case "script_stopout"
                    retVal = 123
Case "script_stop"
                    retVal = 124
Case "script_undoout"
                    retVal = 125
Case "script_undo"
                    retVal = 126


    End Select

    SkinVarToResource = retVal
    
End Function
Private Function GetDefaultSkinValue(ByVal SkinVar As String) As Variant

    Dim retVal As Variant
    Dim skinDir As String
    Dim noneDir As String
    skinDir = AppPath & GraphicsFolder & dbSettings.GetProfileSetting("GraphicsFolder") & "\"
    noneDir = AppPath & GraphicsFolder & "(None)\"
    
    Select Case LCase(SkinVar)
    
        Case "uplevel"
            retVal = RetrunUseFolder(skinDir, noneDir, "client\uplevel_over.bmp")
        Case "back"
            retVal = RetrunUseFolder(skinDir, noneDir, "client\back_over.bmp")
        Case "forward"
            retVal = RetrunUseFolder(skinDir, noneDir, "client\forward_over.bmp")
        Case "stop"
            retVal = RetrunUseFolder(skinDir, noneDir, "client\stop_over.bmp")
        Case "refresh"
            retVal = RetrunUseFolder(skinDir, noneDir, "client\refresh_over.bmp")
        Case "newfolder"
            retVal = RetrunUseFolder(skinDir, noneDir, "client\newfolder_over.bmp")
        Case "delete"
            retVal = RetrunUseFolder(skinDir, noneDir, "client\delete_over.bmp")
        Case "cut"
            retVal = RetrunUseFolder(skinDir, noneDir, "client\cut_over.bmp")
        Case "copy"
            retVal = RetrunUseFolder(skinDir, noneDir, "client\copy_over.bmp")
        Case "paste"
            retVal = RetrunUseFolder(skinDir, noneDir, "client\paste_over.bmp")
        Case "go"
            retVal = RetrunUseFolder(skinDir, noneDir, "client\go_over.bmp")
        Case "close"
            retVal = RetrunUseFolder(skinDir, noneDir, "client\close_over.bmp")
        Case "browse"
            retVal = RetrunUseFolder(skinDir, noneDir, "client\browse_over.bmp")
    
        Case "uplevelout"
            retVal = RetrunUseFolder(skinDir, noneDir, "client\upLevel_out.bmp")
        Case "backout"
            retVal = RetrunUseFolder(skinDir, noneDir, "client\back_out.bmp")
        Case "forwardout"
            retVal = RetrunUseFolder(skinDir, noneDir, "client\forward_out.bmp")
        Case "stopout"
            retVal = RetrunUseFolder(skinDir, noneDir, "client\stop_out.bmp")
        Case "refreshout"
            retVal = RetrunUseFolder(skinDir, noneDir, "client\refresh_out.bmp")
        Case "newfolderout"
            retVal = RetrunUseFolder(skinDir, noneDir, "client\newfolder_out.bmp")
        Case "deleteout"
            retVal = RetrunUseFolder(skinDir, noneDir, "client\delete_out.bmp")
        Case "cutout"
            retVal = RetrunUseFolder(skinDir, noneDir, "client\cut_out.bmp")
        Case "copyout"
            retVal = RetrunUseFolder(skinDir, noneDir, "client\copy_out.bmp")
        Case "pasteout"
            retVal = RetrunUseFolder(skinDir, noneDir, "client\paste_out.bmp")
        Case "goout"
            retVal = RetrunUseFolder(skinDir, noneDir, "client\go_out.bmp")
        Case "closeout"
            retVal = RetrunUseFolder(skinDir, noneDir, "client\close_out.bmp")
        Case "browseout"
            retVal = RetrunUseFolder(skinDir, noneDir, "client\browse_out.bmp")
    
        Case "schedule_open"
            retVal = RetrunUseFolder(skinDir, noneDir, "schedule\open_over.bmp")
        Case "schedule_save"
            retVal = RetrunUseFolder(skinDir, noneDir, "schedule\save_over.bmp")
        Case "schedule_add"
            retVal = RetrunUseFolder(skinDir, noneDir, "schedule\add_over.bmp")
        Case "schedule_edit"
            retVal = RetrunUseFolder(skinDir, noneDir, "schedule\edit_over.bmp")
        Case "schedule_delete"
            retVal = RetrunUseFolder(skinDir, noneDir, "schedule\delete_over.bmp")
        Case "schedule_up"
            retVal = RetrunUseFolder(skinDir, noneDir, "schedule\up_over.bmp")
        Case "schedule_down"
            retVal = RetrunUseFolder(skinDir, noneDir, "schedule\down_over.bmp")
        Case "schedule_run"
            retVal = RetrunUseFolder(skinDir, noneDir, "schedule\run_over.bmp")
        Case "schedule_stop"
            retVal = RetrunUseFolder(skinDir, noneDir, "schedule\stop_over.bmp")
        Case "schedule_events"
            retVal = RetrunUseFolder(skinDir, noneDir, "schedule\events_over.bmp")
        Case "schedule_servicestart"
            retVal = RetrunUseFolder(skinDir, noneDir, "schedule\servicestart_over.bmp")
        Case "schedule_servicestop"
            retVal = RetrunUseFolder(skinDir, noneDir, "schedule\servicestop_over.bmp")
            
        Case "schedule_openout"
            retVal = RetrunUseFolder(skinDir, noneDir, "schedule\open_out.bmp")
        Case "schedule_saveout"
            retVal = RetrunUseFolder(skinDir, noneDir, "schedule\save_out.bmp")
        Case "schedule_addout"
            retVal = RetrunUseFolder(skinDir, noneDir, "schedule\add_out.bmp")
        Case "schedule_editout"
            retVal = RetrunUseFolder(skinDir, noneDir, "schedule\edit_out.bmp")
        Case "schedule_deleteout"
            retVal = RetrunUseFolder(skinDir, noneDir, "schedule\delete_out.bmp")
        Case "schedule_upout"
            retVal = RetrunUseFolder(skinDir, noneDir, "schedule\up_out.bmp")
        Case "schedule_downout"
            retVal = RetrunUseFolder(skinDir, noneDir, "schedule\down_out.bmp")
        Case "schedule_runout"
            retVal = RetrunUseFolder(skinDir, noneDir, "schedule\run_out.bmp")
        Case "schedule_stopout"
            retVal = RetrunUseFolder(skinDir, noneDir, "schedule\stop_out.bmp")
        Case "schedule_eventsout"
            retVal = RetrunUseFolder(skinDir, noneDir, "schedule\events_out.bmp")
    
        Case "schedule_servicestartout"
            retVal = RetrunUseFolder(skinDir, noneDir, "schedule\servicestart_out.bmp")
        Case "schedule_servicestopout"
            retVal = RetrunUseFolder(skinDir, noneDir, "schedule\servicestop_out.bmp")
    
        Case "script_newproject"
            retVal = RetrunUseFolder(skinDir, noneDir, "script\new_over.bmp")
        Case "script_openproject"
            retVal = RetrunUseFolder(skinDir, noneDir, "script\open_over.bmp")
        Case "script_saveproject"
            retVal = RetrunUseFolder(skinDir, noneDir, "script\save_over.bmp")
        Case "script_addfile"
            retVal = RetrunUseFolder(skinDir, noneDir, "script\add_over.bmp")
        Case "script_removefile"
            retVal = RetrunUseFolder(skinDir, noneDir, "script\remove_over.bmp")
        Case "script_undo"
            retVal = RetrunUseFolder(skinDir, noneDir, "script\undo_over.bmp")
        Case "script_redo"
            retVal = RetrunUseFolder(skinDir, noneDir, "script\redo_over.bmp")
        Case "script_cut"
            retVal = RetrunUseFolder(skinDir, noneDir, "script\cut_over.bmp")
        Case "script_copy"
            retVal = RetrunUseFolder(skinDir, noneDir, "script\copy_over.bmp")
        Case "script_paste"
            retVal = RetrunUseFolder(skinDir, noneDir, "script\paste_over.bmp")
        Case "script_find"
            retVal = RetrunUseFolder(skinDir, noneDir, "script\find_over.bmp")
        Case "script_stop"
            retVal = RetrunUseFolder(skinDir, noneDir, "script\stop_over.bmp")
        Case "script_run"
            retVal = RetrunUseFolder(skinDir, noneDir, "script\run_over.bmp")
    
        Case "script_newprojectout"
            retVal = RetrunUseFolder(skinDir, noneDir, "script\new_out.bmp")
        Case "script_openprojectout"
            retVal = RetrunUseFolder(skinDir, noneDir, "script\open_out.bmp")
        Case "script_saveprojectout"
            retVal = RetrunUseFolder(skinDir, noneDir, "script\save_out.bmp")
        Case "script_addfileout"
            retVal = RetrunUseFolder(skinDir, noneDir, "script\add_out.bmp")
        Case "script_removefileout"
            retVal = RetrunUseFolder(skinDir, noneDir, "script\remove_out.bmp")
        Case "script_undoout"
            retVal = RetrunUseFolder(skinDir, noneDir, "script\undo_out.bmp")
        Case "script_redoout"
            retVal = RetrunUseFolder(skinDir, noneDir, "script\redo_out.bmp")
        Case "script_cutout"
            retVal = RetrunUseFolder(skinDir, noneDir, "script\cut_out.bmp")
        Case "script_copyout"
            retVal = RetrunUseFolder(skinDir, noneDir, "script\copy_out.bmp")
        Case "script_pasteout"
            retVal = RetrunUseFolder(skinDir, noneDir, "script\paste_out.bmp")
        Case "script_findout"
            retVal = RetrunUseFolder(skinDir, noneDir, "script\find_out.bmp")
        Case "script_stopout"
            retVal = RetrunUseFolder(skinDir, noneDir, "script\stop_out.bmp")
        Case "script_runout"
            retVal = RetrunUseFolder(skinDir, noneDir, "script\run_out.bmp")
    
        Case "script_toolover_transparentcolor"
            retVal = "FF00FF"
        Case "script_toolout_transparentcolor"
            retVal = "FF00FF"
    
        Case "favorites_graphic"
            retVal = RetrunUseFolder(skinDir, noneDir, "fav.bmp")
        Case "list_background_graphic"
            retVal = ""
        Case "list_background_resize"
            retVal = "auto"
            
        Case "toolover_transparentcolor"
            retVal = "FF00FF"
        Case "toolout_transparentcolor"
            retVal = "FF00FF"
            
        Case "schedule_toolover_transparentcolor"
            retVal = "FF00FF"
        Case "schedule_toolout_transparentcolor"
            retVal = "FF00FF"
            
        Case "logview_backcolor"
            retVal = SystemColorConstants.vbWindowBackground
                        
        Case "logview_textcolor"
            retVal = "7F7F7F"
        Case "logview_incommingcolor"
            retVal = "00FF00"
        Case "logview_outgoingcolor"
            retVal = "0000FF"
        Case "logview_errorcolor"
            retVal = "FF0000"
        Case "logview_highlightcolor"
            retVal = "000000"
    
        Case "list_transparentcolor"
            retVal = "FFFFFF"

        Case "list_backcolor"
            retVal = SystemColorConstants.vbWindowBackground
        Case "list_textcolor"
            retVal = SystemColorConstants.vbWindowText
            
        Case "address_backcolor"
            retVal = SystemColorConstants.vbWindowBackground
        Case "address_textcolor"
            retVal = SystemColorConstants.vbWindowText
    
        Case "drivelist_backcolor"
            retVal = SystemColorConstants.vbWindowBackground
        Case "drivelist_textcolor"
            retVal = SystemColorConstants.vbWindowText
    
        Case "transferlist_backcolor"
            retVal = SystemColorConstants.vbWindowBackground
        Case "transferlist_textcolor"
            retVal = SystemColorConstants.vbWindowText
        Case "icecue_shadow"
            retVal = SystemColorConstants.vb3DShadow
        Case "icecue_hilite"
            retVal = SystemColorConstants.vb3DHighlight
        Case "sizers_normal"
            retVal = &H8000000F
        Case "sizers_moving"
            retVal = &H808080
        
        Case "script_commentcolor"
            retVal = "008000"
        Case "script_errorcolor"
            retVal = "FF0000"
        Case "script_operatorcolor"
            retVal = "404040"
        Case "script_statementcolor"
            retVal = "000080"
        Case "script_textcolor"
            retVal = "000000"
        Case "script_userdefinedcolor"
            retVal = "000000"
        Case "script_valuecolor"
            retVal = "808080"

        Case "toolbarbutton_spacer"
            retVal = "default"
        Case "toolbarbutton_width"
            retVal = 24
        Case "toolbarbutton_height"
            retVal = 24
        

    End Select
    
    GetDefaultSkinValue = retVal

End Function

Public Function SetIcecue(ByRef Line As Control, ByVal SkinVar As String)
    If LCase(GetCollectSkinValue(SkinVar)) = "none" Then
        Line.Visible = False
    Else
        Line.BorderColor = GetSkinColor(SkinVar)
    End If
End Function

Public Function GetSkinDimension(ByVal SkinVar As String) As Single

    GetSkinDimension = GetCollectSkinValue(SkinVar)

End Function

Public Function GetSkinColor(ByVal SkinVar As String) As Long

    GetSkinColor = GetColor(GetCollectSkinValue(SkinVar))

End Function

Public Function GetSkinColorHTML(ByVal SkinVar As String) As String

    GetSkinColorHTML = GetCollectSkinValue(SkinVar)

End Function

Public Function GetCollectSkinValue(ByVal SkinVar As String) As Variant
    If SkinVars Is Nothing Then
        Set SkinVars = New NTNodes10.Collection
        
        Dim ini As String
        Dim fna As String
        fna = AppPath + GraphicsFolder + "(None)\" + GraphicsINI
        
        Dim ID As String
        
        If PathExists(fna) Then
            
            ini = ReadFile(fna)
            
        Else
            ini = StrConv(LoadResData(100, "INI"), vbUnicode)
            
        End If
        
        Do While Not ini = ""
            ID = RemoveNextArg(ini, vbCrLf)
            If (Not (ID = "")) And (Not (Left(Trim(ID), 1) = ";")) Then
                If (Not (Left(ID, 1) = "[")) And (InStr(ID, "=") > 0) Then
                    fna = RemoveNextArg(ID, "=")
                    If SkinVars.Exists(Trim(LCase(fna))) Then SkinVars.Remove fna
                    SkinVars.Add ID, Trim(LCase(fna))
                End If
            End If
        Loop
    
        fna = AppPath + GraphicsFolder + dbSettings.GetProfileSetting("GraphicsFolder") + "\" + GraphicsINI

        If PathExists(fna) Then
            
            ini = ReadFile(fna)
            
            Do While Not ini = ""
                ID = RemoveNextArg(ini, vbCrLf)
                If (Not (ID = "")) And (Not (Left(Trim(ID), 1) = ";")) Then
                    If (Not (Left(ID, 1) = "[")) And (InStr(ID, "=") > 0) Then
                        fna = RemoveNextArg(ID, "=")
                        If SkinVars.Exists(Trim(LCase(fna))) Then SkinVars.Remove Trim(LCase(fna))
                        SkinVars.Add ID, Trim(LCase(fna))
                    End If
                End If
            Loop
            
        End If
    
    End If
    Dim ret As Variant
    
    If SkinVars.Exists(Trim(LCase(SkinVar))) Then
        If LCase(Trim(SkinVars.Item(Trim(LCase(SkinVar))))) = "default" Then
            ret = GetDefaultSkinValue(Trim(LCase(SkinVar)))
        Else
            ret = SkinVars.Item(Trim(LCase(SkinVar)))
        End If
    Else
        ret = GetDefaultSkinValue(Trim(LCase(SkinVar)))
    End If
        
    GetCollectSkinValue = ret
End Function
'Public Function ValueExists(ByRef col As Collection, ByVal Var As String) As Boolean
'    On Error Resume Next
'    Dim n As Variant
'    n = col.Item(Var)))
'    If Not Err.Number = 0 Then
'        Err.Clear
'        ValueExists = False
'    Else
'        ValueExists = True
'    End If
'    On Error GoTo 0
'End Function
Public Sub SetPicture(ByVal SkinVar As String, ByRef tControl)
    Dim fso As New Scripting.FileSystemObject

    Dim sSetting As String
    Dim sFile As String
    Dim sDefault As Boolean
    sDefault = True
    
    sSetting = GetCollectSkinValue(SkinVar)
    
    sFile = AppPath + GraphicsFolder + dbSettings.GetProfileSetting("GraphicsFolder") + "\" + sSetting
    
    If sSetting <> "" And sSetting <> "default" And sSetting <> "none" And fso.FileExists(sFile) Then
        tControl.Picture = LoadPicture(sFile)
        sDefault = False
    End If
    
    If sDefault And sSetting <> "none" Then
        
        sSetting = GetDefaultSkinValue(SkinVar)
        
        If sSetting <> "" And sSetting <> "none" And fso.FileExists(sSetting) Then
            tControl.Picture = LoadPicture(sSetting)
        End If
    End If
    
    If sSetting = "none" Then
        tControl.Picture = LoadPicture("")
    End If

    Set fso = Nothing
End Sub

Public Function GetColor(ByVal HTMLVal As Variant, Optional ByRef Red As Integer = -1, Optional ByRef Green As Integer = -1, Optional ByRef Blue As Integer = -1) As Variant

    If Red = -1 Then
        If TypeName(HTMLVal) = "Long" Then
            GetColor = HTMLVal
        ElseIf TypeName(HTMLVal) = "String" Then
            Dim r As Integer
            Dim g As Integer
            Dim b As Integer
            r = val("&H" & Left(HTMLVal, 2))
            g = val("&H" & Mid(HTMLVal, 3, 2))
            b = val("&H" & Right(HTMLVal, 2))
            GetColor = RGB(r, g, b)
        End If
    Else
        Red = val("&H" & Left(HTMLVal, 2))
        Green = val("&H" & Mid(HTMLVal, 3, 2))
        Blue = val("&H" & Right(HTMLVal, 2))
        GetColor = RGB(Red, Green, Blue)
    End If
        
End Function

Public Sub LoadButton(imgList, ByVal Key As String, ByVal iIndex As Integer)
    
    Dim ImgX As ListImage
    Dim fso As New Scripting.FileSystemObject
    Dim res As Long

    If fso.FileExists(AppPath + GraphicsFolder + dbSettings.GetProfileSetting("GraphicsFolder") + "\" + GetDefaultSkinValue(Key)) Then
        Set ImgX = imgList.ListImages.Add(iIndex, Key, LoadPicture(AppPath + GraphicsFolder + dbSettings.GetProfileSetting("GraphicsFolder") + "\" + GetDefaultSkinValue(Key)))
    Else
        res = SkinVarToResource(Key)
        If res <> 0 Then
            Set ImgX = imgList.ListImages.Add(iIndex, Key, PictureFromByteStream(LoadResData(res, "CUSTOM")))
        End If
    End If

    Set fso = Nothing
    
End Sub

Public Function PictureFromByteStream(b() As Byte) As Object 'IPicture
    Dim LowerBound As Long
    Dim ByteCount  As Long
    Dim hMem  As Long
    Dim lpMem  As Long
    Dim IID_IPicture(15)
    Dim istm As Object 'stdole.IUnknown

    On Error GoTo Err_Init
    If UBound(b, 1) < 0 Then
        Exit Function
    End If
    
    LowerBound = LBound(b)
    ByteCount = (UBound(b) - LowerBound) + 1
    hMem = GlobalAlloc(&H2, ByteCount)
    If hMem <> 0 Then
        lpMem = GlobalLock(hMem)
        If lpMem <> 0 Then
            RtlMoveMemory ByVal lpMem, b(LowerBound), ByteCount
            Call GlobalUnlock(hMem)
            If CreateStreamOnHGlobal(hMem, 1, istm) = 0 Then

                If CLSIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IPicture(0)) = 0 Then
                  Call OleLoadPicture(ByVal ObjPtr(istm), ByteCount, 0, IID_IPicture(0), PictureFromByteStream)
                End If
            End If
        End If
    End If
    
    Exit Function
    
Err_Init:
    If Err.Number = 9 Then
        'Uninitialized array
        MsgBox "Unable to load picture."
    Else
        MsgBox Err.Number & " - " & Err.Description
    End If
End Function



