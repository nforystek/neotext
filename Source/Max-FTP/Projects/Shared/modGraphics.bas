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
    Dim retval As Long
    
    Select Case LCase(SkinVar)
    
        Case "backout"
            retval = 101
        Case "back"
            retval = 102
        Case "browseout"
            retval = 103
        Case "browse"
            retval = 104
        Case "closeout"
            retval = 105
        Case "close"
            retval = 106
        Case "copyout"
            retval = 107
        Case "copy"
            retval = 108
        Case "cutout"
            retval = 109
        Case "cut"
            retval = 110
        Case "deleteout"
            retval = 111
        Case "delete"
            retval = 112
        Case "forwardout"
            retval = 113
        Case "forward"
            retval = 114
        Case "goout"
            retval = 115
        Case "go"
            retval = 116
        Case "leftout"
            retval = 117
        Case "left"
            retval = 118
        Case "newfolderout"
            retval = 119
        Case "newfolder"
            retval = 120
        Case "pasteout"
            retval = 121
        Case "paste"
            retval = 122
        Case "refreshout"
            retval = 123
        Case "refresh"
            retval = 124
        Case "rightout"
            retval = 125
        Case "right"
            retval = 126
        Case "stopout"
            retval = 127
        Case "stop"
            retval = 128
        Case "uplevelout"
            retval = 129
        Case "uplevel"
            retval = 130
            
            

        Case "schedule_addout"
                    retval = 131
        Case "schedule_add"
                    retval = 132
        Case "schedule_deleteout"
                    retval = 133
        Case "schedule_delete"
                    retval = 134
        Case "schedule_downout"
                    retval = 135
        Case "schedule_down"
                    retval = 136
        Case "schedule_editout"
                    retval = 137
        Case "schedule_edit"
                    retval = 138
        Case "schedule_eventsout"
                    retval = 139
        Case "schedule_events"
                    retval = 140
        Case "schedule_loadout"
                    retval = 141
        Case "schedule_load"
                    retval = 142
        Case "schedule_runselectedout", "schedule_runallout", "schedule_runout"
                    retval = 143
        Case "schedule_runselected", "schedule_runall", "schedule_run"
                    retval = 144

        Case "schedule_saveout"
                    retval = 145
        Case "schedule_save"
                    retval = 146
        Case "schedule_servicestartout"
                    retval = 147
        Case "schedule_servicestart"
                    retval = 148
        Case "schedule_servicestopout"
                    retval = 149
        Case "schedule_servicestop"
                    retval = 150
        Case "schedule_stopout"
                    retval = 151
        Case "schedule_stop"
                    retval = 152
        Case "schedule_upout"
                    retval = 153
        Case "schedule_up"
                    retval = 154
                    
        Case "favorites_graphic"
            retval = 155
                    
Case "script_addfileout"
                    retval = 101
Case "script_addfile"
                    retval = 102
Case "script_copyout"
                    retval = 103
Case "script_copy"
                    retval = 104
Case "script_cutout"
                    retval = 105
Case "script_cut"
                    retval = 106
Case "script_findout"
                    retval = 107
Case "script_find"
                    retval = 108
Case "script_newprojectout"
                    retval = 109
Case "script_newproject"
                    retval = 110
Case "script_openprojectout"
                    retval = 111
Case "script_openproject"
                    retval = 112
Case "script_pasteout"
                    retval = 113
Case "script_paste"
                    retval = 114
Case "script_redoout"
                    retval = 115
Case "script_redo"
                    retval = 116
Case "script_removefileout"
                    retval = 117
Case "script_removefile"
                    retval = 118
Case "script_runout"
                    retval = 119
Case "script_run"
                    retval = 120
Case "script_saveprojectout"
                    retval = 121
Case "script_saveproject"
                    retval = 122
Case "script_stopout"
                    retval = 123
Case "script_stop"
                    retval = 124
Case "script_undoout"
                    retval = 125
Case "script_undo"
                    retval = 126


    End Select

    SkinVarToResource = retval
    
End Function
Private Function GetDefaultSkinValue(ByVal SkinVar As String) As Variant

    Dim retval As Variant
    Dim skinDir As String
    Dim noneDir As String
    skinDir = AppPath & GraphicsFolder & dbSettings.GetProfileSetting("GraphicsFolder") & "\"
    noneDir = AppPath & GraphicsFolder & "(None)\"
    
    Select Case LCase(SkinVar)
    
        Case "uplevel"
            retval = RetrunUseFolder(skinDir, noneDir, "client\uplevel_over.bmp")
        Case "back"
            retval = RetrunUseFolder(skinDir, noneDir, "client\back_over.bmp")
        Case "forward"
            retval = RetrunUseFolder(skinDir, noneDir, "client\forward_over.bmp")
        Case "stop"
            retval = RetrunUseFolder(skinDir, noneDir, "client\stop_over.bmp")
        Case "refresh"
            retval = RetrunUseFolder(skinDir, noneDir, "client\refresh_over.bmp")
        Case "newfolder"
            retval = RetrunUseFolder(skinDir, noneDir, "client\newfolder_over.bmp")
        Case "delete"
            retval = RetrunUseFolder(skinDir, noneDir, "client\delete_over.bmp")
        Case "cut"
            retval = RetrunUseFolder(skinDir, noneDir, "client\cut_over.bmp")
        Case "copy"
            retval = RetrunUseFolder(skinDir, noneDir, "client\copy_over.bmp")
        Case "paste"
            retval = RetrunUseFolder(skinDir, noneDir, "client\paste_over.bmp")
        Case "go"
            retval = RetrunUseFolder(skinDir, noneDir, "client\go_over.bmp")
        Case "close"
            retval = RetrunUseFolder(skinDir, noneDir, "client\close_over.bmp")
        Case "browse"
            retval = RetrunUseFolder(skinDir, noneDir, "client\browse_over.bmp")
    
        Case "uplevelout"
            retval = RetrunUseFolder(skinDir, noneDir, "client\upLevel_out.bmp")
        Case "backout"
            retval = RetrunUseFolder(skinDir, noneDir, "client\back_out.bmp")
        Case "forwardout"
            retval = RetrunUseFolder(skinDir, noneDir, "client\forward_out.bmp")
        Case "stopout"
            retval = RetrunUseFolder(skinDir, noneDir, "client\stop_out.bmp")
        Case "refreshout"
            retval = RetrunUseFolder(skinDir, noneDir, "client\refresh_out.bmp")
        Case "newfolderout"
            retval = RetrunUseFolder(skinDir, noneDir, "client\newfolder_out.bmp")
        Case "deleteout"
            retval = RetrunUseFolder(skinDir, noneDir, "client\delete_out.bmp")
        Case "cutout"
            retval = RetrunUseFolder(skinDir, noneDir, "client\cut_out.bmp")
        Case "copyout"
            retval = RetrunUseFolder(skinDir, noneDir, "client\copy_out.bmp")
        Case "pasteout"
            retval = RetrunUseFolder(skinDir, noneDir, "client\paste_out.bmp")
        Case "goout"
            retval = RetrunUseFolder(skinDir, noneDir, "client\go_out.bmp")
        Case "closeout"
            retval = RetrunUseFolder(skinDir, noneDir, "client\close_out.bmp")
        Case "browseout"
            retval = RetrunUseFolder(skinDir, noneDir, "client\browse_out.bmp")
    
        Case "schedule_open"
            retval = RetrunUseFolder(skinDir, noneDir, "schedule\open_over.bmp")
        Case "schedule_save"
            retval = RetrunUseFolder(skinDir, noneDir, "schedule\save_over.bmp")
        Case "schedule_add"
            retval = RetrunUseFolder(skinDir, noneDir, "schedule\add_over.bmp")
        Case "schedule_edit"
            retval = RetrunUseFolder(skinDir, noneDir, "schedule\edit_over.bmp")
        Case "schedule_delete"
            retval = RetrunUseFolder(skinDir, noneDir, "schedule\delete_over.bmp")
        Case "schedule_up"
            retval = RetrunUseFolder(skinDir, noneDir, "schedule\up_over.bmp")
        Case "schedule_down"
            retval = RetrunUseFolder(skinDir, noneDir, "schedule\down_over.bmp")
        Case "schedule_run"
            retval = RetrunUseFolder(skinDir, noneDir, "schedule\run_over.bmp")
        Case "schedule_stop"
            retval = RetrunUseFolder(skinDir, noneDir, "schedule\stop_over.bmp")
        Case "schedule_events"
            retval = RetrunUseFolder(skinDir, noneDir, "schedule\events_over.bmp")
        Case "schedule_servicestart"
            retval = RetrunUseFolder(skinDir, noneDir, "schedule\servicestart_over.bmp")
        Case "schedule_servicestop"
            retval = RetrunUseFolder(skinDir, noneDir, "schedule\servicestop_over.bmp")
            
        Case "schedule_openout"
            retval = RetrunUseFolder(skinDir, noneDir, "schedule\open_out.bmp")
        Case "schedule_saveout"
            retval = RetrunUseFolder(skinDir, noneDir, "schedule\save_out.bmp")
        Case "schedule_addout"
            retval = RetrunUseFolder(skinDir, noneDir, "schedule\add_out.bmp")
        Case "schedule_editout"
            retval = RetrunUseFolder(skinDir, noneDir, "schedule\edit_out.bmp")
        Case "schedule_deleteout"
            retval = RetrunUseFolder(skinDir, noneDir, "schedule\delete_out.bmp")
        Case "schedule_upout"
            retval = RetrunUseFolder(skinDir, noneDir, "schedule\up_out.bmp")
        Case "schedule_downout"
            retval = RetrunUseFolder(skinDir, noneDir, "schedule\down_out.bmp")
        Case "schedule_runout"
            retval = RetrunUseFolder(skinDir, noneDir, "schedule\run_out.bmp")
        Case "schedule_stopout"
            retval = RetrunUseFolder(skinDir, noneDir, "schedule\stop_out.bmp")
        Case "schedule_eventsout"
            retval = RetrunUseFolder(skinDir, noneDir, "schedule\events_out.bmp")
    
        Case "schedule_servicestartout"
            retval = RetrunUseFolder(skinDir, noneDir, "schedule\servicestart_out.bmp")
        Case "schedule_servicestopout"
            retval = RetrunUseFolder(skinDir, noneDir, "schedule\servicestop_out.bmp")
    
        Case "script_newproject"
            retval = RetrunUseFolder(skinDir, noneDir, "script\new_over.bmp")
        Case "script_openproject"
            retval = RetrunUseFolder(skinDir, noneDir, "script\open_over.bmp")
        Case "script_saveproject"
            retval = RetrunUseFolder(skinDir, noneDir, "script\save_over.bmp")
        Case "script_addfile"
            retval = RetrunUseFolder(skinDir, noneDir, "script\add_over.bmp")
        Case "script_removefile"
            retval = RetrunUseFolder(skinDir, noneDir, "script\remove_over.bmp")
        Case "script_undo"
            retval = RetrunUseFolder(skinDir, noneDir, "script\undo_over.bmp")
        Case "script_redo"
            retval = RetrunUseFolder(skinDir, noneDir, "script\redo_over.bmp")
        Case "script_cut"
            retval = RetrunUseFolder(skinDir, noneDir, "script\cut_over.bmp")
        Case "script_copy"
            retval = RetrunUseFolder(skinDir, noneDir, "script\copy_over.bmp")
        Case "script_paste"
            retval = RetrunUseFolder(skinDir, noneDir, "script\paste_over.bmp")
        Case "script_find"
            retval = RetrunUseFolder(skinDir, noneDir, "script\find_over.bmp")
        Case "script_stop"
            retval = RetrunUseFolder(skinDir, noneDir, "script\stop_over.bmp")
        Case "script_run"
            retval = RetrunUseFolder(skinDir, noneDir, "script\run_over.bmp")
    
        Case "script_newprojectout"
            retval = RetrunUseFolder(skinDir, noneDir, "script\new_out.bmp")
        Case "script_openprojectout"
            retval = RetrunUseFolder(skinDir, noneDir, "script\open_out.bmp")
        Case "script_saveprojectout"
            retval = RetrunUseFolder(skinDir, noneDir, "script\save_out.bmp")
        Case "script_addfileout"
            retval = RetrunUseFolder(skinDir, noneDir, "script\add_out.bmp")
        Case "script_removefileout"
            retval = RetrunUseFolder(skinDir, noneDir, "script\remove_out.bmp")
        Case "script_undoout"
            retval = RetrunUseFolder(skinDir, noneDir, "script\undo_out.bmp")
        Case "script_redoout"
            retval = RetrunUseFolder(skinDir, noneDir, "script\redo_out.bmp")
        Case "script_cutout"
            retval = RetrunUseFolder(skinDir, noneDir, "script\cut_out.bmp")
        Case "script_copyout"
            retval = RetrunUseFolder(skinDir, noneDir, "script\copy_out.bmp")
        Case "script_pasteout"
            retval = RetrunUseFolder(skinDir, noneDir, "script\paste_out.bmp")
        Case "script_findout"
            retval = RetrunUseFolder(skinDir, noneDir, "script\find_out.bmp")
        Case "script_stopout"
            retval = RetrunUseFolder(skinDir, noneDir, "script\stop_out.bmp")
        Case "script_runout"
            retval = RetrunUseFolder(skinDir, noneDir, "script\run_out.bmp")
    
        Case "script_toolover_transparentcolor"
            retval = "FF00FF"
        Case "script_toolout_transparentcolor"
            retval = "FF00FF"
    
        Case "favorites_graphic"
            retval = RetrunUseFolder(skinDir, noneDir, "fav.bmp")
        Case "list_background_graphic"
            retval = ""
        Case "list_background_resize"
            retval = "auto"
            
        Case "toolover_transparentcolor"
            retval = "FF00FF"
        Case "toolout_transparentcolor"
            retval = "FF00FF"
            
        Case "schedule_toolover_transparentcolor"
            retval = "FF00FF"
        Case "schedule_toolout_transparentcolor"
            retval = "FF00FF"
            
        Case "logview_backcolor"
            retval = SystemColorConstants.vbWindowBackground
                        
        Case "logview_textcolor"
            retval = "7F7F7F"
        Case "logview_incommingcolor"
            retval = "00FF00"
        Case "logview_outgoingcolor"
            retval = "0000FF"
        Case "logview_errorcolor"
            retval = "FF0000"
        Case "logview_highlightcolor"
            retval = "000000"
    
        Case "list_transparentcolor"
            retval = "FFFFFF"

        Case "list_backcolor"
            retval = SystemColorConstants.vbWindowBackground
        Case "list_textcolor"
            retval = SystemColorConstants.vbWindowText
            
        Case "address_backcolor"
            retval = SystemColorConstants.vbWindowBackground
        Case "address_textcolor"
            retval = SystemColorConstants.vbWindowText
    
        Case "drivelist_backcolor"
            retval = SystemColorConstants.vbWindowBackground
        Case "drivelist_textcolor"
            retval = SystemColorConstants.vbWindowText
    
        Case "transferlist_backcolor"
            retval = SystemColorConstants.vbWindowBackground
        Case "transferlist_textcolor"
            retval = SystemColorConstants.vbWindowText
        Case "icecue_shadow"
            retval = SystemColorConstants.vb3DShadow
        Case "icecue_hilite"
            retval = SystemColorConstants.vb3DHighlight
        Case "sizers_normal"
            retval = &H8000000F
        Case "sizers_moving"
            retval = &H808080
        
        Case "script_commentcolor"
            retval = "008000"
        Case "script_errorcolor"
            retval = "FF0000"
        Case "script_operatorcolor"
            retval = "404040"
        Case "script_statementcolor"
            retval = "000080"
        Case "script_textcolor"
            retval = "000000"
        Case "script_userdefinedcolor"
            retval = "000000"
        Case "script_valuecolor"
            retval = "808080"

        Case "toolbarbutton_spacer"
            retval = "default"
        Case "toolbarbutton_width"
            retval = 24
        Case "toolbarbutton_height"
            retval = 24
        

    End Select
    
    GetDefaultSkinValue = retval

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
            r = Val("&H" & Left(HTMLVal, 2))
            g = Val("&H" & Mid(HTMLVal, 3, 2))
            b = Val("&H" & Right(HTMLVal, 2))
            GetColor = RGB(r, g, b)
        End If
    Else
        Red = Val("&H" & Left(HTMLVal, 2))
        Green = Val("&H" & Mid(HTMLVal, 3, 2))
        Blue = Val("&H" & Right(HTMLVal, 2))
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



