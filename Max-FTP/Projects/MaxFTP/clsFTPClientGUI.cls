VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFTPClientGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private Const Border = 4

Public Function LoadGUI(ByRef myForm, ByVal Index As Integer)
    With myForm
   
        .NoResize = True
        
        .Image1(Index).Stretch = False
        SetPicture "list_background_graphic", .Image1(Index)
        
        .pView(Index).SortKey = dbSettings.GetClientSetting("wColumnKey" & Trim(Index))
        .pView(Index).SortOrder = dbSettings.GetClientSetting("wColumnSort" & Trim(Index))
        
        .vSizer.BackColor = GetSkinColor("sizers_normal")
        .hSizer.BackColor = GetSkinColor("sizers_normal")
    
        .pDummyView(Index).BackColor = GetSkinColor("logview_backcolor")
        
        .pView(Index).BackColor = GetSkinColor("list_backcolor")
        .pView(Index).ForeColor = GetSkinColor("list_textcolor")
        
        .Picture1(Index).BackColor = GetSkinColor("list_backcolor")
        
        .setLocation(Index).BackColor = GetSkinColor("address_backcolor")
        .setLocation(Index).ForeColor = GetSkinColor("address_textcolor")
        
        .pViewDrives(Index).BackColor = GetSkinColor("drivelist_backcolor")
        .pViewDrives(Index).ForeColor = GetSkinColor("drivelist_textcolor")
        
        .ListView1.BackColor = GetSkinColor("transferlist_backcolor")
        .ListView1.ForeColor = GetSkinColor("transferlist_textcolor")
       
        SetIcecue .Line1(0), "icecue_shadow"
        SetIcecue .Line1(2), "icecue_shadow"
        SetIcecue .Line5(0), "icecue_shadow"
        SetIcecue .Line4(0), "icecue_shadow"
        SetIcecue .Line7(0), "icecue_shadow"
        SetIcecue .Line1(1), "icecue_shadow"
        SetIcecue .Line5(1), "icecue_shadow"
        SetIcecue .Line4(1), "icecue_shadow"
        SetIcecue .Line7(1), "icecue_shadow"
        SetIcecue .Line7(2), "icecue_shadow"
        SetIcecue .Line4(2), "icecue_shadow"
        SetIcecue .Line5(2), "icecue_shadow"
        SetIcecue .Line5(3), "icecue_shadow"
        SetIcecue .Line1(3), "icecue_shadow"
        SetIcecue .Line4(3), "icecue_shadow"
        SetIcecue .Line7(3), "icecue_shadow"
        SetIcecue .Line1(4), "icecue_shadow"
        SetIcecue .Line7(4), "icecue_shadow"
        SetIcecue .Line4(4), "icecue_shadow"
        SetIcecue .Line5(4), "icecue_shadow"
        SetIcecue .Line1(5), "icecue_shadow"
        SetIcecue .Line5(5), "icecue_shadow"
        SetIcecue .Line4(5), "icecue_shadow"
        SetIcecue .Line7(5), "icecue_shadow"
                
        SetIcecue .Line6(0), "icecue_hilite"
        SetIcecue .Line2(0), "icecue_hilite"
        SetIcecue .Line3(0), "icecue_hilite"
        SetIcecue .Line8(0), "icecue_hilite"
        SetIcecue .Line2(0), "icecue_hilite"
        SetIcecue .Line2(1), "icecue_hilite"
        SetIcecue .Line6(1), "icecue_hilite"
        SetIcecue .Line3(1), "icecue_hilite"
        SetIcecue .Line8(1), "icecue_hilite"
        SetIcecue .Line2(2), "icecue_hilite"
        SetIcecue .Line8(2), "icecue_hilite"
        SetIcecue .Line6(5), "icecue_hilite"
        SetIcecue .Line3(2), "icecue_hilite"
        SetIcecue .Line6(2), "icecue_hilite"
        SetIcecue .Line2(3), "icecue_hilite"
        SetIcecue .Line6(3), "icecue_hilite"
        SetIcecue .Line3(3), "icecue_hilite"
        SetIcecue .Line8(3), "icecue_hilite"
        SetIcecue .Line2(4), "icecue_hilite"
        SetIcecue .Line3(4), "icecue_hilite"
        SetIcecue .Line6(4), "icecue_hilite"
        SetIcecue .Line8(4), "icecue_hilite"
        SetIcecue .Line2(5), "icecue_hilite"
        SetIcecue .Line3(5), "icecue_hilite"
        SetIcecue .Line8(5), "icecue_hilite"
        SetIcecue .Line6(5), "icecue_hilite"
        
        Set .ListView1.SmallIcons = frmMain.imgOperations(0)

        .UserControls(Index).Buttons.Clear
        Set .UserControls(Index).ImageList = frmMain.imgClient(0)
        Set .UserControls(Index).DisabledImageList = frmMain.imgClient(0)
        Set .UserControls(Index).HotImageList = frmMain.imgClient(1)
        
        .UserGo(Index).Buttons.Clear
        Set .UserGo(Index).ImageList = frmMain.imgClient16(0)
        Set .UserGo(Index).DisabledImageList = frmMain.imgClient16(0)
        Set .UserGo(Index).HotImageList = frmMain.imgClient16(1)
        
        Dim btnX As Button
        
        Set btnX = .UserControls(Index).Buttons.Add(1, "uplevel", , 0, "uplevelout")
        Set btnX = .UserControls(Index).Buttons.Add(2, "back", , 0, "backout")
        Set btnX = .UserControls(Index).Buttons.Add(3, "forward", , 0, "forwardout")
        Set btnX = .UserControls(Index).Buttons.Add(4, , , 3)
        If GetCollectSkinValue("toolbarbutton_spacer") = "none" Then btnX.Visible = False
        Set btnX = .UserControls(Index).Buttons.Add(5, "stop", , 0, "stopout")
        Set btnX = .UserControls(Index).Buttons.Add(6, "refresh", , 0, "refreshout")
        Set btnX = .UserControls(Index).Buttons.Add(7, , , 3)
        If GetCollectSkinValue("toolbarbutton_spacer") = "none" Then btnX.Visible = False
        Set btnX = .UserControls(Index).Buttons.Add(8, "newfolder", , 0, "newfolderout")
        Set btnX = .UserControls(Index).Buttons.Add(9, "delete", , 0, "deleteout")
        Set btnX = .UserControls(Index).Buttons.Add(10, , , 3)
        If GetCollectSkinValue("toolbarbutton_spacer") = "none" Then btnX.Visible = False
        Set btnX = .UserControls(Index).Buttons.Add(11, "cut", , 0, "cutout")
        Set btnX = .UserControls(Index).Buttons.Add(12, "copy", , 0, "copyout")
        Set btnX = .UserControls(Index).Buttons.Add(13, "paste", , 0, "pasteout")
        
        Set btnX = .UserGo(Index).Buttons.Add(1, "go", , 0, "goout")
        Set btnX = .UserGo(Index).Buttons.Add(2, "close", , 0, "closeout")
        Set btnX = .UserGo(Index).Buttons.Add(3, "browse", , 0, "browseout")
        
        .SetClientGUIState Index, st_StartUp
        
        .ClientGUILoaded(Index) = True
        
        .NoResize = False
        
        EnableTransferInfo myForm, dbSettings.GetProfileSetting("MultiThread")

    End With
End Function
Public Function UnloadGUI(ByRef myForm, ByVal Index As Integer)
    With myForm

        .ClearFileView Index
        
        .UserControls(Index).Buttons.Clear
        Set .UserControls(Index).ImageList = Nothing
        Set .UserControls(Index).DisabledImageList = Nothing
        Set .UserControls(Index).HotImageList = Nothing
        
        .UserGo(Index).Buttons.Clear
        Set .UserGo(Index).ImageList = Nothing
        Set .UserGo(Index).DisabledImageList = Nothing
        Set .UserGo(Index).HotImageList = Nothing
        
        
        Set .ListView1.SmallIcons = Nothing
        
        .ClientGUILoaded(Index) = False

    End With
    
End Function
Public Sub PaintClientGUI(ByRef myForm, ByVal Index As Integer)
    With myForm
    
        On Error Resume Next
        If .NoResize Then Exit Sub
        
        Dim xPixel As Integer
        Dim yPixel As Integer
        xPixel = (Screen.TwipsPerPixelX)
        yPixel = (Screen.TwipsPerPixelY)
                
        Dim cHeight As Integer
        Dim cTop As Integer
        Dim cWidth As Integer
        
        If .dContainer(GetContainerIndex(0, Index)).Visible Then
            .dContainer(GetContainerIndex(0, Index)).Move 0, 0, .userGUI(Index).Width, ((GetSkinDimension("toolbarbutton_height") + 6) * Screen.TwipsPerPixelY) + (yPixel * 4)
            .UserControls(Index).Move yPixel * 2, yPixel * 2, .dContainer(GetContainerIndex(0, Index)).Width - (yPixel * 4)
        End If
            
        If .dContainer(GetContainerIndex(1, Index)).Visible Then
            If .dContainer(GetContainerIndex(0, Index)).Visible Then
                .dContainer(GetContainerIndex(1, Index)).Top = .dContainer(GetContainerIndex(0, Index)).Height + (yPixel)
            Else
                .dContainer(GetContainerIndex(1, Index)).Top = 0
            End If
            .dContainer(GetContainerIndex(1, Index)).Left = 0
            .dContainer(GetContainerIndex(1, Index)).Width = .userGUI(Index).Width
            .dContainer(GetContainerIndex(1, Index)).Height = (22 * Screen.TwipsPerPixelY) + (yPixel * 4)
            .pViewDrives(Index).Move yPixel * 2, yPixel * 2, .dContainer(GetContainerIndex(1, Index)).Width - (yPixel * 4)
        End If
        
        If .dContainer(GetContainerIndex(2, Index)).Visible Then
            If .dContainer(GetContainerIndex(1, Index)).Visible Then
                .dContainer(GetContainerIndex(2, Index)).Top = .dContainer(GetContainerIndex(1, Index)).Top + .dContainer(GetContainerIndex(1, Index)).Height + (yPixel)
            Else
                If .dContainer(GetContainerIndex(0, Index)).Visible Then
                    .dContainer(GetContainerIndex(2, Index)).Top = .dContainer(GetContainerIndex(0, Index)).Height + (yPixel)
                Else
                    .dContainer(GetContainerIndex(2, Index)).Top = 0
                End If
            End If
            .dContainer(GetContainerIndex(2, Index)).Left = 0
            .dContainer(GetContainerIndex(2, Index)).Width = .userGUI(Index).Width
            .dContainer(GetContainerIndex(2, Index)).Height = (22 * Screen.TwipsPerPixelY) + (yPixel * 4)
            .pAddressBar(Index).Move yPixel * 2, yPixel * 2, .dContainer(GetContainerIndex(2, Index)).Width - (yPixel * 4)
            .UserGo(Index).Width = 47 * xPixel
            .UserGo(Index).Height = 36 * yPixel
            .UserGo(Index).Left = .dContainer(GetContainerIndex(2, Index)).Width - .UserGo(Index).Width - (yPixel * 4)
            .UserGo(Index).Top = 0
        End If
        
        cTop = yPixel * 2
        cHeight = .userGUI(Index).Height '- 20
        cWidth = .userGUI(Index).Width
        
        If .dContainer(GetContainerIndex(0, Index)).Visible Then
            cTop = cTop + .dContainer(GetContainerIndex(0, Index)).Height
            cHeight = cHeight - .dContainer(GetContainerIndex(0, Index)).Height
        End If
        If .dContainer(GetContainerIndex(1, Index)).Visible Then
            cTop = cTop + .dContainer(GetContainerIndex(1, Index)).Height
            cHeight = cHeight - .dContainer(GetContainerIndex(1, Index)).Height
        End If
        If .dContainer(GetContainerIndex(2, Index)).Visible Then
            cTop = cTop + .dContainer(GetContainerIndex(2, Index)).Height
            cHeight = cHeight - .dContainer(GetContainerIndex(2, Index)).Height
        End If
            
        If .mnuShowStatusBar.Checked Then
            .pStatus(Index).Visible = True
            .pProgress(Index).Visible = False
            cHeight = cHeight - (26 * Screen.TwipsPerPixelY)
            
            .pStatus(Index).Move 0, (.userGUI(Index).Height - .pStatus(Index).Height), .userGUI(Index).Width, (20 * Screen.TwipsPerPixelY)
            .pProgress(Index).Move 0, (.userGUI(Index).Height - .pProgress(Index).Height), .userGUI(Index).Width, (20 * Screen.TwipsPerPixelY)
            
        Else
            cHeight = cHeight - (2 * Screen.TwipsPerPixelY)
            .pStatus(Index).Visible = False
            .pProgress(Index).Visible = False
        End If
        
        .pView(Index).Move 0, cTop, cWidth, cHeight
        .Picture1(Index).Move 0, cTop, cWidth, cHeight
        SizeInnerImage myForm, CByte(Index)
        
        .pDummyView(Index).Move 0, 0, .Picture1(Index).ScaleWidth, .Picture1(Index).ScaleHeight
    
    End With
End Sub
Private Sub SizeInnerImage(ByRef myForm, ByRef Index As Byte)
    With myForm
        Select Case LCase(GetCollectSkinValue("list_background_resize"))
            Case "auto"
                .Image1(Index).Stretch = True
                If .Picture1(Index).ScaleWidth > 0 Then
                    .Image1(Index).Left = (.Picture1(Index).ScaleWidth / 2) - (.Image1(Index).Width / 2)
                End If
                If .Picture1(Index).ScaleHeight > 0 Then
                    .Image1(Index).Top = (.Picture1(Index).ScaleHeight / 2) - (.Image1(Index).Height / 2)
                End If
            Case "fit"
                .Image1(Index).Stretch = True
                
                .Image1(Index).Left = 0
                If .Picture1(Index).ScaleWidth > 0 Then
                    .Image1(Index).Width = .Picture1(Index).ScaleWidth
                End If
                .Image1(Index).Top = 0
                If .Picture1(Index).ScaleHeight > 0 Then
                    .Image1(Index).Height = .Picture1(Index).ScaleHeight
                End If
            
        End Select
    End With
End Sub
Public Sub PaintInfoGUI(ByRef myForm)
    With myForm

        On Error Resume Next
        
        If .NoResize Then Exit Sub
        
        .ListView1.Move 0, 0, .userInfo.Width, .userInfo.Height
        
        Err.Clear
    
    End With

End Sub

Public Function GetContainerIndex(ByVal ConNum As Integer, GUIIndex As Integer) As Integer

    Dim retVal As Integer
    Select Case ConNum
        Case 0
        Select Case GUIIndex
            Case 0
                retVal = 0
            Case 1
                retVal = 3
        End Select
        Case 1
        Select Case GUIIndex
            Case 0
                retVal = 1
            Case 1
                retVal = 4
        End Select
        Case 2
        Select Case GUIIndex
            Case 0
                retVal = 2
            Case 1
                retVal = 5
        End Select
    End Select
    GetContainerIndex = retVal
    
End Function

Public Sub dContainerResize(ByRef myForm, Index As Integer)
    With myForm

        On Error Resume Next
        
        If .NoResize Then Exit Sub
        
        Dim xPixel As Integer
        Dim yPixel As Integer
        xPixel = (Screen.TwipsPerPixelX)
        yPixel = (Screen.TwipsPerPixelY)
        
        .Line1(Index).X1 = 0
        .Line1(Index).X2 = 0
        .Line1(Index).Y1 = yPixel
        .Line1(Index).Y2 = (.dContainer(Index).Height - (2 * yPixel))
        
        .Line2(Index).X1 = xPixel
        .Line2(Index).X2 = xPixel
        .Line2(Index).Y1 = yPixel
        .Line2(Index).Y2 = (.dContainer(Index).Height - (2 * yPixel))
        
        .Line3(Index).X1 = .dContainer(Index).Width - yPixel
        .Line3(Index).X2 = .dContainer(Index).Width - yPixel
        .Line3(Index).Y1 = 0
        .Line3(Index).Y2 = .dContainer(Index).Height - yPixel
        
        .Line4(Index).X1 = .dContainer(Index).Width - (xPixel * 2)
        .Line4(Index).X2 = .dContainer(Index).Width - (xPixel * 2)
        .Line4(Index).Y1 = yPixel
        .Line4(Index).Y2 = (.dContainer(Index).Height - (2 * yPixel))
        
        .Line5(Index).X1 = 0
        .Line5(Index).X2 = .dContainer(Index).Width
        .Line5(Index).Y1 = 0
        .Line5(Index).Y2 = 0
        
        .Line6(Index).X1 = xPixel
        .Line6(Index).X2 = .dContainer(Index).Width
        .Line6(Index).Y1 = yPixel
        .Line6(Index).Y2 = yPixel
        
        .Line7(Index).X1 = xPixel
        .Line7(Index).X2 = .dContainer(Index).Width - (xPixel * 2)
        .Line7(Index).Y1 = .dContainer(Index).Height - (yPixel * 2)
        .Line7(Index).Y2 = .dContainer(Index).Height - (yPixel * 2)
        
        .Line8(Index).X1 = 0
        .Line8(Index).X2 = .dContainer(Index).Width
        .Line8(Index).Y1 = .dContainer(Index).Height - yPixel
        .Line8(Index).Y2 = .dContainer(Index).Height - yPixel
        
        Err.Clear

    End With
End Sub

Public Sub FormResize(ByRef myForm)
    With myForm

        On Error Resume Next
        
        If .NoResize Then Exit Sub
        
        If .mnuDoubleWin.Checked Then
            If Not .userGUI(0).Visible Then .userGUI(0).Visible = True
            If Not .userGUI(1).Visible Then .userGUI(1).Visible = True
        
            .vSizer.Visible = True
            If .mnuMultiThread.Checked Then
                If Not .userInfo.Visible = True Then .userInfo.Visible = True
                .hSizer.Visible = True
            
                If .ReCenterSizers Then
                    .hSizer.Top = (0.68 * .ScaleHeight)
                End If
            
                .hSizer.Left = (Border * Screen.TwipsPerPixelX)
                .hSizer.Width = .ScaleWidth - ((Border * Screen.TwipsPerPixelX) * 2)
                If .hSizer.Top < (Border * Screen.TwipsPerPixelY) * 2 Then .hSizer.Top = (Border * Screen.TwipsPerPixelY) * 2
                If .hSizer.Top + ((Border * Screen.TwipsPerPixelY) * 2) > .ScaleHeight Then .hSizer.Top = .ScaleHeight - ((Border * Screen.TwipsPerPixelY) * 2)
            
            Else
                If Not .userInfo.Visible = False Then .userInfo.Visible = False
                .hSizer.Visible = False
            End If
            
            If .ReCenterSizers Then
                .vSizer.Left = (.ScaleWidth / 2) - ((Border * Screen.TwipsPerPixelX) / 2)
            End If
            
            If .vSizer.Left < (Border * Screen.TwipsPerPixelX) * 2 Then .vSizer.Left = (Border * Screen.TwipsPerPixelX) * 2
            If .vSizer.Left + ((Border * Screen.TwipsPerPixelX) * 2) > .ScaleWidth Then .vSizer.Left = .ScaleWidth - ((Border * Screen.TwipsPerPixelX) * 2)
            
            .vSizer.Top = (Border * Screen.TwipsPerPixelY)
            .userInfo.Left = (Border * Screen.TwipsPerPixelX)
            .userInfo.Width = (.ScaleWidth) - ((Border * Screen.TwipsPerPixelX) * 2)
                        
            .userGUI(0).Move (Border * Screen.TwipsPerPixelX), 0, (.vSizer.Left) - ((Border * Screen.TwipsPerPixelX)) - Screen.TwipsPerPixelX
            
            .userGUI(1).Move (.vSizer.Left + .vSizer.Width) + Screen.TwipsPerPixelX, 0, (.ScaleWidth) - (.vSizer.Left + .vSizer.Width) - (Border * Screen.TwipsPerPixelX)
            
            If .mnuMultiThread.Checked Then
                .userInfo.Top = (.hSizer.Top + .hSizer.Height + Screen.TwipsPerPixelY)
                .userGUI(0).Height = (.hSizer.Top)
                .userGUI(1).Height = (.hSizer.Top)
                .userInfo.Height = .ScaleHeight - .userInfo.Top - (Border * Screen.TwipsPerPixelY)
                .vSizer.Height = .userInfo.Top - (Border * Screen.TwipsPerPixelY) - Screen.TwipsPerPixelY
                PaintInfoGUI myForm
            Else
                .userGUI(0).Height = (.ScaleHeight) - (Border * Screen.TwipsPerPixelY) - Screen.TwipsPerPixelY
                .userGUI(1).Height = (.ScaleHeight) - (Border * Screen.TwipsPerPixelY) - Screen.TwipsPerPixelY
                .vSizer.Height = (.ScaleHeight) - (Border * Screen.TwipsPerPixelY)
            End If
            
            PaintClientGUI myForm, 0
            PaintClientGUI myForm, 1
        
        Else
            If Not .userGUI(0).Visible Then .userGUI(0).Visible = True
            If .userGUI(1).Visible Then .userGUI(1).Visible = False
            
            .vSizer.Visible = False
            If .mnuMultiThread.Checked Then
                If Not .userInfo.Visible = True Then .userInfo.Visible = True
                .hSizer.Visible = True
                
                If .hSizer.Top < (Border * Screen.TwipsPerPixelY) * 2 Then .hSizer.Top = (Border * Screen.TwipsPerPixelY) * 2
                If .hSizer.Top + ((Border * Screen.TwipsPerPixelY) * 2) > .ScaleHeight Then .hSizer.Top = .ScaleHeight - ((Border * Screen.TwipsPerPixelY) * 2)
                .hSizer.Move (Border * Screen.TwipsPerPixelX), , .ScaleWidth - ((Border * Screen.TwipsPerPixelX) * 2)
                
                If .ReCenterSizers Then
                    .hSizer.Top = (0.68 * .ScaleHeight)
                End If
                
                .userInfo.Move (Border * Screen.TwipsPerPixelX), , (.ScaleWidth) - ((Border * Screen.TwipsPerPixelX) * 2)
            
            Else
                If Not .userInfo.Visible = False Then .userInfo.Visible = False
                .hSizer.Visible = False
            End If
           
            .userGUI(0).Move (Border * Screen.TwipsPerPixelX), (Border * Screen.TwipsPerPixelY), .ScaleWidth - ((Border * Screen.TwipsPerPixelX) * 2)
        
            If .mnuMultiThread.Checked Then
                .userInfo.Top = (.hSizer.Top + .hSizer.Height + Screen.TwipsPerPixelY)
                .userGUI(0).Height = (.hSizer.Top) - ((Border * Screen.TwipsPerPixelY)) - Screen.TwipsPerPixelY
                .userInfo.Height = .ScaleHeight - .userInfo.Top - (Border * Screen.TwipsPerPixelY)
                PaintInfoGUI myForm
            Else
                .userGUI(0).Height = (.ScaleHeight) - (Border * Screen.TwipsPerPixelY) - (Screen.TwipsPerPixelY * 2)
            End If
        
            PaintClientGUI myForm, 0
        End If

    Err.Clear
    
    End With
End Sub

Public Sub ViewDoubleWindow(ByRef myForm, ByVal IsDouble As Boolean, Optional ByVal SetValue As Boolean = True)
    With myForm
    
        .mnuDoubleWin.Checked = IsDouble
        
        If IsDouble Then
            LoadGUI myForm, 1
        Else
            UnloadGUI myForm, 1
        End If
        
        FormResize myForm
    
    End With
End Sub
Public Sub ViewTransferInfo(ByRef myForm, ByVal IsVisible As Boolean, Optional SetValue As Boolean = True)
    With myForm

        .mnuMultiThread.Checked = IsVisible
            
         .ReCenterSizers = True
        FormResize myForm
        .ReCenterSizers = False
        FormResize myForm
    
    End With
End Sub

Public Sub EnableTransferInfo(ByRef myForm, ByVal enabled As Boolean)
    With myForm
        
        .mnuMultiThread.Checked = enabled
        
        FormResize myForm

    End With
End Sub

Public Sub ViewStatusBar(ByRef myForm, ByVal IsVisible As Boolean, Optional SetValue As Boolean = True)
    With myForm
    
        .mnuShowStatusBar.Checked = IsVisible

        PaintClientGUI myForm, 0
        If .mnuDoubleWin.Checked Then PaintClientGUI myForm, 1

    End With
End Sub
Public Sub ViewToolBar(ByRef myForm, ByVal IsVisible As Boolean, Optional SetValue As Boolean = True)
    With myForm
    
        .mnuShowToolBar.Checked = IsVisible

        .dContainer(GetContainerIndex(0, 1)).Visible = IsVisible
        .dContainer(GetContainerIndex(0, 0)).Visible = IsVisible
        
        PaintClientGUI myForm, 0
        If .mnuDoubleWin.Checked Then PaintClientGUI myForm, 1
        
    End With
End Sub
Public Sub ViewDriveList(ByRef myForm, ByVal IsVisible As Boolean, Optional SetValue As Boolean = True)
    With myForm
        
        .mnuShowDriveList.Checked = IsVisible

        .dContainer(GetContainerIndex(1, 1)).Visible = IsVisible
        .dContainer(GetContainerIndex(1, 0)).Visible = IsVisible
        PaintClientGUI myForm, 0
        If .mnuDoubleWin.Checked Then PaintClientGUI myForm, 1
    
    End With
End Sub
Public Sub ViewAddressBar(ByRef myForm, ByVal IsVisible As Boolean, Optional SetValue As Boolean = True)
    With myForm
    
        .mnuShowAddressBar.Checked = IsVisible

        .dContainer(GetContainerIndex(2, 1)).Visible = IsVisible
        .dContainer(GetContainerIndex(2, 0)).Visible = IsVisible
        PaintClientGUI myForm, 0
        If .mnuDoubleWin.Checked Then PaintClientGUI myForm, 1

    End With
End Sub

Public Sub EnableFileMenu(ByRef myForm, ByVal Index As Integer)
    With myForm
    
        .FocusIndex = Index
        
        Dim selCount As Integer
        Dim IsConnected As Boolean
        
        Dim myClient As NTAdvFTP61.Client
        Set myClient = .SetMyClient(Index)
        IsConnected = myClient.IsConnected
        Set myClient = Nothing
        
        selCount = GetSelectedCount(.pView(Index))
        
        .mnuOpen.enabled = (IsConnected And .GetState(Index) = st_ProcessSuccess) And (selCount = 1)
        .mnuCut.enabled = (IsConnected And .GetState(Index) = st_ProcessSuccess) And (selCount > 0)
        .mnuCopy.enabled = (IsConnected And .GetState(Index) = st_ProcessSuccess) And (selCount > 0)
        .mnuPaste.enabled = (IsConnected And .GetState(Index) = st_ProcessSuccess)
        .mnuNewFolder.enabled = (IsConnected And .GetState(Index) = st_ProcessSuccess)
        .mnuDelete.enabled = (IsConnected And .GetState(Index) = st_ProcessSuccess) And (selCount > 0)
        .mnuRename.enabled = (IsConnected And .GetState(Index) = st_ProcessSuccess) And (selCount = 1)
    
        .mnuSelAll.enabled = (IsConnected And .GetState(Index) = st_ProcessSuccess)
        .mnuSelFiles.enabled = (IsConnected And .GetState(Index) = st_ProcessSuccess)
        .mnuSelFolders.enabled = (IsConnected And .GetState(Index) = st_ProcessSuccess)
        .mnuWildCard.enabled = (IsConnected And .GetState(Index) = st_ProcessSuccess)
    
        .mnuRefresh.enabled = IsConnected
        .mnuStop.enabled = (IsConnected And .GetState(Index) = st_Processing)
    
        .mnuConnect.enabled = Not IsConnected
        .mnuDisconnect = IsConnected

    End With
End Sub


