VERSION 5.00
Begin VB.UserControl AutoType 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "AutoType.ctx":0000
   Begin VB.ListBox lstFolders 
      Height          =   840
      Left            =   1095
      TabIndex        =   3
      Top             =   2415
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.ListBox lstHistory 
      Height          =   1035
      Left            =   930
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.TextBox txtTypeIn 
      Height          =   315
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.ComboBox cmbTypeIn 
      Height          =   315
      Left            =   15
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   3705
   End
End
Attribute VB_Name = "AutoType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary

Enum AutoCompleteStyles
    None = 0
    ListComplete = 1
    TextComplete = 2
End Enum

Private Const Default_Text = ""

Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_HIDE = 0

Private Const LB_FINDSTRING As Long = &H18F
Private Const CB_FINDSTRING As Long = &H14C
Private Const CB_SHOWDROPDOWN As Long = &H14F
Private Const CB_GETCOUNT = &H146
Private Const CB_GETCURSEL = &H147
Private Const CB_GETEDITSEL = &H140
Private Const CB_SELECTSTRING = &H14D
Private Const CB_SETCURSEL = &H14E
Private Const CB_SETEDITSEL = &H142

Private cancelTracking As Boolean

Private isEnabled As Boolean
Private myCompleteStyle As Integer
Private myTracking As Integer
Private maxHistory As Integer
Private isHistoryVisible As Boolean
Private myToolTipText As String
Private isURLEnabled As Boolean
Private isPathEnabled As Boolean
Private isReadOnly As Boolean
Private pAllowDuplicates As Boolean
Private pRootFolders As Boolean

Public Event Change()
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Private pDefaultHWnd As Long
Private CurrentSelection As Long

Public Property Get BackColor() As Long
    BackColor = txtTypeIn.BackColor
End Property
Public Property Let BackColor(ByVal newVal As Long)
    cmbTypeIn.BackColor = newVal
    txtTypeIn.BackColor = newVal
End Property

Public Property Get ForeColor() As Long
    ForeColor = txtTypeIn.ForeColor
End Property
Public Property Let ForeColor(ByVal newVal As Long)
    cmbTypeIn.ForeColor = newVal
    txtTypeIn.ForeColor = newVal
End Property

Private Function GetURLText(ByVal pText As String, Optional ByRef pLen As Integer = 0)
    
    If Len(pText) <= 7 Then
        If LCase(pText) = Left("http://", Len(pText)) Or LCase(pText) = Left("ftp://", Len(pText)) Then
            GetURLText = pText
        Else
            GetURLText = "ftp://" & pText
            If pLen > 0 Then
                pLen = pLen + 6
            End If
        End If
    Else
        If LCase(Left(pText, 7)) <> "http://" And LCase(Left(pText, 6)) <> "ftp://" Then
            GetURLText = "ftp://" & pText
            If pLen > 0 Then
                pLen = pLen + 6
            End If
        Else
            GetURLText = pText
        End If
    End If
End Function

Private Sub TrackText()
On Error GoTo catch
    If Not cancelTracking Then
        cancelTracking = True
        Dim srhText
        Dim srhLen As Integer
        Dim lstIndex As Long
        If isHistoryVisible Then
            Set srhText = cmbTypeIn
        Else
            Set srhText = txtTypeIn
        End If
        srhLen = Len(srhText.Text)
        If (srhText <> "") And (srhLen > 1) Then
            Dim newText As String
            Select Case myCompleteStyle
                Case AutoCompleteStyles.TextComplete
                    If isURLEnabled Then
                        lstIndex = SendMessageStr(lstHistory.hwnd, LB_FINDSTRING, -1, GetURLText(srhText.Text))
                        If lstIndex = -1 Then lstIndex = SendMessageStr(lstHistory.hwnd, LB_FINDSTRING, -1, srhText.Text)
                    Else
                        lstIndex = SendMessageStr(lstHistory.hwnd, LB_FINDSTRING, -1, srhText.Text)
                    End If
                    If (lstIndex > -1) And (srhText.Text <> "") Then
                        srhText.Text = lstHistory.List(lstIndex)
                        srhText.SelStart = srhLen
                        srhText.SelLength = Len(lstHistory.List(lstIndex)) - srhLen
                    End If
                Case AutoCompleteStyles.ListComplete
                    Dim lstUse


                    If isPathEnabled Then
                        Dim fso As Object
                        Set fso = CreateObject("Scripting.FileSystemObject")
                        Dim pFolder As String
                        If InStr(srhText.Text, "\") > 0 Then
                            pFolder = Left(srhText.Text, InStrRev(srhText.Text, "\") - 1)
                        Else
                            pFolder = srhText.Text
                        End If
                        If isURLEnabled Then
                            lstIndex = SendMessageStr(lstHistory.hwnd, LB_FINDSTRING, -1, GetURLText(pFolder))
                            If lstIndex = -1 Then lstIndex = SendMessageStr(lstHistory.hwnd, LB_FINDSTRING, -1, pFolder)
                        Else
                            lstIndex = SendMessageStr(lstHistory.hwnd, LB_FINDSTRING, -1, pFolder)
                        End If
                        If (fso.FolderExists(pFolder) And (lstIndex = -1)) And isPathEnabled Then
                            lstFolders.Clear
                            Dim f As Object 'Scripting.Folder
                            Dim fItem As Object 'Scripting.Folder
                            Set f = fso.GetFolder(pFolder)
                            For Each fItem In f.SubFolders
                                lstFolders.AddItem fItem
                            Next
                            For Each fItem In f.Files
                                lstFolders.AddItem fItem
                            Next
                    
                            Set lstUse = lstFolders
                            newText = srhText.Text
                        Else
                            Set lstUse = lstHistory
                            newText = srhText.Text
                        End If
                        Set fso = Nothing
                    Else
                        Set lstUse = lstHistory
                        newText = srhText.Text
                    End If
                    If isURLEnabled Then
                        lstIndex = SendMessageStr(lstUse.hwnd, LB_FINDSTRING, -1, GetURLText(newText))
                        If lstIndex = -1 Then lstIndex = SendMessageStr(lstUse.hwnd, LB_FINDSTRING, -1, newText)
                    Else
                        lstIndex = SendMessageStr(lstUse.hwnd, LB_FINDSTRING, -1, newText)
                    End If
                    If (lstIndex > -1) And (newText <> "") Then
                        ResizeComboList
                        
                        ShowWindow
                                                
                        frmCombo.lstMatch.Clear
                        Set frmCombo.srhText = Nothing
                        Set frmCombo.srhText = srhText
                        Dim oldIndex As Long
                        lstIndex = -1
                        Do
                            oldIndex = lstIndex
                            If isURLEnabled Then
                                lstIndex = SendMessageStr(lstUse.hwnd, LB_FINDSTRING, oldIndex, GetURLText(newText))
                                If lstIndex = -1 Then lstIndex = SendMessageStr(lstUse.hwnd, LB_FINDSTRING, oldIndex, newText)
                            Else
                                lstIndex = SendMessageStr(lstUse.hwnd, LB_FINDSTRING, oldIndex, newText)
                            End If
                            If (lstIndex > -1) And Not (lstIndex <= oldIndex) Then
                                frmCombo.lstMatch.AddItem lstUse.List(lstIndex)
                            End If
                        Loop Until (lstIndex = -1) Or (lstIndex <= oldIndex)
                    Else
                        HideWindow
                    End If
                    
                    CurrentSelection = -1
            End Select

        End If
        cancelTracking = False
    End If
Exit Sub
catch:
    Err.Clear
End Sub

Public Function ShowHistory()
    If Not isHistoryVisible Then
        If frmCombo.Visible Then HideWindow
        Dim r As Long
        r = SendMessageLong(cmbTypeIn.hwnd, CB_SHOWDROPDOWN, True, 0)
    End If
End Function
Public Function HideHistory()
    If Not isHistoryVisible Then
        Dim r As Long
        r = SendMessageLong(cmbTypeIn.hwnd, CB_SHOWDROPDOWN, False, 0)
    End If
End Function

Public Sub SetFocus()
    On Error Resume Next
    If isHistoryVisible Then
        SetFocusAPI cmbTypeIn.hwnd
    Else
        SetFocusAPI txtTypeIn.hwnd
    End If
    Err.Clear
End Sub

Public Property Get SelStart() As Integer
    If isHistoryVisible Then
        SelStart = cmbTypeIn.SelStart
    Else
        SelStart = txtTypeIn.SelStart
    End If
End Property
Public Property Let SelStart(ByVal newValue As Integer)
    If isHistoryVisible Then
        cmbTypeIn.SelStart = newValue
    Else
        txtTypeIn.SelStart = newValue
    End If
End Property

Public Property Get SelLength() As Integer
    If isHistoryVisible Then
        SelLength = cmbTypeIn.SelLength
    Else
        SelLength = txtTypeIn.SelLength
    End If
End Property
Public Property Let SelLength(ByVal newValue As Integer)
    If isHistoryVisible Then
        cmbTypeIn.SelLength = newValue
    Else
        txtTypeIn.SelLength = newValue
    End If
End Property

Public Sub SetToList(ByVal lstBox As Variant)
    Me.Clear
    Dim cnt As Integer
    If lstBox.ListCount > 0 Then
        For cnt = 0 To lstBox.ListCount - 1
            Me.AddItem lstBox.List(cnt)
        Next
    End If
End Sub

Public Sub AddItem(ByVal lstText As String)
    If (lstHistory.ListCount < maxHistory) Or maxHistory = 0 Then
        Dim lstIndex As Integer
        If lstHistory.ListCount > 0 Then
            For lstIndex = 0 To (lstHistory.ListCount - 1)
                If Trim(LCase(lstHistory.List(lstIndex))) = Trim(LCase(lstText)) Then
                    If Not pAllowDuplicates Then Exit Sub
                End If
            Next
        End If
        lstHistory.AddItem lstText
        cmbTypeIn.AddItem lstText
        
        If Not pAllowDuplicates Then
        Else
            lstHistory.AddItem lstText
            cmbTypeIn.AddItem lstText
        End If
    End If
End Sub
Public Sub RemoveItem(ByVal lstIndex As Integer)
    lstHistory.RemoveItem lstIndex
    cmbTypeIn.RemoveItem lstIndex
End Sub
Public Sub Clear()
    Dim sTmp As String
    sTmp = Me.Text
    
    Do Until lstHistory.ListCount <= 0
        lstHistory.RemoveItem 0
    Loop
    Do Until cmbTypeIn.ListCount <= 0
        cmbTypeIn.RemoveItem 0
    Loop
    
    Me.Text = sTmp
End Sub

Public Property Get ListCount() As Integer
    ListCount = lstHistory.ListCount
End Property

Public Property Get HistorySize() As Integer
    HistorySize = maxHistory
End Property
Public Property Let HistorySize(ByVal newValue As Integer)
    maxHistory = newValue
End Property

Public Property Get AllowDuplicates() As Boolean
    AllowDuplicates = pAllowDuplicates
End Property
Public Property Let AllowDuplicates(ByVal newValue As Boolean)
    pAllowDuplicates = newValue
End Property

Public Property Get RootFolders() As Boolean
    RootFolders = pRootFolders
End Property
Public Property Let RootFolders(ByVal newValue As Boolean)
    pRootFolders = newValue
End Property

Public Property Get HistoryVisible() As Boolean
    HistoryVisible = isHistoryVisible
End Property
Public Property Let HistoryVisible(ByVal newValue As Boolean)
    isHistoryVisible = newValue
    cmbTypeIn.Visible = newValue
    txtTypeIn.Visible = Not newValue
End Property

Public Property Get ToolTipText() As String
    ToolTipText = myToolTipText
End Property
Public Property Let ToolTipText(ByVal newValue As String)
    myToolTipText = newValue
    cmbTypeIn.ToolTipText = myToolTipText
    txtTypeIn.ToolTipText = myToolTipText
End Property

Public Property Get DefaultHWnd() As Long
    If pDefaultHWnd = 0 Then
        DefaultHWnd = UserControl.Parent.hwnd
    Else
        DefaultHWnd = pDefaultHWnd
    End If
End Property
Public Property Let DefaultHWnd(ByVal newVal As Long)
    If Not isHooked Then
        pDefaultHWnd = newVal
    End If
End Property

Public Property Get Enabled() As Boolean
    Enabled = isEnabled
End Property
Public Property Let Enabled(ByVal newValue As Boolean)
    isEnabled = newValue
    cmbTypeIn.Enabled = isEnabled
    txtTypeIn.Enabled = isEnabled
End Property

Public Property Get ReadOnly() As Boolean
    ReadOnly = isReadOnly
End Property
Public Property Let ReadOnly(ByVal newValue As Boolean)
    isReadOnly = newValue
    cmbTypeIn.Locked = isReadOnly
    txtTypeIn.Locked = isReadOnly
End Property
Public Property Get Text() As String
    On Error Resume Next
    If isHistoryVisible Then
        Text = IIf(isHistoryVisible, cmbTypeIn.Text, txtTypeIn.Text)
    End If
End Property
Public Property Let Text(ByVal newText As String)
    On Error Resume Next
    cancelTracking = True
    If Not isReadOnly Then
        If Not cmbTypeIn.Text = newText Then cmbTypeIn.Text = newText
        If Not txtTypeIn.Text = newText Then txtTypeIn.Text = newText
    End If
    txtObject_Change
    cancelTracking = False
End Property

Public Property Get CompleteStyle() As AutoCompleteStyles
    CompleteStyle = myCompleteStyle
End Property
Public Property Let CompleteStyle(ByVal newValue As AutoCompleteStyles)
    myCompleteStyle = newValue
End Property

Public Property Get URLEnabled() As Boolean
    URLEnabled = isURLEnabled
End Property
Public Property Let URLEnabled(ByVal newValue As Boolean)
    isURLEnabled = newValue
End Property

Public Property Get PathEnabled() As Boolean
    PathEnabled = isPathEnabled
End Property
Public Property Let PathEnabled(ByVal newValue As Boolean)
    isPathEnabled = newValue
End Property

Private Sub cmbTypeIn_DropDown()
    HideWindow
End Sub

Private Sub cmbTypeIn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Then
    
    End If
    If isHistoryVisible Then
        txtObject_KeyDown KeyCode, Shift
    End If
End Sub
Private Sub cmbTypeIn_KeyPress(KeyAscii As Integer)
    If isHistoryVisible Then
        txtObject_KeyPress KeyAscii
    End If
End Sub
Private Sub cmbTypeIn_KeyUp(KeyCode As Integer, Shift As Integer)
    If isHistoryVisible Then
        txtObject_KeyUp KeyCode, Shift
    End If
End Sub
Private Sub cmbTypeIn_Change()
    If isHistoryVisible And Not cancelTracking Then
        TrackText
        txtObject_Change
    End If
End Sub
Private Sub cmbTypeIn_Click()
    If isHistoryVisible Then
        txtObject_Click
        txtObject_Change
    End If
End Sub
Private Sub cmbTypeIn_DblClick()
    If isHistoryVisible Then
        txtObject_DblClick
    End If
End Sub
Private Sub cmbTypeIn_GotFocus()
    If isHistoryVisible Then
        txtObject_GotFocus
    End If
End Sub
Private Sub cmbTypeIn_LostFocus()
    If isHistoryVisible Then
        txtObject_LostFocus
    End If
End Sub

Private Sub txtTypeIn_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not isHistoryVisible Then
        txtObject_KeyDown KeyCode, Shift
    End If
End Sub
Private Sub txtTypeIn_KeyPress(KeyAscii As Integer)
    If Not isHistoryVisible Then
        txtObject_KeyPress KeyAscii
    End If
End Sub
Private Sub txtTypeIn_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not isHistoryVisible Then
        txtObject_KeyUp KeyCode, Shift
    End If
End Sub
Private Sub txtTypeIn_Change()
    If (Not isHistoryVisible) And Not cancelTracking Then
        TrackText
        txtObject_Change
    End If
End Sub
Private Sub txtTypeIn_Click()
    If Not isHistoryVisible Then
        txtObject_Click
    End If
End Sub
Private Sub txtTypeIn_DblClick()
    If Not isHistoryVisible Then
        txtObject_DblClick
    End If
End Sub
Private Sub txtTypeIn_GotFocus()
    If Not isHistoryVisible Then
        txtObject_GotFocus
    End If
End Sub
Private Sub txtTypeIn_LostFocus()
    If Not isHistoryVisible Then
        txtObject_LostFocus
    End If
End Sub

Private Sub txtObject_KeyDown(KeyCode As Integer, Shift As Integer)

    If (frmCombo.lstMatch.ListCount > 0) And (frmCombo.Visible = True) Then
        Select Case KeyCode
            Case 38
                KeyCode = 0
            Case 40
                SetFocusAPI frmCombo.lstMatch.hwnd
                frmCombo.lstMatch.ListIndex = 0
               
                KeyCode = 0
            Case 13
                If isHistoryVisible Then
                    AddItem cmbTypeIn.Text
                Else
                    AddItem txtTypeIn.Text
                End If
                
                HideWindow
        End Select
    End If
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub txtObject_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub txtObject_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub txtObject_Change()
    RaiseEvent Change
End Sub
Private Sub txtObject_Click()
    RaiseEvent Click
End Sub
Private Sub txtObject_DblClick()
    RaiseEvent DblClick
End Sub
Private Sub txtObject_GotFocus()
    txtTypeIn.SelStart = 0
    txtTypeIn.SelLength = Len(txtTypeIn.Text)
End Sub
Private Sub txtObject_LostFocus()
    If frmCombo.Visible Then
        If (Not (GetActiveWindow = frmCombo.hwnd)) And _
            (Not (GetActiveWindow = UserControl.hwnd)) And _
            (Not (GetActiveWindow = frmCombo.lstMatch.hwnd)) And _
            (Not (GetActiveWindow = cmbTypeIn.hwnd)) Then
            
            HideWindow
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    cancelTracking = False
    Load frmCombo
End Sub
Private Sub UserControl_InitProperties()
    Me.RootFolders = False
    Me.AllowDuplicates = False
    Me.CompleteStyle = AutoCompleteStyles.ListComplete
    Me.HistorySize = 0
    Me.HistoryVisible = True
    Me.Enabled = True
    Me.URLEnabled = True
    Me.PathEnabled = True
    Me.ToolTipText = ""
    Me.Text = Default_Text
    Me.ReadOnly = False
    Me.BackColor = SystemColorConstants.vbWindowBackground
    Me.ForeColor = SystemColorConstants.vbWindowText
    
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Me.RootFolders = .ReadProperty("RootFolders", False)
        Me.AllowDuplicates = .ReadProperty("AllowDuplicates", False)
        Me.CompleteStyle = .ReadProperty("CompleteStyle", AutoCompleteStyles.ListComplete)
        Me.HistorySize = .ReadProperty("HistorySize", 0)
        Me.HistoryVisible = .ReadProperty("HistoryVisible", True)
        Me.Enabled = .ReadProperty("Enabled", True)
        Me.URLEnabled = .ReadProperty("URLEnabled", True)
        Me.PathEnabled = .ReadProperty("PathEnabled", True)
        Me.ToolTipText = .ReadProperty("ToolTipText", "")
        Me.Text = .ReadProperty("Text", Default_Text)
        Me.ReadOnly = .ReadProperty("ReadOnly", False)
        Me.BackColor = .ReadProperty("BackColor", SystemColorConstants.vbWindowBackground)
        Me.ForeColor = .ReadProperty("ForeColor", SystemColorConstants.vbWindowText)
    End With
End Sub

Private Sub ShowWindow()
    
    
    modEditor.ShowWindow frmCombo.hwnd, SW_SHOWNOACTIVATE
    SendMessageLong frmCombo.hwnd, SW_SHOWNOACTIVATE, False, 0&
    TopMostForm frmCombo, True, True
    Hook DefaultHWnd
    
End Sub

Private Sub UserControl_Terminate()
    Unload frmCombo
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "RootFolders", pRootFolders, False
        .WriteProperty "AllowDuplicates", pAllowDuplicates, False
        .WriteProperty "CompleteStyle", myCompleteStyle, AutoCompleteStyles.ListComplete
        .WriteProperty "HistorySize", maxHistory, 0
        .WriteProperty "HistoryVisible", isHistoryVisible, True
        .WriteProperty "Enabled", isEnabled, True
        .WriteProperty "URLEnabled", isURLEnabled, True
        .WriteProperty "PathEnabled", isPathEnabled, True
        .WriteProperty "ToolTipText", myToolTipText, ""
        If isHistoryVisible Then
            .WriteProperty "Text", cmbTypeIn.Text, Default_Text
        Else
            .WriteProperty "Text", txtTypeIn.Text, Default_Text
        End If
        .WriteProperty "ReadOnly", isReadOnly, ""
        .WriteProperty "BackColor", txtTypeIn.BackColor, SystemColorConstants.vbWindowBackground
        .WriteProperty "ForeColor", txtTypeIn.ForeColor, SystemColorConstants.vbWindowText
        
    End With
End Sub
Private Sub UserControl_Resize()
    cmbTypeIn.Top = 0
    cmbTypeIn.Left = 0
    txtTypeIn.Top = 0
    txtTypeIn.Left = 0
    UserControl.Height = cmbTypeIn.Height
    cmbTypeIn.Width = UserControl.Width
    txtTypeIn.Width = UserControl.Width
    
    ResizeComboList
End Sub

Private Sub ResizeComboList()
    On Error Resume Next
    Dim rctCombo As RectType
    GetWindowRect cmbTypeIn.hwnd, rctCombo
    
    frmCombo.Top = (rctCombo.Bottom * Screen.TwipsPerPixelY)
    frmCombo.Left = (rctCombo.Left * Screen.TwipsPerPixelX)
    frmCombo.Width = ((rctCombo.Right - rctCombo.Left) * Screen.TwipsPerPixelX)
    frmCombo.ResizeBox
    Err.Clear
End Sub








Attribute 