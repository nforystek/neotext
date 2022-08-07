#Const [True] = -1
#Const [False] = 0
Attribute VB_Name = "modMenuUp"



#Const modMenuUp = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module

Public Enum MenuActions
    FileMenu = 2
End Enum

Public Enum SubMenuActions
    OpenMenu = 0
    CutMenu = 2
    CopyMenu = 3
    PasteMenu = 4
    NewMenu = 6
    DeleteMenu = 7
    RenameMenu = 8
    SelectMenu = 10
    RefreshMenu = 12
    StopMenu = 13
    ConnectMenu = 15
    DisconnectMenu = 16
End Enum

Public Enum SelectMenuActions
    SelectAll = 0
    SelectFiles = 1
    SelectFolders = 2
    SelectPattern = 3
End Enum
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long        ':( Missing Scope
Private Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal Y As Long, ByVal hwnd As Long, ByVal lptpm As Any) As Long
Private Const MF_BYPOSITION = &H400&
Private Const TPM_LEFTALIGN = &H0&

Public Sub PopUp(ByVal FormHWnd As Long, ByVal FormMenu As Long)

    Dim MousePoint As POINTAPI

    GetCursorPos MousePoint

    TrackPopupMenuEx GetSubMenu(GetMenu(FormHWnd), FormMenu), TPM_LEFTALIGN, MousePoint.x, MousePoint.Y, FormHWnd, ByVal 0&

End Sub
