#Const [True] = -1
#Const [False] = 0
Attribute VB_Name = "modForms"



#Const modForms = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_NOZORDER = &H4

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Sub GetFormByHWND(ByRef tForm As Form, ByVal hwnd As Long)

    Dim cForm
    For Each cForm In Forms
        If cForm.hwnd = hwnd Then
            Set tForm = cForm
        End If
    Next

End Sub

Public Function TopMostForm(ByRef frm As Form, ByVal SetOnTop As Boolean, Optional ByVal SwapFlags As Boolean = False) As Long
    With frm
        TopMostForm = SetWindowPos(.hwnd, _
                IIf(SetOnTop, HWND_TOPMOST, HWND_NOTOPMOST), _
                .Left / Screen.TwipsPerPixelX, _
                .Top / Screen.TwipsPerPixelY, _
                .Width / Screen.TwipsPerPixelX, _
                .Height / Screen.TwipsPerPixelY, _
                IIf(SwapFlags, SWP_NOACTIVATE Or SWP_SHOWWINDOW, 0))
    End With
End Function

Public Function IsFormVisible(ByVal frmHWnd As Long) As Boolean

    Dim cnt As Integer
    Dim vis As Boolean
    vis = False
    For cnt = 0 To Forms.count - 1
        If Forms(cnt).hwnd = frmHWnd Then
            If Forms(cnt).Visible Then
                vis = True
                Exit For
            End If
        End If
    Next
    
    IsFormVisible = vis

End Function

Public Function GetForm(ByVal frmHWnd As Long) As Form

    Dim cnt As Integer
    Dim frm As Form
    Set frm = Nothing
    For cnt = 0 To Forms.count - 1
        If Forms(cnt).hwnd = frmHWnd Then
            Set frm = Forms(cnt)
        End If
    Next
    
    Set GetForm = frm

End Function
