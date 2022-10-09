Attribute VB_Name = "modDocking"
Option Explicit

Option Private Module

Private SingleForm As MDIForm
Private DockedForm As Collection

Public Static Property Get MDIForm() As Form
    Set MDIForm = SingleForm
End Property
Public Static Property Get Count() As Long
    If DockedForm Is Nothing Then
        Count = 0
    Else
        Count = DockedForm.Count
    End If
End Property
Public Static Property Set Instance(ByRef RHS As Form)
    If RHS Is Nothing Then
        If DockedForm Is Nothing Then
            If Not SingleForm Is Nothing Then Set SingleForm = Nothing
        Else
            DockedForm.Remove 1
            If DockedForm.Count = 0 Then Set DockedForm = Nothing
        End If
    Else
        If SingleForm Is Nothing Then
            Set SingleForm = RHS
        Else
            If DockedForm Is Nothing Then
                Set DockedForm = New Collection
            End If
        End If
        DockedForm.Add RHS, "H" & RHS.hwnd
    End If
End Property
