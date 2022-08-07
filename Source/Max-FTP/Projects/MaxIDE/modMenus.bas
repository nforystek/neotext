#Const [True] = -1
#Const [False] = 0
Attribute VB_Name = "modMenus"
#Const modMenus = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module

Public CancelTabClick As Boolean

Public Sub RefreshWindowMenu()

    With frmMainIDE
        Dim mnu
        For Each mnu In .mnuWindow
            If mnu.Index <> 0 Then Unload mnu
        Next
        .WindowTabs.Tabs.Clear
       
        Dim frm As Form, childCnt As Integer, ntab
        childCnt = 0
        For Each frm In Forms
            If TypeName(frm) = "frmScriptPage" Or TypeName(frm) = "frmIEPage" Then
                If frm.MDIChild And frm.Visible Then
                    childCnt = childCnt + 1
                    Load .mnuWindow(childCnt)
                    .mnuWindow(childCnt).Visible = True
                    .mnuWindow(childCnt).Caption = "&" & childCnt & " " & frm.Caption
                    .mnuWindow(childCnt).Tag = frm.hwnd
                    .mnuWindow(childCnt).Checked = False
                    Set ntab = .WindowTabs.Tabs.Add(childCnt, , frm.Caption)
                    ntab.Tag = frm.hwnd
                End If
            End If
        Next
        
        .mnuWindow(0).Visible = (childCnt > 0)
    End With
End Sub
Public Sub ShowSelectedWindow(ByVal WinCaption As String)
    With frmMainIDE
        Dim frm As Form
        For Each frm In Forms
            If TypeName(frm) = "frmScriptPage" Or TypeName(frm) = "frmIEPage" Then
                If WinCaption = frm.Caption Then
                    frm.Show
                    frm.ZOrder 0
                End If
            End If
        Next
    End With
End Sub
Public Sub SelectWindowMenu(ByVal WinCaption As String)
    CancelTabClick = True
    With frmMainIDE
        Dim cnt As Integer
        If .mnuWindow.Count > 0 Then
            For cnt = 1 To .mnuWindow.Count - 1
                If WinCaption = Mid(.mnuWindow(cnt).Caption, 4) Then
                    .mnuWindow(cnt).Checked = True
                    .WindowTabs.Tabs(cnt).Selected = True
                Else
                    .mnuWindow(cnt).Checked = False
                End If
            Next
        End If
    End With

    CancelTabClick = False
End Sub

Public Function RefreshToolbarMenu()
    
    With frmMainIDE.ScriptControls
        .Buttons("script_newproject").Enabled = (Not Project.IsRunning)
        .Buttons("script_openproject").Enabled = (Not Project.IsRunning)
        .Buttons("script_saveproject").Enabled = Project.Loaded And (Not Project.IsRunning)
        .Buttons("script_addfile").Enabled = Project.Loaded And (Not Project.IsRunning) And (Not Project.IsTemplate)
        .Buttons("script_removefile").Enabled = Project.Loaded And (Not Project.IsRunning) And (Not Project.IsTemplate) And frmProjectExplorer.AllowDelete
        .Buttons("script_undo").Enabled = EnableUndo
        .Buttons("script_redo").Enabled = EnableRedo
        .Buttons("script_cut").Enabled = EnableCut
        .Buttons("script_copy").Enabled = EnableCopy
        .Buttons("script_paste").Enabled = EnablePaste
        .Buttons("script_find").Enabled = EnableFind
        .Buttons("script_stop").Enabled = Project.Loaded And Project.IsRunning And (Not Project.IsTemplate)
        .Buttons("script_run").Enabled = Project.Loaded And (Not Project.IsRunning) And (Not Project.IsTemplate)
    End With

End Function

Public Function RefreshScriptMenu()
    
    With frmMainIDE
        
        .mnuNew.Enabled = (Not Project.IsRunning)
        .mnuOpen.Enabled = (Not Project.IsRunning)
    
        .mnuSaveProject.Enabled = Project.Loaded And (Not Project.IsRunning)
        .mnuSaveProjectAs.Enabled = Project.Loaded And (Not Project.IsRunning)
        
        .mnuDebugMap.Enabled = Project.Loaded And (Not Project.IsRunning) And (Not Project.IsTemplate)
        .mnuRun.Enabled = Project.Loaded And (Not Project.IsRunning) And (Not Project.IsTemplate)
        .mnuStop.Enabled = Project.Loaded And Project.IsRunning And (Not Project.IsTemplate)
        
        .mnuClose.Enabled = Project.Loaded And (Not Project.IsRunning)
                
        .mnuOpenItem.Enabled = Project.Loaded And (Not Project.IsRunning) And frmProjectExplorer.AllowOpen
        .mnuNewItem.Enabled = Project.Loaded And (Not Project.IsRunning) And (Not Project.IsTemplate)
        .mnuDeleteItem.Enabled = Project.Loaded And (Not Project.IsRunning) And (Not Project.IsTemplate) And frmProjectExplorer.AllowDelete
        .mnuRename.Enabled = Project.Loaded And (Not Project.IsRunning) And frmProjectExplorer.AllowRename
        .mnuExamine.Enabled = Project.Loaded And (Not Project.IsRunning)
        

        .mnuExamples.Enabled = (Not Project.IsRunning)
        .mnuExamples2.Enabled = (Not Project.IsRunning)
        .mnuExperiments.Enabled = (Not Project.IsRunning)
    End With

End Function
Public Function RefreshEditMenu()
    With frmMainIDE
        .mnuUndo.Enabled = EnableUndo
        .mnuRedo.Enabled = EnableRedo
    
        .mnuCut.Enabled = EnableCut
        .mnuCopy.Enabled = EnableCopy
        
        .mnuPaste.Enabled = EnablePaste
        
        .mnuDelete.Enabled = EnableDelete
        .mnuSelectAll.Enabled = EnableSelectAll
        
        .mnuComment.Enabled = EnableComment
        .mnuUncomment.Enabled = EnableComment
        .mnuFind.Enabled = EnableFind

    End With
End Function
Private Function EnableComment() As Boolean
    With frmMainIDE
        If Not .ActiveForm Is Nothing Then
            If TypeName(.ActiveForm) = "frmScriptPage" Then
                EnableComment = True
            Else
                EnableComment = False
            End If
        Else
            EnableComment = False
        End If
    End With
End Function
Private Function EnableFind() As Boolean
    With frmMainIDE
        If Not .ActiveForm Is Nothing Then
            If TypeName(.ActiveForm) = "frmScriptPage" Then
                EnableFind = True
            Else
                EnableFind = False
            End If
        Else
            EnableFind = False
        End If
    End With
End Function
Private Function EnableUndo() As Boolean
    With frmMainIDE
        If Not .ActiveForm Is Nothing Then
            If TypeName(.ActiveForm) = "frmScriptPage" Then
                EnableUndo = .ActiveForm.CodeEdit1.CanUndo And (Not .ActiveForm.Locked)
            Else
                EnableUndo = False
            End If
        Else
            EnableUndo = False
        End If
    End With
End Function
Private Function EnableRedo() As Boolean
    With frmMainIDE
        If Not .ActiveForm Is Nothing Then
            If TypeName(.ActiveForm) = "frmScriptPage" Then
                EnableRedo = .ActiveForm.CodeEdit1.CanRedo And (Not .ActiveForm.Locked)
            Else
                EnableRedo = False
            End If
        Else
            EnableRedo = False
        End If
    End With
End Function

Private Function EnableCopy() As Boolean
    With frmMainIDE
        If Not .ActiveForm Is Nothing Then
            If TypeName(.ActiveForm) = "frmScriptPage" Then
                EnableCopy = (.ActiveForm.CodeEdit1.SelLength > 0)
            Else
                EnableCopy = False
            End If
        Else
            EnableCopy = False
        End If
    End With
End Function
Private Function EnableCut() As Boolean
    With frmMainIDE
        If Not .ActiveForm Is Nothing Then
            If TypeName(.ActiveForm) = "frmScriptPage" Then
                EnableCut = (.ActiveForm.CodeEdit1.SelLength > 0) And (Not .ActiveForm.Locked)
            Else
                EnableCut = False
            End If
        Else
            EnableCut = False
        End If
    End With
End Function
Private Function EnablePaste() As Boolean
    With frmMainIDE
        If Not .ActiveForm Is Nothing Then
            If TypeName(.ActiveForm) = "frmScriptPage" Then
                EnablePaste = (Len(Clipboard.GetText(vbCFText)) > 0) And (Not .ActiveForm.Locked)
            Else
                EnablePaste = False
            End If
        Else
            EnablePaste = False
        End If
    End With
End Function
Private Function EnableDelete() As Boolean
    With frmMainIDE
        If Not .ActiveForm Is Nothing Then
            If TypeName(.ActiveForm) = "frmScriptPage" Then
                EnableDelete = (.ActiveForm.CodeEdit1.SelLength > 0) And (Not .ActiveForm.Locked)
            Else
                EnableDelete = False
            End If
        Else
            EnableDelete = False
        End If
    End With
End Function
Private Function EnableSelectAll() As Boolean
    With frmMainIDE
        If Not .ActiveForm Is Nothing Then
            If TypeName(.ActiveForm) = "frmScriptPage" Then
                EnableSelectAll = (Len(.ActiveForm.CodeEdit1.Text) > 0)
            Else
                EnableSelectAll = False
            End If
        Else
            EnableSelectAll = False
        End If
    End With
End Function




Attribute 