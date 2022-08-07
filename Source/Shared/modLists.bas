#Const [True] = -1
#Const [False] = 0



Attribute VB_Name = "modLists"
#Const modLists = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module

Public Sub ColumnSortClick(ByVal ListView1 As ListView, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ListView1.SortKey <> ColumnHeader.SubItemIndex Then
        ListView1.SortKey = ColumnHeader.SubItemIndex
    Else
        If ListView1.SortOrder = 0 Then
            ListView1.SortOrder = 1
        Else
            ListView1.SortOrder = 0
        End If
    End If
End Sub

Public Function GetSelectedCount(ByRef xList As Control) As Integer
    Dim cnt As Integer
    Dim selCnt As Integer
    selCnt = 0
    Select Case TypeName(xList)
        Case "ListView"
            For cnt = 1 To xList.ListItems.Count
                If xList.ListItems(cnt).Selected Then selCnt = selCnt + 1
            Next
        Case "ListBox", "ComboBox"
            For cnt = 0 To xList.ListCount - 1
                If xList.Selected(cnt) Then selCnt = selCnt + 1
            Next
    End Select
    GetSelectedCount = selCnt
End Function

Public Function GetItemCount(ByRef xList As Control, ByVal Item As String) As Integer

    Dim cnt As Integer
    Dim found As Integer
    found = 0
    Select Case TypeName(xList)
        Case "ListView"
            cnt = 1
            Do Until cnt > xList.ListItems.Count Or found > 0
                If LCase(Trim(xList.ListItems(cnt).Text)) = LCase(Trim(Item)) Then found = found + 1
                cnt = cnt + 1
            Loop
        Case "ListBox", "ComboBox"
            cnt = 0
            Do Until cnt > xList.ListCount - 1 Or found > 0
                If LCase(Trim(xList.List(cnt))) = LCase(Trim(Item)) Then found = found + 1
                cnt = cnt + 1
            Loop
    End Select
    GetItemCount = found

End Function

Public Function IsOnList(ByRef tList As Control, ByVal Item As String) As Integer

    Dim cnt As Integer
    Dim found As Integer
    cnt = 0
    found = -1
    Do Until cnt = tList.ListCount Or found > -1
        If LCase(Trim(tList.List(cnt))) = LCase(Trim(Item)) Then found = cnt
        cnt = cnt + 1
    Loop
    IsOnList = found

End Function
Public Function IsOnListEx(ByRef tList As Control, ByVal Item As String) As Integer

    Dim cnt As Integer
    Dim found As Integer
    cnt = 0
    found = -1
    Do Until cnt = tList.ListCount Or found > -1
        If InStr(LCase(Trim(tList.List(cnt))), LCase(Trim(Item))) > 0 Or InStr(LCase(Trim(Item)), LCase(Trim(tList.List(cnt)))) > 0 Then found = cnt
        cnt = cnt + 1
    Loop
    IsOnListEx = found

End Function

Public Function IsOnListItemsEx(ByRef tList As ListView, ByVal Item As String) As Integer

    Dim cnt As Integer
    Dim found As Integer
    cnt = 0
    found = -1
    Do Until cnt = tList.ListItems.Count Or found > -1
        cnt = cnt + 1
        If InStr(tList.ListItems(cnt).Text, Item) > 0 Then found = cnt
    Loop
    IsOnListItemsEx = found

End Function

Public Function IsOnListItems(ByRef tList As ListView, ByVal Item As String) As Integer

    Dim cnt As Integer
    Dim found As Integer
    cnt = 1
    found = 0
    Do Until cnt > tList.ListItems.Count Or found > 0
        If LCase(Trim(tList.ListItems(cnt).Text)) = LCase(Trim(Item)) Then found = cnt
        cnt = cnt + 1
    Loop
    IsOnListItems = found

End Function

Public Function IsOnView(ByRef tList As TreeView, ByVal Item As String) As Integer

    Dim cnt As Integer
    Dim found As Integer
    cnt = 0
    found = -1
    Do Until cnt = tList.Nodes.Count Or found > -1
        cnt = cnt + 1
        If LCase(Trim(tList.Nodes(cnt).Text)) = LCase(Trim(Item)) Then found = cnt
    Loop
    IsOnView = found

End Function

Public Function IsPathOnItems(ByRef tList As ListView, ByVal Item As String) As Integer

    Dim cnt As Integer
    Dim found As Integer
    Dim valCheck As String
    cnt = 1
    found = 0
    Item = LCase(Trim(Item))
    Do Until cnt > tList.ListItems.Count Or found > 0
        valCheck = LCase(Trim(tList.ListItems(cnt).Text))
        If valCheck = Item Then
            found = cnt
        Else
            If Left(valCheck, 1) = "/" Or Left(Item, 1) = "\" Then
                If Mid(valCheck, 2) = Item Then
                    found = cnt
                End If
            End If
        End If
        cnt = cnt + 1
    Loop
    IsPathOnItems = found

End Function
