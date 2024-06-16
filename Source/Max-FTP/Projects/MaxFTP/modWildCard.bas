Attribute VB_Name = "modWildCard"





#Const modWildCard = -1
Option Explicit
'TOP DOWN
Option Compare Text
Option Private Module

Public Sub SelectWildCard(ByVal pView As Control, ByVal Items As String, ByVal WildCard As String)
    Dim lstItem
    Select Case LCase(Trim(Items))
        Case "all"
            For Each lstItem In pView.ListItems
                If Left(lstItem.Text, 1) = "/" Then
                    lstItem.Selected = (Mid(lstItem.Text, 2) Like WildCard)
                Else
                    lstItem.Selected = (lstItem.Text Like WildCard)
                End If
            Next
        Case "folders"
            For Each lstItem In pView.ListItems
                If Left(lstItem.Text, 1) = "/" Then
                    lstItem.Selected = (Mid(lstItem.Text, 2) Like WildCard)
                Else
                    lstItem.Selected = False
                End If
            Next
        Case "files"
            For Each lstItem In pView.ListItems
                If Left(lstItem.Text, 1) <> "/" Then
                    lstItem.Selected = (lstItem.Text Like WildCard)
                Else
                    lstItem.Selected = False
                End If
            Next
    End Select
End Sub
