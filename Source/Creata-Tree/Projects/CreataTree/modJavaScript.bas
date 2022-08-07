Attribute VB_Name = "modJavaScript"
#Const modJavaScript = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module

Private BaseItem As clsItem

Public Function vBool(ByVal tBool As Boolean) As String
    vBool = IIf(tBool, "true", "false")
End Function
Public Function TreeImage(ByRef nBase As clsItem, ByVal Image As String) As String
    With nBase
        If (Not .Value("UsePlusMinus")) And (Not .Value("UseTreeLines")) Then
            TreeImage = "''"
        Else
            Select Case Image
                Case "plus", "minus"
                    TreeImage = IIf(.Value("UsePlusMinus"), """Media/PlusMinus/" & nBase.Value("PlusMinusColor") & "/" & Image & ".gif""", """""")
                Case Else
                    TreeImage = IIf(.Value("UseTreeLines"), """Media/TreeLines/" & nBase.Value("TreeLineColor") & "/" & Image & ".gif""", """""")
            End Select
        End If
    End With
        
End Function
Public Function MediaImage(ByRef nItem As clsItem, ByVal Image As String) As String
    MediaImage = IIf(nItem.Value("Use" & Image), """Media/" & Replace(nItem.Value(Image), "\", "/") & """", """""")
End Function
Public Function GetFont(ByRef nItem As clsItem) As String
    If nItem.Cast("UseFont") = "d" Then
        GetFont = ",""" & BaseItem.Value("FontFamily") & """" & _
                    ",""" & BaseItem.Value("FontSize") & """" & _
                    ",""" & BaseItem.Value("FontColor") & """"
    Else
        GetFont = ",""" & IIf(nItem.Value("UseFont"), LCase(nItem.Value("FontFamily")), "") & """" & _
                    ",""" & IIf(nItem.Value("UseFont"), LCase(nItem.Value("FontSize")), "") & """" & _
                    ",""" & IIf(nItem.Value("UseFont"), nItem.Value("FontColor"), "") & """"
    End If
End Function
Public Function GetTarget(ByRef nItem As clsItem) As String
    If nItem.Cast("UseLinkTarget") = "d" Then
        GetTarget = ",""" & BaseItem.Value("LinkTarget") & """"
    Else
        GetTarget = ",""" & IIf(nItem.Value("UseLinkTarget"), nItem.Value("LinkTarget"), "") & """"
    End If

End Function
Public Function GenerateTree(ByRef nBase As clsItem) As String
    Set BaseItem = nBase
    
    Dim txt As String
    With nBase
        txt = txt & "var menuBase = new menuObject(" & .Value("Top") & _
                                                    "," & .Value("Left") & _
                                                    "," & .Value("Height") & _
                                                    "," & vBool(.Value("StaticView")) & _
                                                    "," & "'Media/202.gif'" & _
                                                    "," & vBool(.Value("StretchBullets")) & _
                                                    "," & vBool(.Value("UsePlusMinus")) & _
                                                    "," & TreeImage(nBase, "plus") & _
                                                    "," & TreeImage(nBase, "minus") & _
                                                    "," & vBool(.Value("UseTreelines")) & _
                                                    "," & TreeImage(nBase, "top") & _
                                                    "," & TreeImage(nBase, "mid") & _
                                                    "," & TreeImage(nBase, "btm") & _
                                                    "," & TreeImage(nBase, "hline") & _
                                                    "," & TreeImage(nBase, "vline") & ");" & vbCrLf
    End With
    
    Dim cnt As Long
    Dim nItem As clsItem
    For Each nItem In nBase.SubItems
        If Not nItem.IsTemplate Then
            txt = txt & GenerateItem(nItem, "menuBase.menuItems[" & cnt & "]")
            cnt = cnt + 1
        End If
    Next
    
    Set BaseItem = Nothing
    GenerateTree = txt
End Function
Private Function GenerateItem(ByRef nItem As clsItem, ByVal ScriptPrefix As String) As String
    Dim txt As String
    With nItem
        txt = txt & ScriptPrefix & " = new menuItemObject(" & _
                                                            MediaImage(nItem, "Collapsed") & _
                                                            "," & MediaImage(nItem, "Expanded") & _
                                                            "," & MediaImage(nItem, "MouseOut") & _
                                                            "," & MediaImage(nItem, "MouseOver") & _
                                                            "," & IIf(.Value("UseText"), """" & .Value("Text") & """", """""") & _
                                                            GetFont(nItem) & _
                                                            ",""" & .Value("LinkURL") & """" & _
                                                            GetTarget(nItem) & _
                                                            ",""" & .Value("LinkToolTip") & """" & _
                                                            "," & IIf((nItem.SubItems.Count > 0), vBool(.Value("Opened")), "false") & ");" & vbCrLf
    
    End With
    
    Dim cnt As Long
    Dim nSubItem As clsItem
    For Each nSubItem In nItem.SubItems
        txt = txt & GenerateItem(nSubItem, ScriptPrefix & ".subMenu[" & cnt & "]")
        cnt = cnt + 1
    Next

    GenerateItem = txt
    
End Function
