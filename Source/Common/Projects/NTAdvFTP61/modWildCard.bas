Attribute VB_Name = "modWildCard"

#Const modWildCard = -1
Option Explicit
Option Compare Text

Option Private Module

Public Function TestWildCard(ByVal Text As String, ByVal Wild As String) As Boolean
    TestWildCard = Text Like Wild
End Function
    












