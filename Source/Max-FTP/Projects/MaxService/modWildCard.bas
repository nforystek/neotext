Attribute VB_Name = "modWildCard"
#Const modWildCard = -1
Option Explicit
'TOP DOWN
Option Compare Text
Option Private Module

Public Function WildCardMatch(ByVal Text As String, ByVal WildCard As String)
    WildCardMatch = (Text Like WildCard) Or (Text = WildCard)
End Function
