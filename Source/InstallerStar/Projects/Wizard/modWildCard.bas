Attribute VB_Name = "modWildCard"
Option Explicit
'TOP DOWN

Option Compare Text


Public Function LikeCompare(ByVal text As String, ByVal wild As String) As Boolean
    LikeCompare = text Like wild
End Function
