Attribute VB_Name = "modMain"
Option Explicit

Option Private Module

Public Declare Function vbaObjSet Lib "msvbvm60.dll" Alias "__vbaObjSet" (dstObject As Any, ByVal srcObjPtr As Long) As Long
Public Declare Function vbaObjSetAddref Lib "msvbvm60.dll" Alias "__vbaObjSetAddref" (dstObject As Any, ByVal srcObjPtr As Long) As Long

Public Strings As Strings

Public Sub Main()
    Set Strings = New Strings
    
End Sub

