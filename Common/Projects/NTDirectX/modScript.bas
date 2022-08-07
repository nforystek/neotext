Attribute VB_Name = "modScript"
Option Explicit

'#################################################################
'### These objects made public are added to ScriptControl also ###
'### are the only thing global exposed to modFactory.Execute   ###
'#################################################################

Public Globals As New Globals
Public Molecules As New NTNodes10.Collection
Public Bindings As New Bindings

