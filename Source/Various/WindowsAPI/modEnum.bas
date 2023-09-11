Attribute VB_Name = "modEnum"
Option Explicit
' ------------------------------------------------------------------------
'
'    WIN32API.TXT -- Win32 API Declarations for Visual Basic
'
'              Copyright (C) 1994-98 Microsoft Corporation
'
'  This file is required for the Visual Basic 6.0 version of the APILoader.
'  Older versions of this file will not work correctly with the version
'  6.0 APILoader.  This file is backwards compatible with previous releases
'  of the APILoader with the exception that Constants are no longer declared
'  as Global or Public in this file.
'
'  This file contains only the Const, Type,
'  and Public Declare statements for  Win32 APIs.
'
'  You have a royalty-free right to use, modify, reproduce and distribute
'  this file (and/or any modified version) in any way you find useful,
'  provided that you agree that Microsoft has no warranty, obligation or
'  liability for its contents.  Refer to the Microsoft Windows Programmer's
'  Reference for further information.
'
' ------------------------------------------------------------------------

Public Enum g_netSID_NAME_USE
   SidTypeUser = 1&
   SidTypeGroup = 2&
   SidTypeDomain = 3&
   SidTypeAlias = 4&
   SidTypeWellKnownGroup = 5&
   SidTypeDeletedAccount = 6&
   SidTypeInvalid = 7&
   SidTypeUnknown = 8&
End Enum
