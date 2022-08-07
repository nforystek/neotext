VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RemindMe Utility"
   ClientHeight    =   7320
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   12828
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   12828
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox DefaultJScript 
      Height          =   5655
      Left            =   6405
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "frmMain.frx":6042
      Top             =   480
      Width           =   5805
   End
   Begin VB.TextBox DefaultVBScript 
      Height          =   5655
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmMain.frx":6C0D
      Top             =   585
      Width           =   5805
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

'Comment to catch libraries used


''**********************************************************************************************************
''Note: Any comment beginning with RemindMe is a functional RemindMe comment, they
''are used to add GUI functionality to procedure/functions in the operation wizard
''
''   Usage: 'RemindMe:Enum:Name(DisplayedEnum:Value, DisplayedEnum:Value, DisplayedEnum:Value, ....)
''   Usage: 'RemindMe:Function:Name(DisplayedParam:Type, DisplayedParam:Type, DisplayedParam:Type, ....)
''
''Types for function parameters can be Boolean, Numeric, String, Browse, or <Enum Name>
''The name field of functions must be the same as the actual scripted function
''Declared RemindMe enumerators can be used in as types for function parameters
''Underscores are replaced with spaces (" ") in display, double underscores are replaced with dashes (" - ")
''**********************************************************************************************************
'
''############ Enum Examples: (The following enum examples are used in distributed RemindMe functions)
''RemindMe:Enum:IconTypes(Information:64, Exclamation:48, Critical:16)
''RemindMe:Enum:FocusTypes(Normal_Focus:1, Normal_No_Focus:4, Minimized_Focus:2, Minimized_No_Focus:6, Maximized_Focus:3, Hidden:0)
'
''############ Function Examples: (The following function examples are distributed as RemindMe operations able to be scheduled)
''RemindMe:Function:Popup_Window(Message:String, Title:String, Icon:IconTypes)
'Public Function Popup_Window(aMessage, aTitle, aIcon)
'    Dim popup
'    Set popup = CreateObject("NTPopup21.Window")
'    With popup
'        .Message = aMessage
'        .Icon = aIcon
'        .Title = aTitle
'        .Visible = True
'    End With
'End Function
'
''RemindMe:Function:Popup_Window_with_Sound(Message:String, Title:String, Icon:IconTypes, SoundFile:Browse, Loop_Enabled:Boolean, Loop_Times:Numeric)
'Public Function Popup_Window_with_Sound(aMessage, aTitle, aIcon, aSoundFile, aLoopEnabled, aLoopTimes)
'    Dim popup
'    Set popup = CreateObject("NTPopup21.Window")
'    With popup
'        .Message = aMessage
'        .Icon = aIcon
'        .Title = aTitle
'        .Visible = True
'    End With
'    Dim sound
'    Set sound = CreateObject("NTSound20.Player")
'    With sound
'        .FileName = aSoundFile
'        .LoopEnabled = aLoopEnabled
'        .LoopTimes = aLoopTimes
'        .PlaySound
'    End With
'End Function
'
''RemindMe:Function:Run_Program(FileName:Browse, Params:String, Focus:FocusTypes)
'Public Function Run_Program(aFileName, aParams, aFocus)
'    Dim shell
'    Set shell = CreateObject("NTShell22.Process")
'    shell.Run aFileName, aParams, aFocus, False
'End Function
'
''RemindMe:Function:Play_Sound(SoundFile:Browse, Loop_Enabled:Boolean, Loop_Times:Numeric)
'Public Function Play_Sound(aSoundFile, aLoopEnabled, aLoopTimes)
'    Dim sound
'    Set sound = CreateObject("NTSound20.Player")
'    With sound
'        .FileName = aSoundFile
'        .LoopEnabled = aLoopEnabled
'        .LoopTimes = aLoopTimes
'        .PlaySound
'    End With
'End Function
'
''RemindMe:Function:Open_Website(WebLink:String)
'Public Function Open_Website(aWebLink)
'    Dim web
'    Set web = CreateObject("NTShell22.Internet")
'    web.OpenWebsite aWebLink
'End Function

Private Sub Form_Load()

End Sub
