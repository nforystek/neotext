VERSION 5.00
Begin VB.UserControl ftpClient 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   CanGetFocus     =   0   'False
   ClientHeight    =   312
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   312
   EditAtDesignTime=   -1  'True
   ForwardFocus    =   -1  'True
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ftpClient.ctx":0000
   ScaleHeight     =   312
   ScaleMode       =   0  'User
   ScaleWidth      =   315
End
Attribute VB_Name = "ftpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Private pLinkForm As frmFTPClientGUI
Private pLinkIndex As Long
Private pIndex As Long
Private pIsSource As Boolean
Private pListItems As NTNodes10.Collection
Private pListToItems As NTNodes10.Collection
Private pPassList As NTNodes10.Collection

Private WithEvents pClient As NTAdvFTP61.Client
Attribute pClient.VB_VarHelpID = -1

Public Event DataComplete(ByVal ProgressType As NTAdvFTP61.ProgressTypes)
Public Event DataProgress(ByVal ProgressType As NTAdvFTP61.ProgressTypes, ByVal ReceivedBytes As Double)
Public Event Error(ByVal Number As Long, ByVal Source As String, ByVal Description As String)
Public Event ItemListing(ByVal ItemName As String, ByVal ItemSize As String, ByVal ItemDate As String, ByVal ItemAccess As String)
Public Event LogMessage(ByVal MessageType As NTAdvFTP61.MessageTypes, ByVal AddedText As String)

Public Action As String
Public Overwrite As Long
Public Element As Long
Public PassCount As Integer
Public ItmIndex As Integer
Public IsFolder As Boolean
Public ItemName As String
Public ItemSize As String
Public ItemDate As String
Public DestItemSize As String
Public DestItemDate As String
    
Friend Property Get obj() As NTAdvFTP61.Client
    Set obj = pClient
End Property
Friend Property Set obj(ByRef RHS As NTAdvFTP61.Client)
    Set pClient = RHS
End Property
    
Friend Property Get PassList() As NTNodes10.Collection
    Set PassList = pPassList
End Property
Friend Property Set PassList(ByRef RHS As NTNodes10.Collection)
    Set pPassList = RHS
End Property

Friend Property Get ListItems() As NTNodes10.Collection
    Set ListItems = pListItems
End Property
Friend Property Set ListItems(ByRef RHS As NTNodes10.Collection)
    Set pListItems = RHS
End Property
    
Friend Property Get ListToItems() As NTNodes10.Collection
    Set ListToItems = pListToItems
End Property
Friend Property Set ListToItems(ByRef RHS As NTNodes10.Collection)
    Set pListToItems = RHS
End Property


Friend Property Get Index() As Long
    Index = pIndex
End Property
Friend Property Let Index(ByRef RHS As Long)
    pIndex = RHS
End Property

Friend Property Get IsSource() As Boolean
    IsSource = pIsSource
End Property

Friend Property Get LinkForm() As frmFTPClientGUI
    Set LinkForm = pLinkForm
End Property

Friend Property Get LinkIndex() As Long
    LinkIndex = pLinkIndex
End Property

Friend Property Get LinkObj() As NTAdvFTP61.Client
    If Not pLinkForm Is Nothing Then
        Set LinkObj = pLinkForm.ftpClient(pLinkIndex).obj
    End If
End Property

Public Function UnLink()
    Link Nothing, -1
End Function

Public Static Function Link(ByRef LinkToForm As frmFTPClientGUI, ByVal LinkToIndex As Long)
'    Static stack As Boolean
'
'    If Not stack Then
'        stack = True
'        pIsSource = False
'
'        If (LinkToForm Is Nothing) And (Not (LinkForm Is Nothing)) Then
'            LinkForm.ftpClient(LinkIndex).Link Nothing, Index
'        ElseIf (Not (LinkToForm Is Nothing)) And (LinkForm Is Nothing) Then
'            LinkToForm.ftpClient(LinkToIndex).Link UserControl.Parent, Index
'        End If
'
'
'        Set pLinkForm = LinkToForm
'        pLinkIndex = LinkToIndex
'        pIsSource = True
'        stack = False
'    Else
'        pIsSource = False
'    End If
    

End Function

Private Sub pClient_DataComplete(ByVal ProgressType As NTAdvFTP61.ProgressTypes)
    RaiseEvent DataComplete(ProgressType)
End Sub

Private Sub pClient_DataProgress(ByVal ProgressType As NTAdvFTP61.ProgressTypes, ByVal ReceivedBytes As Double)
    RaiseEvent DataProgress(ProgressType, ReceivedBytes)
End Sub

Private Sub pClient_Error(ByVal Number As Long, ByVal Source As String, ByVal Description As String)
    RaiseEvent Error(Number, Source, Description)
End Sub

Private Sub pClient_ItemListing(ByVal ItemName As String, ByVal ItemSize As String, ByVal ItemDate As String, ByVal ItemAccess As String)
    RaiseEvent ItemListing(ItemName, ItemSize, ItemDate, ItemAccess)
End Sub

Private Sub pClient_LogMessage(ByVal MessageType As NTAdvFTP61.MessageTypes, ByVal AddedText As String)
    RaiseEvent LogMessage(MessageType, AddedText)
End Sub

Private Sub UserControl_Terminate()
    Set pClient = Nothing
    Set pLinkForm = Nothing
    Set pListItems = Nothing
    Set pListToItems = Nothing
    Set pPassList = Nothing
End Sub
