


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary

Private myFTPClient1
Private myFTPClient2

Friend Property Get FTPClient1() As NTAdvFTP61.Client
    Set FTPClient1 = myFTPClient1
End Property

Friend Property Set FTPClient1(ByVal newval As NTAdvFTP61.Client)
    Set myFTPClient1 = newval
End Property

Friend Property Get FTPClient2() As NTAdvFTP61.Client
    Set FTPClient2 = myFTPClient2
End Property

Friend Property Set FTPClient2(ByVal newval As NTAdvFTP61.Client)
    Set myFTPClient2 = newval
End Property

Friend Sub InitializeClients()

    Set myFTPClient1 = CreateObject("NTAdvFTP61.Client")
    Set myFTPClient2 = CreateObject("NTAdvFTP61.Client")
    
End Sub

Friend Sub TerminateClients()

    Set myFTPClient1 = Nothing
    Set myFTPClient2 = Nothing
    
End Sub

Public Property Get AppProcessId() As Long
    AppProcessId = GetCurrentProcessId
End Property
Public Property Get AppThreadId() As Long
    AppThreadId = GetCurrentThreadId
End Property
Public Property Get AppHInstance() As Long
    AppHInstance = App.hInstance
End Property
Public Property Get AppComments() As String
    AppComments = App.Comments
End Property
Public Property Get AppCompanyName() As String
    AppCompanyName = App.CompanyName
End Property
Public Property Get AppFileDescription() As String
    AppFileDescription = App.FileDescription
End Property
Public Property Get AppLegalCopyright() As String
    AppLegalCopyright = App.LegalCopyright
End Property
Public Property Get AppLegalTrademarks() As String
    AppLegalTrademarks = App.LegalTrademarks
End Property
Public Property Get AppMajor() As Integer
    AppMajor = App.Major
End Property
Public Property Get AppMinor() As Integer
    AppMinor = App.Minor
End Property
Public Property Get AppProductName() As String
    AppProductName = App.ProductName
End Property
Public Property Get AppRevision() As Integer
    AppRevision = App.Revision
End Property
Public Property Get AppTitle() As String
    AppTitle = App.Title
End Property

Public Function WindowCreate() As Long
    Dim newClient3 As New frmFTPClientGUI
    newClient3.LoadClient
    newClient3.ShowClient
    WindowCreate = newClient3.hwnd
End Function

Public Sub WindowUnload(ByVal hwnd As Long)
    Dim frm As Form
    For Each frm In Forms
        If TypeName(frm) = "frmFTPClientGUI" Then
            If frm.hwnd = hwnd Then
                Unload frm
            End If
        End If
    Next
End Sub

Public Function WindowExists(ByVal hwnd As Long) As Boolean
    Dim frm As Form
    For Each frm In Forms
        If TypeName(frm) = "frmFTPClientGUI" Then
            If frm.hwnd = hwnd Then WindowExists = True
        End If
    Next
End Function

Public Sub QuitApplication() 'we need a hop out component callie, else a crash
    SetTimer WindowInitialize(False), ObjPtr(Me), 500, AddressOf ShutDownActiveX
End Sub