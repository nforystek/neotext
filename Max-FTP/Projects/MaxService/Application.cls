VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
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

Private myFTPClient1 As Object
Attribute myFTPClient1.VB_VarHelpID = -1
Private myFTPClient2 As Object
Attribute myFTPClient2.VB_VarHelpID = -1

Friend Property Get FTPClient1()
    Set FTPClient1 = myFTPClient1
End Property

Friend Property Set FTPClient1(ByRef newval)
    Set myFTPClient1 = newval
End Property

Friend Property Get FTPClient2()
    Set FTPClient2 = myFTPClient2
End Property

Friend Property Set FTPClient2(ByRef newval)
    Set myFTPClient2 = newval
End Property

Friend Sub InitializeClients()

    Set myFTPClient1 = New NTAdvFTP61.Client 'CreateObject("NTAdvFTP61.Client")
    Set myFTPClient2 = New NTAdvFTP61.Client ' CreateObject("NTAdvFTP61.Client")

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
