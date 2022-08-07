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
Public Property Get AppMajor() As String
    AppMajor = App.Major
End Property
Public Property Get AppMinor() As String
    AppMinor = App.Minor
End Property
Public Property Get AppProductName() As String
    AppProductName = App.ProductName
End Property
Public Property Get AppRevision() As String
    AppRevision = App.Revision
End Property
Public Property Get AppTitle() As String
    AppTitle = App.Title
End Property


