VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsISOFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strName     As String
Private strDOSName  As String
Private strJolName  As String
Private strPath     As String
Private lngSize     As Long
Private lngSortNum  As Long
Private lngLBA      As Long
Private dateCreat   As Date

Public Property Get CreationDate( _
) As Date

    CreationDate = dateCreat
End Property

Public Property Let CreationDate( _
    ByVal dtNewDate As Date _
)

    dateCreat = dtNewDate
End Property

Public Property Get SortNumber( _
) As Long

    SortNumber = lngSortNum
End Property

Public Property Let SortNumber( _
    ByVal lngNewVal As Long _
)

    lngSortNum = lngNewVal
End Property

Public Property Get LBA( _
) As Long

    LBA = lngLBA
End Property

Public Property Let LBA( _
    ByVal lngNewVal As Long _
)

    lngLBA = lngNewVal
End Property

Public Property Get JolietName( _
) As String

    JolietName = strJolName
End Property

Public Property Let JolietName( _
    ByVal strVal As String _
)

    strJolName = strVal
End Property

Public Property Get DOSName( _
) As String

    DOSName = strDOSName
End Property

Public Property Let DOSName( _
    ByVal strVal As String _
)

    strDOSName = strVal
End Property

Public Property Get name( _
) As String

    name = strName
End Property

Public Property Let name( _
    ByVal strVal As String _
)

    strName = strVal
End Property

Public Property Get LocalPath( _
) As String

    LocalPath = strPath
End Property

Public Property Let LocalPath( _
    ByVal strVal As String _
)

    strPath = strVal
End Property

Public Property Get Size( _
) As Long

    Size = lngSize
End Property

Public Property Let Size( _
    ByVal lngVal As Long _
)

    lngSize = lngVal
End Property
