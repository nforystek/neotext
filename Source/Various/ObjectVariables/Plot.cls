VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Plot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pX As Single
Private pY As Single

Public Property Get X() As Single
    X = pX
End Property
Public Property Let X(ByVal RHS As Single)
    pX = RHS
End Property

Public Property Get Y() As Single
    Y = pY
End Property
Public Property Let Y(ByVal RHS As Single)
    pY = RHS
End Property

'The example debug when running this project havily depends
'on setting the "Procedure Attributes" under the Tools menu
'for the ToString() property setting to be "Procedure ID:"
'as "(Default)," checking "Hide this member," and checking
'"User Interface Default," for classes Plot, Point and Plane.

Public Property Get ToString()
Attribute ToString.VB_UserMemId = 0
Attribute ToString.VB_MemberFlags = "240"
    ToString = "[" & Round(pX, 6) & ", " & Round(pY, 6) & "]"
End Property
Public Property Let ToString(ByVal RHS)
    Select Case TypeName(RHS)
        Case "String"
            RHS = Trim(RHS)
            If Left(RHS, 1) = "[" And Right(RHS, 1) = "]" Then
                RHS = Mid(RHS, 2, Len(RHS) - 2) & ","
            Else
                If InStr(RHS, ",") = 1 Then RHS = RHS & ","
            End If
            If CountWord(RHS, ",") = 2 Then
                pX = ParseNumerical(RHS)
                pY = ParseNumerical(RHS)
            End If
        Case "Plot"
            pX = RHS.X
            pY = RHS.Y
    End Select
End Property

