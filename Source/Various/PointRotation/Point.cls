VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Point"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements Plot

Private pZ As Double

Public Property Get X() As Double
    X = Plot.X
End Property
Public Property Let X(ByVal RHS As Double)
    Plot.X = RHS
End Property

Public Property Get Y() As Double
    Y = Plot.Y
End Property
Public Property Let Y(ByVal RHS As Double)
    Plot.Y = RHS
End Property

Public Property Get Z() As Double
    Z = pZ
End Property
Public Property Let Z(ByVal RHS As Double)
    pZ = RHS
End Property

'The example debug when running this project havily depends
'on setting the "Procedure Attributes" under the Tools menu
'for the ToString() property setting to be "Procedure ID:"
'as "(Default)," checking "Hide this member," and checking
'"User Interface Default," for classes Plot, Point and Plane.

Public Property Get ToString()
Attribute ToString.VB_UserMemId = 0
Attribute ToString.VB_MemberFlags = "240"
    ToString = "[" & Round(Plot.X, 6) & ", " & Round(Plot.Y, 6) & ", " & Round(pZ, 6) & "]"
End Property
Public Property Let ToString(ByVal RHS)
    Select Case TypeName(RHS)
        Case "String"
            RHS = Trim(RHS)
            If Left(RHS, 1) = "[" And Right(RHS, 1) = "]" Then
                RHS = Mid(RHS, 2, Len(RHS) - 2) & ","
            Else
                If CountWord(RHS, ",") = 2 Then RHS = RHS & ","
            End If
            If CountWord(RHS, ",") = 3 Then
                Plot.X = ParseNumerical(RHS)
                Plot.Y = ParseNumerical(RHS)
                pZ = ParseNumerical(RHS)
            End If
        Case "Point"
            pZ = RHS.Z
    End Select
    Select Case TypeName(RHS)
        Case "Plot", "Point"
            Plot.X = RHS.X
            Plot.Y = RHS.Y
    End Select
End Property

Private Sub Class_Initialize()
    vbaObjSetAddref Plot, ObjPtr(New Plot)
End Sub

Private Sub Class_Terminate()
    vbaObjSet Plot, ObjPtr(Nothing)
End Sub

Private Property Let Plot_ToString(ByVal RHS)
    Plot.ToString = RHS
End Property

Private Property Get Plot_ToString()
    Plot_ToString = Plot.ToString
End Property

Private Property Let Plot_X(ByVal RHS As Double)
    Plot.X = RHS
End Property

Private Property Get Plot_X() As Double
    Plot_X = Plot.X
End Property

Private Property Let Plot_Y(ByVal RHS As Double)
    Plot.Y = RHS
End Property

Private Property Get Plot_Y() As Double
    Plot_Y = Plot.Y
End Property

