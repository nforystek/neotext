Attribute VB_Name = "Module1"
Option Explicit

Public Sub RealDegreeAngle(ByRef Angle As Double)
    'input an angle, and ensures it is with-in
    '0.001 to 360 degrees, no neg/zero angles.
    Dim tmp As Double
    If Angle > 360 Then 'above 360
        tmp = Angle - 360
        'invalid numbers can hang it,
        'no change, so tmp<>Angle too
        Do While tmp > 360 And tmp <> Angle
            tmp = tmp - 360
        Loop
        Angle = tmp
    End If
    If Angle <= 0 Then 'zero or below
        tmp = Angle + 360
        'invalid numbers can hang it,
        'no change, so tmp<>Angle too
        Do While tmp <= 0 And tmp <> Angle
            tmp = tmp + 360
        Loop
        Angle = tmp
    End If
End Sub

Public Sub RealDegreeAngles(ByRef Angles As Point)
    '3 axis version of RealDegreeAngle(Angle)
    RealDegreeAngle Angles.X
    RealDegreeAngle Angles.Y
    RealDegreeAngle Angles.Z
End Sub

Public Sub JoinAngles(ByRef BeforeAngles As Point, ByRef AfterAngles As Point, ByRef JoinedAngles As Point)

    RealDegreeAngles BeforeAngles
    RealDegreeAngles AfterAngles
    
    If JoinedAngles Is Nothing Then Set JoinedAngles = New Point
    
    With JoinedAngles
        .X = CDec(CStr(AfterAngles.X) & "." & PaddingLeft(CStr(BeforeAngles.X), 3))
        .Y = CDec(CStr(AfterAngles.Y) & "." & PaddingLeft(CStr(BeforeAngles.Y), 3))
        .Z = CDec(CStr(AfterAngles.Z) & "." & PaddingLeft(CStr(BeforeAngles.Z), 3))
    End With

End Sub

Public Sub PartAngles(ByRef JoinedAngles As Point, ByRef BeforeAngles As Point, ByRef AfterAngles As Point)

    If BeforeAngles Is Nothing Then Set BeforeAngles = New Point
    If AfterAngles Is Nothing Then Set AfterAngles = New Point

    With BeforeAngles
        .X = CDec(PaddingRight(RemoveArg(CStr(JoinedAngles.X), "."), 3))
        .Y = CDec(PaddingRight(RemoveArg(CStr(JoinedAngles.Y), "."), 3))
        .Z = CDec(PaddingRight(RemoveArg(CStr(JoinedAngles.Z), "."), 3))
    End With
    With AfterAngles
        .X = CDec(NextArg(CStr(JoinedAngles.X), "."))
        .Y = CDec(NextArg(CStr(JoinedAngles.Y), "."))
        .Z = CDec(NextArg(CStr(JoinedAngles.Z), "."))
    End With
    
End Sub

Public Sub JoinCompoundAngles(ByRef BeforeAngles As Point, ByRef AfterAngles As Point, ByRef CompoundAngles As Point)

    RealDegreeAngles BeforeAngles
    RealDegreeAngles AfterAngles
    
    If CompoundAngles Is Nothing Then Set CompoundAngles = New Point
    
    With CompoundAngles
        .X = CDec(CStr(AfterAngles.X * 1000) & "." & RemoveArg(CStr(BeforeAngles.X / 1000), "."))
        .Y = CDec(CStr(AfterAngles.Y * 1000) & "." & RemoveArg(CStr(BeforeAngles.Y / 1000), "."))
        .Z = CDec(CStr(AfterAngles.Z * 1000) & "." & RemoveArg(CStr(BeforeAngles.Z / 1000), "."))
    End With

End Sub

Public Sub PartCompoundAngles(ByRef CompoundAngles As Point, ByRef BeforeAngles As Point, ByRef AfterAngles As Point)

    If BeforeAngles Is Nothing Then Set BeforeAngles = New Point
    If AfterAngles Is Nothing Then Set AfterAngles = New Point

    With BeforeAngles
        .X = CDbl("." & RemoveArg(CStr(CompoundAngles.X), ".")) * 1000
        .Y = CDbl("." & RemoveArg(CStr(CompoundAngles.Y), ".")) * 1000
        .Z = CDbl("." & RemoveArg(CStr(CompoundAngles.Z), ".")) * 1000
    End With
    With AfterAngles
        .X = CDbl(NextArg(CStr(CompoundAngles.X), ".")) / 1000
        .Y = CDbl(NextArg(CStr(CompoundAngles.Y), ".")) / 1000
        .Z = CDbl(NextArg(CStr(CompoundAngles.Z), ".")) / 1000
    End With

End Sub

Public Function RandomAngles() As Point
    Randomize
    Set RandomAngles = New Point
    With RandomAngles
        .X = Round(RandomPositive(0.001, 360), 0)
        .Y = Round(RandomPositive(0.001, 360), 0)
        .Z = Round(RandomPositive(0.001, 360), 0)
    End With
    RealDegreeAngles RandomAngles
End Function
Public Sub Main()
    
    Dim test1 As New Point
    Dim test2 As New Point
    Dim test3 As New Point
    Dim test4 As New Point
    
    Dim test5 As New Point
    Dim test6 As New Point
    Dim test7 As New Point
    Dim test8 As New Point

    Dim join1 As New Point
    Dim join2 As New Point

    Dim join3 As New Point
    Dim join4 As New Point
    
    Dim comp1 As New Point
    
    Debug.Print "BEGIN"
    
    'make four random 3 axis degree angles
    test1 = RandomAngles
    Debug.Print test1;
    
    test2 = RandomAngles
    Debug.Print test2;

    test3 = RandomAngles
    Debug.Print test3;
    
    test4 = RandomAngles
    Debug.Print test4
    
    'join the first pair
    JoinAngles test1, test2, join1
    Debug.Print join1;
    
    'join a second pair
    JoinAngles test3, test4, join2
    Debug.Print join2
    
    'compund the two new joined pair
    JoinCompoundAngles join1, join2, comp1
    Debug.Print comp1
        
    'with comp1 get the two compound joined abgles
    PartCompoundAngles comp1, join3, join4
    Debug.Print join3;
    Debug.Print join4
    
    'with the first compound join get the joined angles
    PartAngles join3, test5, test6
    Debug.Print test5;
    Debug.Print test6;

    'with the second compound join get the joined angles
    PartAngles join4, test7, test8
    Debug.Print test7;
    Debug.Print test8
    
    If test1 <> test5 Or test2 <> test6 Or test3 <> test7 Or test4 <> test8 Or join1 <> join3 Or join2 <> join4 Then
        Err.Raise 8, , "The angles did not successfully compound and decompound!"
    End If

    
    Debug.Print "END"
    
End Sub
