Attribute VB_Name = "Module1"
#Const Module1 = -1
Option Explicit

Public Function RandomPrime(Optional ByVal AllowDecimals As Boolean = True) As Double
    Dim PrimeResult As Double
    Dim Factors As New Collection
    Dim Remix As New Collection
    'ten multiples of ten offset into starting with half notes and excluding
    'what woudl be then 90 and 100 for 9 and 10, rather became 11 and 12
    Remix.Add 0.5:   Remix.Add 5
    Remix.Add 10:    Remix.Add 20
    Remix.Add 30:    Remix.Add 40
    Remix.Add 50:    Remix.Add 60
    Remix.Add 70:    Remix.Add 80
    Randomize 'reseet the RND function for
    
    Do
        
        'mixing up a new set of the increments
        'using a simple hi/lo recreation suffle
        Do Until Remix.Count = 0
            If Round(Rnd, 0) = 1 Then
                Factors.Add Remix.Item(Remix.Count)
                Remix.Remove Remix.Count
            Else
                Factors.Add Remix.Item(1)
                Remix.Remove 1
            End If
        Loop

        PrimeResult = PrimeFactor(Factors)
        
        'exclude
        If Not AllowDecimals Then
            If InStr(CStr(PrimeResult), ".") > 0 Then
                'recreate the remixing collection
                Do Until Factors.Count = 0
                    Remix.Add Factors.Item(1)
                    Factors.Remove 1
                Loop
                PrimeResult = 0 'set to zero to contnue
            End If
        End If
       
    Loop Until PrimeResult <> 0
    'return the  prime
    RandomPrime = PrimeResult
End Function

Public Function Medium(ByVal Factor As Double) As Double
    'any number starting with 5 and trailing 0s's
    'returns a 33.33 variant in mode of timed grade
    'which is factor divided by seven days in a week
    'divided by 24 hours in a day, multiplied by 8
    'hours in a work day, multipied by 14 days in
    'a 2 weeks pay period to prime mean timing
    Medium = ((((Factor / 7) / 24) * 8) * 14)
    
'?Medium(90) = 60
'?Medium(80) = 53.3333333333333
'?Medium(70) = 46.6666666666667
'?Medium(60) = 40
'?Medium(50) = 33.3333333333333
'?Medium(40) = 26.6666666666667
'?Medium(30) = 20
'?Medium(20) = 13.3333333333333
'?Medium(10) = 6.66666666666667
'?Medium(05) = 3.33333333333333
'?Medium(.5) = 0.333333333333333
End Function

Private Function PrimeFactor(ByRef Factors As Collection) As Double
    'offset paired subtraction by
    'four quarters of a whole half
    'mid coupling the net result

    Debug.Print
    PrimeFactor = ( _
                    ( _
                        Medium(Factors(10)) - Medium(Factors(1)) _
                    ) + ( _
                            Medium(Factors(3)) + Medium(Factors(4)) + _
                            Medium(Factors(5)) + Medium(Factors(6)) - _
                        ( _
                            Medium(Factors(7)) + Medium(Factors(8)) _
                        ) _
                    ) - Medium(Factors(9))) + Medium(Factors(2) _
                )
    Debug.Print "((" & Medium(Factors(10)) & "-" & Medium(Factors(1)) & ")+(" & Medium(Factors(3)) & "+" & Medium(Factors(4)) & "+" & Medium(Factors(5)) & "+" & Medium(Factors(6)) & "-( " & Medium(Factors(7)) & "+" & Medium(Factors(8)) & "))-" & Medium(Factors(9)) & ")+" & Medium(Factors(2)) & "=";
    
End Function

Public Sub Main()

    Do
        Debug.Print RandomPrime(False)
        DoEvents
    Loop Until False

End Sub
