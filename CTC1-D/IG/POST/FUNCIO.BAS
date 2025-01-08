Attribute VB_Name = "Módulo2"

Public Function FMax(FArray)
    Static Lai, Laj, Uai, Uaj, Maxx
    
    Lai = LBound(FArray, 1) + 1
    Laj = LBound(FArray, 2) + 1
    Uai = UBound(FArray, 1)
    Uaj = UBound(FArray, 2)
    Maxx = FArray(Lai, Laj)
    
    For j = Laj To Uaj
        For i = Lai To Uai
            If Maxx < FArray(i, j) Then _
               Maxx = FArray(i, j)
        Next i
    Next j
    
    FMax = Maxx
End Function


Public Function FMin(FArray)
    Static Lai, Laj, Uai, Uaj, Minn
    
    Lai = LBound(FArray, 1) + 1
    Laj = LBound(FArray, 2) + 1
    Uai = UBound(FArray, 1)
    Uaj = UBound(FArray, 2)
    Minn = FArray(Lai, Laj)
    
    For j = Laj To Uaj
        For i = Lai To Uai
            If Minn > FArray(i, j) Then _
               Minn = FArray(i, j)
        Next i
    Next j
    
    FMin = Minn
End Function
Function Log10(x) As Double
        Log10 = Log(x) / Log(10)
End Function


Function Reo(x, n As Integer) As Double
Dim Om, CC, Sc, Sd As Double
    Select Case x
        Case Is <> 0
            Om = Fix(Log10(Abs(x)))
            CC = "0."
            For i = 1 To n
                CC = CC & "0"
            Next i
            Sc = x / 10 ^ Om
            Sd = Format(Sc, CC)
            Reo = Format(IIf(Abs(Sc - Sd) > 5 * 10 ^ -(n + 1), (Sd + 10 ^ -n) * 10 ^ Om, Sd * 10 ^ Om), "#.00e+##")
        Case 0
            Reo = 0
    End Select
End Function

